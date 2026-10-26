# -*- coding: utf-8 -*-
"""
族谱编辑器 (Genealogy Editor)
读取 WPS/Word .doc 族谱，解析人物关系，提供查询、树状展示与完整编辑能力。

用法:
    python 族谱查看器.py               # 启动图形界面(自动读取上次打开的族谱)
    python 族谱查看器.py 某个族谱.json  # 直接打开指定族谱文件
    python 族谱查看器.py --test       # 仅解析并打印统计信息(无界面)
    python 族谱查看器.py --regen      # 强制从 .doc 重新解析并刷新 genealogy.json

通用族谱:
    软件不绑定任何家族。文件菜单可"新建族谱/打开族谱/另存为"，
    每份族谱是一个独立的 .json 文件，可以直接发给别人编辑。
    各地称呼不同，可在"编辑 -> 称呼设置"里自定义全部亲属称谓，
    称谓随族谱文件一起保存，发给谁都是同一套称呼。

数据:
    子女/兄弟关系由 father 字段自动反推, 无需手工维护子列表。
    编辑在软件内完成并保存到 .json, 也可手动编辑 json(改 gender / father / birth_order 等)。

依赖: olefile (仅解析 .doc 时需要)。使用 .json 文件可完全离线运行。
"""

import os
import re
import sys
import json
import time
import shutil
import base64
import io
import datetime

def _base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _res_path(name):
    """优先 exe 同目录; 打包环境下若不存在则回退到内嵌资源(_MEIPASS)。"""
    p = os.path.join(_base_dir(), name)
    if os.path.exists(p):
        return p
    if getattr(sys, 'frozen', False):
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass and os.path.exists(os.path.join(meipass, name)):
            return os.path.join(meipass, name)
    return p


DOC_PATH = _res_path("031国和第八代孙  贤命(5)(1).doc")
CACHE_PATH = os.path.join(_base_dir(), "genealogy.txt")
# 内置示例数据(打包资源); 正式数据以用户打开/保存的 .json 文件为准
if getattr(sys, 'frozen', False):
    _mp = getattr(sys, '_MEIPASS', _base_dir())
    DATA_PATH = os.path.join(_mp, "genealogy.json")
else:
    DATA_PATH = os.path.join(_base_dir(), "genealogy.json")
DATA_PATH_WRITABLE = os.path.join(_base_dir(), "genealogy.json")  # --regen 写这里
CONFIG_PATH = os.path.join(_base_dir(), "lastfile.txt")           # 记住上次打开的族谱
# 内置示例族谱(明朝皇帝世系, 打包资源); 正式数据以用户打开/保存的 .json 文件为准
if getattr(sys, 'frozen', False):
    BUILTIN_PATH = os.path.join(_mp, "明朝族谱.json")
else:
    BUILTIN_PATH = os.path.join(_base_dir(), "明朝族谱.json")


def get_last_file():
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            p = f.read().strip()
        if p and os.path.exists(p) and p.lower().endswith('.json'):
            return p
    except Exception:
        pass
    return None


def set_last_file(path):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            f.write(path)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 1. 从 .doc 提取文本
# ---------------------------------------------------------------------------
def decode_doc(path):
    """用 olefile 读取 WordDocument 流, 文本以 UTF-16LE 存于偏移 2048 处。"""
    try:
        import olefile
    except ImportError:
        raise RuntimeError("需要 olefile 才能读取 .doc，请运行: pip install olefile")
    ole = olefile.OleFileIO(path)
    word = ole.openstream('WordDocument').read()
    text = word[2048:].decode('utf-16-le', errors='ignore')

    out = []
    for ch in text:
        o = ord(ch)
        if o == 0x000D:
            out.append('\n')
        elif o in (0x0007, 0x000B, 0x0001, 0x0002, 0x0003, 0x0004,
                   0x0005, 0x0006, 0x0011, 0x0012, 0x0013, 0x0014, 0x0015):
            out.append('')
        elif o < 0x20:
            out.append('')
        else:
            out.append(ch)
    s = ''.join(out)

    # 截掉末尾的乱码(非中文明文区域): 出现连续 >=12 个非明文字符即视为乱码起点
    cut = None
    run = 0
    for i, ch in enumerate(s):
        if not _allowed(ch):
            run += 1
            if run >= 12:
                cut = i - run + 1
                break
        else:
            run = 0
    if cut is not None:
        s = s[:cut]
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s


def _allowed(ch):
    o = ord(ch)
    if 0x20 <= o <= 0x7E:
        return True
    if 0x3000 <= o <= 0x303F:   # CJK 符号和标点
        return True
    if 0xFF00 <= o <= 0xFFEF:   # 全角/半角形
        return True
    if 0x4E00 <= o <= 0x9FFF:   # CJK 统一汉字
        return True
    if 0x3400 <= o <= 0x4DBF:   # CJK 扩展A
        return True
    if 0x2000 <= o <= 0x206F:   # 通用标点
        return True
    if o in (0x3007,):           # 〇
        return True
    return False


def load_text(force_doc=False):
    """读取族谱文本。

    默认优先读取已解析排版好的 TXT 底本(genealogy.txt), 便于人工核对与修改;
    若底本不存在(或 force_doc=True), 则从 .doc 完整解析并生成 TXT 底本。
    """
    if not force_doc and os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, encoding='utf-8') as f:
            return f.read()
    if os.path.exists(DOC_PATH):
        try:
            text = decode_doc(DOC_PATH)
            if not force_doc:
                with open(CACHE_PATH, 'w', encoding='utf-8') as f:
                    f.write(text)
            return text
        except Exception as e:
            print("读取 .doc 失败:", e)
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, encoding='utf-8') as f:
            return f.read()
    raise RuntimeError("找不到数据源(.doc 或 genealogy.txt)")


def regenerate_cache():
    """强制从 .doc 重新解析并覆盖 genealogy.txt 底本。"""
    if not os.path.exists(DOC_PATH):
        raise RuntimeError("找不到 .doc 源文件，无法重新生成")
    text = decode_doc(DOC_PATH)
    with open(CACHE_PATH, 'w', encoding='utf-8') as f:
        f.write(text)
    return text


# ---------------------------------------------------------------------------
# 2. 解析人物关系
# ---------------------------------------------------------------------------
_CN_DIGIT = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
             '七': 7, '八': 8, '九': 9, '零': 0}


def cn_to_int(s):
    """把中文数字(支持 十/百)转成整数。"""
    if s == '十':
        return 10
    total, cur = 0, 0
    for ch in s:
        if ch in _CN_DIGIT:
            cur = _CN_DIGIT[ch]
        elif ch == '十':
            total += (cur if cur else 1) * 10
            cur = 0
        elif ch == '百':
            total += cur * 100
            cur = 0
        elif ch == '千':
            total += cur * 1000
            cur = 0
    return total + cur


def clean_name(n):
    n = n.replace(' ', '')
    n = re.sub(r'[（(][^）)]*[)）]', '', n)   # 去掉括号内的异体字注释
    return n


GEN_RE = re.compile(r'^(?:第)?([一二三四五六七八九十百零]+)\s*世\s*([一-龥]{1,2}行)?')
HEAD_RE = re.compile(r'^([一-龥（）()]{2,5})\s')
# 非贪婪; 姓名首字不能是 生/配/字/别/墓 等动词或 子/女/孙/序称,
# 避免把 "生三子" 误认作父亲名
_NAME_EXCL = r'(?![生配字别墓子名女生孙继嗣长次三四五六七八九十百零])'
FATHER_RE = re.compile(_NAME_EXCL + r'([一-龥]{2,4}?)(?:之|继|嗣|长|次|三|四|五|六|七|八|九|十)?(子|女|孙)')
# 赘配: 入赘夫(随妻家, 与妻同世); 配: 普通妻
RU_RE = re.compile(r'赘配\s*([一-龥（）、·]{1,10})')
SPOUSE_RE = re.compile(r'(?<!赘)配\s*([一-龥（）、·]{1,10})')
ZI_RE = re.compile(r'字([一-龥]{1,3})')
ALIAS_RE = re.compile(r'别名([一-龥]{1,4})')
BURIAL_RE = re.compile(r'[坟墓][在]?([^。，；]+)')
KIN_MARK = re.compile(r'(之子|之女|之孙|[一二三四五六七八九十]+子|'
                      r'[一二三四五六七八九十]+女|继子|嗣子|继女|赘|配|字|别名)')


class Person:
    def __init__(self, name, clean, gen=None, gen_label=''):
        self.name = name
        self.clean = clean
        self.gen = gen
        self.gen_label = gen_label
        self.father = None
        self.father_raw = ''
        self.spouse = ''
        self.zi = ''          # 字
        self.aliases = []     # 别名/字 (clean)
        self.burial = ''
        self.extra_children = []   # 兼容保留(已不使用, 见 _kids)
        self._kids = []             # 结构化子女: (name, birth_order, adopted, gender)
        self.birth_order = None     # 排行(1=长,2=次,3=三...)
        self.birth_year = None      # 出生年(取自出生日期前4位, 用于同辈比大小)
        self.birth_date = ''        # 出生日期(自由格式: 1956-03-21 / 1956年三月初五)
        self.deceased = False       # 是否已去世
        self.age_at_death = None    # 享年(岁), 未知则为 None(不显示)
        self.death_year = None      # 去世年份, 用于自动计算享年
        self.adopted = False        # 是否养子/养女
        self.gender = ''            # '男'/'女', 原文明示时填写
        self.ru_zhui = False        # 是否为入赘夫(赘婿)
        self.ru_spouse = None       # 入赘到的配偶(妻方)人物对象
        self.married_in = False     # 是否嫁入(女嫁入夫族, 本宗不载其父母)
        self.spouse_ref = None      # 标准婚姻配偶引用(双向, 非入赘; 嫁入女依此连通夫家)
        self.mother = None          # 母亲(主要用于赘配所生子女回溯本宗)
        self.notes = ''
        self.photo = ''             # 照片(base64 JPEG, 随族谱文件保存)
        self.voices = []            # 语音信息: [{'name','ext','data'(base64)}]
        self._children = None

    @property
    def children(self):
        if self._children is None:
            self._children = [p for p in PERSONS if p.father is self]
            # 养子女排最后, 其余按排行(长子在前), 同排行按姓名
            self._children.sort(key=lambda p: (1 if p.adopted else 0,
                                                p.birth_order if p.birth_order else 999,
                                                p.name))
        return self._children

    def siblings(self):
        if self.father is None:
            return []
        return [p for p in self.father.children if p is not self]


PERSONS = []
CURRENT_TITLE = ''    # 当前族谱名称(如 "温氏族谱"), 随 json 保存
REV_INFO = {          # 族谱修订信息(手工填写), 随 json 保存
    'address': '',    # 族谱地址(祖籍/编修地等基本信息)
    'revisers': '',   # 修订人员(多人用顿号/逗号分隔)
    'date': '',       # 修订日期
    'notes': '',      # 修订说明
}


# ---------------------------------------------------------------------------
# 2.5 称谓设置(各地称呼不同, 可自定义; 随族谱文件保存)
# ---------------------------------------------------------------------------
DEFAULT_TERMS = {
    'self': '自己',
    # 配偶
    'husband': '丈夫',
    'wife': '妻子',
    # 直系长辈
    'father': '父亲',
    'mother': '母亲',
    'grandfather': '祖父（爷爷）',
    'grandmother': '祖母（奶奶）',
    'up_3': '曾祖父',
    'up_4': '高祖父',
    'up_5': '天祖',
    'up_default': '{n}世祖',
    'up_3_female': '曾祖母',
    'up_4_female': '高祖母',
    'up_5_female': '天祖母',
    'up_default_female': '{n}世祖母',
    # 同辈
    'elder_brother': '哥哥',
    'younger_brother': '弟弟',
    'elder_sister': '姐姐',
    'younger_sister': '妹妹',
    'brothers': '兄弟',            # 长幼未知同辈男(堂兄弟)
    'sisters': '姐妹',             # 长幼未知同辈女(堂姐妹)
    'tang_elder_brother': '哥',    # 堂+X, 如 堂兄/堂哥
    'tang_younger_brother': '弟',
    'tang_elder_sister': '姐',
    'tang_younger_sister': '妹',
    # 晚辈
    'son': '儿子',
    'daughter': '女儿',
    'grandson': '孙子',
    'down_3': '曾孙',
    'down_4': '玄孙',
    'down_5': '来孙',
    'down_default': '{n}世孙',
    # 旁系长辈
    'tang_prefix': '堂',           # 堂/表/族 前缀
    'aunt': '姑姑',                # 亲姑
    'tang_aunt': '姑',             # 堂X姑
    'uncle_elder': '大伯',
    'uncle_young': '小叔',
    'uncle': '叔叔',
    'tang_uncle_elder': '伯',      # 堂伯
    'tang_uncle_young': '叔',      # 堂叔
    # 旁系晚辈
    'nephew': '侄子',
    'niece': '侄女',
    'grandnephew': '侄孙',
    # 姻亲称谓(嫁入成员: 依其丈夫一侧推算)
    'daughter_in_law': '儿媳',
    'elder_brother_wife': '嫂子',
    'younger_brother_wife': '弟媳',
    'wife_of_uncle_elder': '伯母',
    'wife_of_uncle_young': '婶婶',
    'wife_of_nephew': '侄媳',
    'wife_of_grandson': '孙媳',
}

KIN_TERMS = dict(DEFAULT_TERMS)


def T(key):
    return KIN_TERMS.get(key) or DEFAULT_TERMS.get(key, key)


TERM_GROUPS = [
    ('配偶', [('husband', '丈夫'), ('wife', '妻子')]),
    ('直系长辈', [('father', '父亲'), ('mother', '母亲'), ('grandfather', '祖父'),
               ('grandmother', '祖母'),
               ('up_3', '曾祖父'), ('up_3_female', '曾祖母'),
               ('up_4', '高祖父'), ('up_4_female', '高祖母'),
               ('up_5', '天祖'), ('up_5_female', '天祖母'),
               ('up_default', '更远祖先({n}占位)'),
               ('up_default_female', '更远女祖({n}占位)')]),
    ('同辈', [('elder_brother', '哥哥'), ('younger_brother', '弟弟'),
            ('elder_sister', '姐姐'), ('younger_sister', '妹妹'),
            ('tang_elder_brother', '堂兄/堂哥用字'), ('tang_younger_brother', '堂弟用字'),
            ('tang_elder_sister', '堂姐用字'), ('tang_younger_sister', '堂妹用字'),
            ('brothers', '同辈男统称(堂X)'), ('sisters', '同辈女统称(堂X)')]),
    ('晚辈', [('son', '儿子'), ('daughter', '女儿'), ('grandson', '孙子'),
            ('down_3', '曾孙'), ('down_4', '玄孙'), ('down_5', '来孙'),
            ('down_default', '更远后裔({n}占位)')]),
    ('旁系长辈', [('tang_prefix', '堂/表前缀'), ('aunt', '亲姑'), ('tang_aunt', '堂姑用字'),
               ('uncle_elder', '伯父(父长兄)'), ('uncle_young', '幼叔(父幼弟)'),
               ('uncle', '叔叔'), ('tang_uncle_elder', '堂伯用字'),
               ('tang_uncle_young', '堂叔用字')]),
    ('旁系晚辈/其他', [('nephew', '侄子'), ('niece', '侄女'), ('grandnephew', '侄孙'),
                 ('self', '自己')]),
    ('姻亲称谓', [('daughter_in_law', '儿媳'), ('elder_brother_wife', '嫂子'),
              ('younger_brother_wife', '弟媳'), ('wife_of_uncle_elder', '伯母'),
              ('wife_of_uncle_young', '婶婶'), ('wife_of_nephew', '侄媳'),
              ('wife_of_grandson', '孙媳')]),
]


# ---------------------------------------------------------------------------
# 2.6 文本解析(仅 --regen 从 .doc 重建时使用)
# ---------------------------------------------------------------------------
def _is_person_head(p):
    m = HEAD_RE.match(p)
    if not m:
        return False
    return bool(KIN_MARK.search(p))


def _split_entries(line):
    pieces = [x.strip() for x in line.split('。')]
    entries, cur = [], None
    for p in pieces:
        if not p:
            continue
        if _is_person_head(p):
            if cur is not None:
                entries.append(cur)
            cur = p
        else:
            cur = (cur + '。' + p) if cur is not None else p
    if cur is not None:
        entries.append(cur)
    return entries


_CN_ORDER = {'长': 1, '次': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
             '七': 7, '八': 8, '九': 9, '十': 10}
ORDER_WORD = {1: '长', 2: '次', 3: '三', 4: '四', 5: '五', 6: '六',
              7: '七', 8: '八', 9: '九', 10: '十'}


def _clean_child(s):
    s = s.strip('。， 、；;')
    s = re.sub(r'[（(][^）)]*[)）]', '', s)
    return s.strip('。， 、；;')


def _remove_spans(text, spans):
    """剔除已匹配的区间, 返回剩余文本(避免兜底正则误吞'兆护之子'等)"""
    if not spans:
        return text
    spans = sorted(spans)
    out, last = [], 0
    for s, e in spans:
        out.append(text[last:s])
        last = e
    out.append(text[last:])
    return ''.join(out)


def _strip_order(part):
    """去掉数量前缀'二：'与排行前缀(长/次/二/三.../十), 返回 (name, order)"""
    part = part.strip()
    part = re.sub(r'^二[：:]', '', part)                 # 数量前缀, 如 女二：
    m = re.match(r'^([长次二三四五六七八九十]+)', part)   # 排行: 长/次/二/三...
    order = _CN_ORDER.get(m.group(1)) if m else None
    if m:
        part = part[m.end():]
    return _clean_child(part), order


def _add_kids(res, body, gender):
    for part in re.split(r'[、，,；;]', body):
        nm, order = _strip_order(part)
        if nm:
            res.append((nm, order, False, gender))


def _extract_children(entry):
    """从一条目提取子女: 返回 [(name, birth_order, adopted, gender), ...]"""
    res = []
    blocks = []
    _stop = r'(?=子[：:]|女[：:]|养[男女][：:]?|生[一二三四五六七八九十]*[子男女][：:]?|$)'
    for mm in re.finditer(r'生[一二三四五六七八九十]*子[：:]\s*([^。墓]+?)' + _stop, entry):
        _add_kids(res, mm.group(1).strip(), '男')
        blocks.append(mm.span())
    for mm in re.finditer(r'生[一二三四五六七八九十]*女[：:]?\s*([^。墓]+?)' + _stop, entry):
        _add_kids(res, mm.group(1).strip(), '女')
        blocks.append(mm.span())
    rest = _remove_spans(entry, blocks)
    for mm in re.finditer(r'女二[：:]\s*([^。墓]+?)' + _stop, rest):
        for part in re.split(r'[、，,；;]', mm.group(1).strip()):
            nm, order = _strip_order(part)
            if nm:
                res.append((nm, order, False, '女'))
    for mm in re.finditer(r'女[：:]\s*([^。，；墓]+?)' + _stop, rest):
        for part in re.split(r'[、，,；;]', mm.group(1).strip()):
            nm, order = _strip_order(part)
            if nm:
                res.append((nm, order, False, '女'))
    for mm in re.finditer(r'养([男女])[：:]\s*([^。，；墓]+?)' + _stop, rest):
        for part in re.split(r'[、，,；;]', mm.group(2).strip()):
            nm, _ = _strip_order(part)
            if nm:
                res.append((nm, None, True, mm.group(1)))
    for mm in re.finditer(r'子[：:]\s*([^。，；墓]+?)' + _stop, rest):
        for part in re.split(r'[、，,；;]', mm.group(1).strip()):
            nm, order = _strip_order(part)
            if nm:
                res.append((nm, order, False, '男'))
    # 去重: 同名(去姓/括号后)保留有排行者, 养优先
    seen = {}
    for name, order, adopted, gender in res:
        if not name:
            continue
        if name in seen:
            o, a, g = seen[name]
            seen[name] = (order if order is not None else o, a or adopted, gender or g)
        else:
            seen[name] = (order, adopted, gender)
    return [(n,) + seen[n] for n in seen]


def _child_prefix(c):
    """子女相对父亲的前缀: 长子/次子/养女..."""
    if c.adopted:
        return '养' + ('女' if c.gender == '女' else '子')
    sex_word = '子' if c.gender == '男' else ('女' if c.gender == '女' else '子')
    if c.birth_order:
        return ORDER_WORD.get(c.birth_order, '') + sex_word
    return ''


def parse_text(text):
    global PERSONS
    PERSONS = []
    by_clean, by_alias = {}, {}

    def get_or_create(name, clean=None, gen=None, gen_label=''):
        if clean is None:
            clean = clean_name(name)
        if clean in by_clean:
            p = by_clean[clean]
        elif clean in by_alias:
            p = by_alias[clean]
        else:
            p = Person(name, clean, gen, gen_label)
            PERSONS.append(p)
            by_clean[clean] = p
        if gen is not None and p.gen is None:
            p.gen = gen
        if gen_label and not p.gen_label:
            p.gen_label = gen_label
        return p

    cur_gen, cur_label = None, ''
    entries_meta = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        m = GEN_RE.match(line)
        if m:
            cur_gen = cn_to_int(m.group(1))
            cur_label = (m.group(2) or '')[:1]
            continue
        if not KIN_MARK.search(line):
            continue
        for entry in _split_entries(line):
            hm = HEAD_RE.match(entry)
            if not hm:
                continue
            name = hm.group(1)
            clean = clean_name(name)
            p = get_or_create(name, clean, cur_gen, cur_label)
            p.notes = entry

            fm = FATHER_RE.search(entry)
            if fm:
                p.father_raw = clean_name(fm.group(1))
                mk = fm.group(2)
                if mk in ('子', '孙') and p.gender != '女':
                    p.gender = '男'
                elif mk == '女':
                    p.gender = '女'
                rank_m = re.search(r'([长次二三四五六七八九十]+)[子女孙]', fm.group(0))
                if rank_m and p.birth_order is None:
                    p.birth_order = _CN_ORDER.get(rank_m.group(1))

            sm = SPOUSE_RE.search(entry)
            if sm:
                p.spouse = clean_name(sm.group(1))
            rm = RU_RE.search(entry)
            if rm:
                sname = clean_name(rm.group(1))
                x = get_or_create(sname, gen=cur_gen, gen_label=cur_label)
                x.ru_zhui = True
                x.spouse = p.name
                p.ru_spouse = x
                p.spouse = sname
                short = sname[1:] if len(sname) > 2 else sname
                by_alias.setdefault(short, x)
                by_alias.setdefault(sname, x)

            zm = ZI_RE.search(entry)
            if zm:
                p.zi = zm.group(1)
                p.aliases.append(clean_name(zm.group(1)))

            am = ALIAS_RE.search(entry)
            if am:
                p.aliases.append(clean_name(am.group(1)))

            bm = BURIAL_RE.search(entry)
            if bm:
                p.burial = bm.group(1).strip('，。 ')

            p._kids = _extract_children(entry)
            entries_meta.append((p, entry))

    for p in PERSONS:
        for a in p.aliases:
            by_alias.setdefault(a, p)

    for p, _ in entries_meta:
        father_target = p.ru_spouse if p.ru_spouse else p
        for cname, order, adopted, gender in p._kids:
            c = get_or_create(cname, gen=(p.gen + 1 if p.gen else None))
            if c.father is None:
                c.father = father_target
            if p.gender == '女' and c.mother is None:
                c.mother = p
            if order is not None:
                c.birth_order = order
            if adopted:
                c.adopted = True
            if gender:
                c.gender = gender

    for p in PERSONS:
        if p.father_raw:
            fc = p.father_raw
            if fc in by_clean:
                p.father = by_clean[fc]
            elif fc in by_alias:
                p.father = by_alias[fc]
            else:
                stub = Person(fc, fc)
                PERSONS.append(stub)
                by_clean[fc] = stub
                by_alias[fc] = stub
                p.father = stub

    name_to_person = {p.clean: p for p in PERSONS}
    for p in PERSONS:
        for cname in p.extra_children:
            cc = clean_name(cname)
            child = name_to_person.get(cc)
            if child is not None and child.father is None:
                child.father = p

    # 赘婿未记生父时, 默认把父亲填为妻子的父亲(随妻族, 连通关系网);
    # 已记载自己生父的不动
    for x in PERSONS:
        if x.ru_zhui and x.father is None:
            wife = getattr(x, 'ru_spouse', None)
            if wife is not None and wife.father is not None:
                x.father = wife.father

    # 世代推断: 已知某人世代, 向上推导未知世数的祖先(父-1, 祖父-2, ...)
    for p in PERSONS:
        if p.gen is None:
            continue
        g = p.gen
        cur = p.father
        while cur is not None and cur.gen is None:
            g -= 1
            cur.gen = g
            cur = cur.father

    roots = [p for p in PERSONS if p.father is None]
    return PERSONS, roots


# ---------------------------------------------------------------------------
def _calendar_popup(anchor, var, pal):
    """轻量内置日历选择器(无第三方依赖): 在 anchor 控件下方弹出月历,
    选择日期后以 YYYY-MM-DD 写入 var。仍可手动输入农历等自由格式。
    pal: {'BG','PANEL','TEXT','SUBTEXT','PRIMARY'} 配色字典。"""
    import tkinter as tk
    from tkinter import ttk
    BG = pal['BG']
    PANEL = pal['PANEL']
    TEXT = pal['TEXT']
    SUBTEXT = pal['SUBTEXT']
    PRIMARY = pal['PRIMARY']
    top = anchor.winfo_toplevel()
    today = datetime.date.today()
    y, m = today.year, today.month
    mm = re.match(r'^(\d{4})-(\d{1,2})', var.get().strip())
    if mm and 1 <= int(mm.group(2)) <= 12:
        y, m = int(mm.group(1)), int(mm.group(2))
    state = {'y': y, 'm': m}

    pop = tk.Toplevel(top)
    pop.title('选择日期')
    pop.transient(top)
    pop.resizable(False, False)
    pop.configure(bg=BG)
    pop.grab_set()
    pop.withdraw()      # 先隐藏, 定位好再显示, 避免在左上角闪一下

    def _leave(ev=None):
        pop.destroy()

    pop.bind('<Escape>', _leave)
    pop.bind('<FocusOut>', _leave)

    hd = tk.Frame(pop, bg=BG)
    hd.pack(fill='x', padx=6, pady=(6, 0))

    def _nav(dy, dm):
        """dy: 年增减, dm: 月增减。"""
        y2, m2 = state['y'] + dy, state['m'] + dm
        while m2 < 1:
            y2 -= 1
            m2 += 12
        while m2 > 12:
            y2 += 1
            m2 -= 12
        state['y'], state['m'] = y2, m2
        _draw()

    def _nav_btn(txt, cmd):
        return tk.Button(hd, text=txt, width=2, relief='flat', bg=PANEL,
                         fg=TEXT, activebackground='#EAEAEA', command=cmd)

    _nav_btn('◀◀', lambda: _nav(-1, 0)).pack(side='left')      # 上一年
    _nav_btn('◀', lambda: _nav(0, -1)).pack(side='left')       # 上一月
    lbl_ym = tk.Label(hd, bg=BG, fg=TEXT,
                      font=('Microsoft YaHei', 10, 'bold'))
    lbl_ym.pack(side='left', expand=True)
    _nav_btn('▶', lambda: _nav(0, 1)).pack(side='right')       # 下一月
    _nav_btn('▶▶', lambda: _nav(1, 0)).pack(side='right')      # 下一年

    # 年/月直接选择: 祖先年代一翻就到, 不用一格一格点
    yr = tk.Frame(pop, bg=BG)
    yr.pack(fill='x', padx=6, pady=(2, 0))
    tk.Label(yr, text='年份', bg=BG, fg=SUBTEXT).pack(side='left')
    y_var = tk.StringVar(value=str(state['y']))
    m_str = tk.StringVar(value=str(state['m']))

    def _on_y(*_):
        try:
            v = int(y_var.get())
        except ValueError:
            return
        if 1500 <= v <= 2100 and v != state['y']:
            state['y'] = v
            _draw()

    def _on_m(*_):
        try:
            v = int(m_str.get())
        except ValueError:
            return
        if 1 <= v <= 12 and v != state['m']:
            state['m'] = v
            _draw()

    sp = tk.Spinbox(yr, from_=1500, to=2100, textvariable=y_var, width=6,
                    justify='center', relief='flat', bg=PANEL,
                    buttonbackground=PANEL, command=_on_y)
    sp.pack(side='left', padx=(2, 10))
    sp.bind('<Return>', _on_y)
    sp.bind('<FocusOut>', _on_y)
    tk.Label(yr, text='月份', bg=BG, fg=SUBTEXT).pack(side='left')
    cb_m = ttk.Combobox(yr, textvariable=m_str, width=3, state='readonly',
                        values=[str(i) for i in range(1, 13)])
    cb_m.bind('<<ComboboxSelected>>', _on_m)
    cb_m.pack(side='left', padx=(2, 0))

    grid = tk.Frame(pop, bg=BG)
    grid.pack(padx=6, pady=4)
    pop._grid = grid
    pop._lbl = lbl_ym

    def _pick(y2, m2, d2):
        var.set(f'{y2:04d}-{m2:02d}-{d2:02d}')
        pop.destroy()

    def _draw():
        y2, m2 = state['y'], state['m']
        pop._lbl.config(text=f'{y2}年{m2}月')
        y_var.set(str(y2))             # 同步年份/月份直接选择控件
        if m_str.get() != str(m2):
            m_str.set(str(m2))
        for w in pop._grid.winfo_children():
            w.destroy()
        for i, wd in enumerate(['一', '二', '三', '四', '五', '六', '日']):
            tk.Label(pop._grid, text=wd, bg=BG, fg=SUBTEXT, width=3,
                     font=('Microsoft YaHei', 8)).grid(row=0, column=i)
        first = datetime.date(y2, m2, 1)
        ndays = (datetime.date(y2 + (m2 == 12), m2 % 12 + 1, 1) -
                 datetime.timedelta(days=1)).day
        off = first.weekday()          # 周一=0
        for d in range(ndays):
            r_, c_ = divmod(off + d, 7)
            is_today = (y2 == today.year and m2 == today.month and d == today.day)
            b = tk.Label(pop._grid, text=str(d + 1), width=3,
                         font=('Microsoft YaHei', 9),
                         bg=PRIMARY if is_today else BG,
                         fg='white' if is_today else TEXT, cursor='hand2')
            b.grid(row=r_ + 1, column=c_, padx=1, pady=1)
            b.bind('<Button-1>',
                   lambda e, dd=d + 1: _pick(y2, m2, dd))

    # 底部按钮只创建一次(放在 _draw 外, 避免翻月时累加)
    ft = tk.Frame(pop, bg=BG)
    ft.pack(fill='x', padx=6, pady=(0, 6))
    tk.Button(ft, text='今天', relief='flat', bg=PANEL,
              command=lambda: _pick(today.year, today.month,
                                    today.day)).pack(side='left')
    tk.Button(ft, text='清除', relief='flat', bg=PANEL,
              command=lambda: (var.set(''), pop.destroy())).pack(side='right')

    _draw()
    anchor.update_idletasks()
    x = anchor.winfo_rootx()
    yv = anchor.winfo_rooty() + anchor.winfo_height() + 2
    pop.update_idletasks()
    sw = anchor.winfo_screenwidth()
    sh = anchor.winfo_screenheight()
    x = min(max(0, x), sw - pop.winfo_reqwidth() - 8)
    yv = min(max(0, yv), sh - pop.winfo_reqheight() - 8)
    pop.geometry(f'+{x}+{yv}')
    pop.deiconify()     # 位置已定好, 再显示
    pop.lift()
    pop.focus_set()


# 3. 数据持久化 (通用族谱 .json 文件)
# ---------------------------------------------------------------------------
def _photo_to_b64(path, max_side=400):
    """读入图片 -> 缩到 max_side 内 -> JPEG base64(便于随 json 一起分发)。"""
    from PIL import Image
    img = Image.open(path)
    img = img.convert('RGB')
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=85)
    return base64.b64encode(buf.getvalue()).decode('ascii')


def _b64_photo(b64, box=(96, 128)):
    """base64 照片 -> 缩放到 box 内 -> ImageTk.PhotoImage; 无照片/失败返回 None。"""
    if not b64:
        return None
    try:
        from PIL import Image, ImageTk
        img = Image.open(io.BytesIO(base64.b64decode(b64)))
        img.thumbnail(box)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


# ---- 语音播放(Windows MCI, 无第三方依赖; 支持 wav/mp3/wma/m4a/aac 等) ----
_VOICE = {'n': 0, 'alias': None, 'path': None}


def _mci_send(cmd):
    """发送 MCI 命令。返回 0 表示成功。"""
    try:
        import ctypes
        return int(ctypes.windll.winmm.mciSendStringW(cmd, None, 0, 0))
    except Exception:
        return -1


def _voice_play(b64, ext):
    """播放 base64 音频。返回 None=成功, 否则为错误信息。"""
    _voice_stop()
    if sys.platform != 'win32':
        return '语音播放仅支持 Windows'
    try:
        data = base64.b64decode(b64)
    except Exception as e:
        return f'音频数据损坏: {e}'
    ext = (ext or 'mp3').lower().lstrip('.').strip()
    tmpdir = os.path.join(os.environ.get('TEMP', os.path.expanduser('~')),
                          'genealogy_voices')
    try:
        os.makedirs(tmpdir, exist_ok=True)
        path = os.path.join(tmpdir, f'gzv_{int(time.time() * 1000)}.{ext}')
        with open(path, 'wb') as f:
            f.write(data)
    except Exception as e:
        return f'无法写入临时文件: {e}'
    _VOICE['n'] += 1
    alias = f'gzv{_VOICE["n"]}'
    if _mci_send(f'open "{path}" alias {alias}') != 0:
        if _mci_send(f'open "{path}" type mpegvideo alias {alias}') != 0:
            try:
                os.remove(path)
            except OSError:
                pass
            return (f'暂不支持该音频格式(.{ext})\n'
                    '支持: wav / mp3 / wma / m4a 等\n'
                    '微信语音(silk/amr)请先转换为 mp3 后再添加')
    _VOICE['alias'] = alias
    _VOICE['path'] = path
    if _mci_send(f'play {alias}') != 0:
        _voice_stop()
        return '播放失败'
    return None


def _voice_playing():
    """是否有语音正在播放。"""
    return _VOICE['alias'] is not None


def _voice_stop():
    """停止播放并清理临时文件。"""
    if _VOICE['alias']:
        _mci_send(f'stop {_VOICE["alias"]}')
        _mci_send(f'close {_VOICE["alias"]}')
        _VOICE['alias'] = None
        if _VOICE['path']:
            try:
                os.remove(_VOICE['path'])
            except OSError:
                pass
        _VOICE['path'] = None


def _person_to_dict(p):
    return {
        "name": p.name,
        "gender": p.gender or "",
        "birth_order": p.birth_order,           # 排行: 1=长 2=次 3=三 ...
        "birth_year": p.birth_year,             # 出生年(取自出生日期前4位)
        "birth_date": p.birth_date or "",       # 出生日期(自由格式)
        "father": (getattr(p, 'father_raw', '')
                   or (p.father.name if p.father else None)) or None,
        "mother": (getattr(p, 'mother_raw', '')
                   or (p.mother.name if getattr(p, 'mother', None) else None)) or None,
        "gen": p.gen,                            # 世数(第X世)
        "gen_label": p.gen_label or "",
        "spouse": p.spouse or "",
        "zi": p.zi or "",
        "aliases": list(p.aliases or []),        # 别名/字
        "burial": p.burial or "",
        "deceased": bool(getattr(p, 'deceased', False)),  # 是否已去世
        "age_at_death": getattr(p, 'age_at_death', None),  # 享年(岁), 未知为 null
        "death_year": getattr(p, 'death_year', None),      # 去世年份, 未知为 null
        "adopted": p.adopted,                    # 是否养子/养女
        "ru_zhui": p.ru_zhui,                    # 是否入赘夫
        "married_in": bool(getattr(p, 'married_in', False)),  # 是否嫁入(女嫁入夫族)
        "notes": p.notes or "",                  # 备注
        "photo": p.photo or "",                  # 照片(base64 JPEG)
        "voices": [dict(v) for v in getattr(p, 'voices', []) or []],  # 语音信息
    }


def rebuild_from_dicts(people):
    """从 dict 列表重建 PERSONS( father/mother/入赘 引用与 _kids )。
    编辑新增/修改/删除后统一走此函数, 保证关系引用一致。"""
    global PERSONS
    PERSONS = []
    by_name = {}
    order = []
    for d in people:
        p = Person(d["name"], clean_name(d["name"]), d.get("gen"), d.get("gen_label", ""))
        p.gender = d.get("gender", "")
        p.birth_order = d.get("birth_order")
        p.birth_year = d.get("birth_year")
        p.birth_date = d.get("birth_date", "") or ""
        p.spouse = d.get("spouse", "")
        p.zi = d.get("zi", "")
        p.aliases = list(d.get("aliases", []) or [])
        p.burial = d.get("burial", "")
        p.deceased = bool(d.get("deceased", False))
        p.age_at_death = d.get("age_at_death")
        p.death_year = d.get("death_year")
        p.adopted = d.get("adopted", False)
        p.ru_zhui = d.get("ru_zhui", False)
        p.married_in = bool(d.get("married_in", False))
        p.notes = d.get("notes", "")
        p.photo = d.get("photo", "") or ""
        p.voices = [dict(v) for v in (d.get("voices") or [])
                    if isinstance(v, dict) and v.get("data")]
        p.father_raw = d.get("father") or ""
        p.mother_raw = d.get("mother") or ""
        PERSONS.append(p)
        order.append(p)
        by_name[d["name"]] = p
        by_name[p.clean] = p
        for a in p.aliases:
            by_name.setdefault(a, p)
    # 设置父亲引用(父亲不在表中则建 stub 祖先)
    for p in order:
        if p.father_raw:
            fc = p.father_raw
            if fc in by_name:
                p.father = by_name[fc]
            else:
                stub = Person(fc, clean_name(fc))
                PERSONS.append(stub)
                by_name[fc] = stub
                by_name[stub.clean] = stub
                p.father = stub
    # 设置母亲引用
    for p in order:
        mc = getattr(p, 'mother_raw', '')
        if mc and mc in by_name:
            p.mother = by_name[mc]
    # 重建 ru_spouse(入赘夫 <-> 妻)
    for p in order:
        if p.ru_zhui and p.spouse and p.spouse in by_name:
            wife = by_name[p.spouse]
            if wife is not p:
                wife.ru_spouse = p
                p.ru_spouse = wife
    # 重建 spouse_ref(标准婚姻配偶双向引用, 非入赘; 嫁入女依此连通夫家)
    for p in order:
        if not p.ru_zhui and p.spouse and p.spouse in by_name:
            spouse = by_name[p.spouse]
            if spouse is not p and not spouse.ru_zhui:
                p.spouse_ref = spouse
                spouse.spouse_ref = p
    # 赘婿未记生父时, 默认把父亲填为妻子的父亲(随妻族, 连通关系网)
    for p in order:
        if getattr(p, 'ru_zhui', False) and p.father is None and p.spouse:
            wife = by_name.get(p.spouse) or by_name.get(clean_name(p.spouse))
            if wife is not None and wife.father is not None:
                p.father = wife.father
    # 重建 _kids
    for p in PERSONS:
        if p.father is not None:
            p.father._kids.append((p.name, p.birth_order, p.adopted, p.gender))
    # 规范化 raw 名字引用(以最终解析结果为准, 保证再次序列化一致)
    for p in PERSONS:
        p.father_raw = p.father.name if p.father is not None else ''
        p.mother_raw = p.mother.name if getattr(p, 'mother', None) is not None else ''
    roots = [p for p in PERSONS if p.father is None]
    return PERSONS, roots


def apply_terms(terms):
    """把 json 里的 terms 应用到全局称谓表(只覆盖有值的键)。"""
    KIN_TERMS.clear()
    KIN_TERMS.update(DEFAULT_TERMS)
    if isinstance(terms, dict):
        for k, v in terms.items():
            if isinstance(v, str) and v.strip():
                KIN_TERMS[k] = v.strip()


def _apply_rev_info(ri):
    """把 json 里的 rev_info 应用到全局 REV_INFO(缺省置空)。"""
    REV_INFO['address'] = ''
    REV_INFO['revisers'] = ''
    REV_INFO['date'] = ''
    REV_INFO['notes'] = ''
    if isinstance(ri, dict):
        for k in ('address', 'revisers', 'date', 'notes'):
            v = ri.get(k)
            if isinstance(v, str):
                REV_INFO[k] = v.strip()


def import_json(path):
    """从指定族谱 .json 重建 PERSONS。返回 (persons, roots)。"""
    global CURRENT_TITLE
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    CURRENT_TITLE = data.get('title', '')
    apply_terms(data.get('terms'))
    _apply_rev_info(data.get('rev_info'))
    persons, roots = rebuild_from_dicts(data.get("people", []))
    return persons, roots


def save_json(path, note=None, title=None):
    """把当前 PERSONS + 称谓表保存为族谱 .json, 返回人数。"""
    global CURRENT_TITLE
    if title is not None:
        CURRENT_TITLE = title
    revs = []
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                revs = json.load(f).get('revisions', [])
        except Exception:
            revs = []
    if note:
        revs.append(note)
    data = {
        "version": 2,
        "title": CURRENT_TITLE or '',
        "help": ("通用族谱数据文件。可在软件内完成全部编辑; 也可手动修改字段: "
                 "gender(男/女)、birth_order(排行: 1=长 2=次...)、father(父亲姓名)、"
                 "gen(世数)、spouse(配偶)、photo(照片base64)、terms(称呼表, 各地称呼不同可自定义)、"
                 "rev_info(修订信息: 地址/修订人员/日期/说明; 族谱名称在 title 字段)。"
                 "children(子女)由软件按 father 自动识别, 请勿手填。"),
        "terms": dict(KIN_TERMS),
        "rev_info": dict(REV_INFO),
        "people": [_person_to_dict(p) for p in PERSONS],
        "revisions": revs,
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return len(PERSONS)


def load_persons():
    """启动加载: 上次打开的文件 > 脚本目录 genealogy.json > 内置示例(仅源码运行)。
    返回 (persons, roots, path, readonly)。readonly=True 表示内置示例(只读, 需另存)。"""
    p = get_last_file()
    if p:
        try:
            persons, roots = import_json(p)
            return persons, roots, p, False
        except Exception as e:
            print("读取上次族谱失败:", e)
    if not getattr(sys, 'frozen', False) and os.path.exists(DATA_PATH):
        try:
            persons, roots = import_json(DATA_PATH)
            return persons, roots, DATA_PATH, False
        except Exception as e:
            print("读取 genealogy.json 失败:", e)
    if getattr(sys, 'frozen', False) and os.path.exists(DATA_PATH):
        try:
            persons, roots = import_json(DATA_PATH)
            return persons, roots, None, True
        except Exception as e:
            print("读取内置示例失败:", e)
    # 内置示例族谱(明朝皇帝世系): 无用户数据/genealogy.json 时, 作为只读示例展示
    if os.path.exists(BUILTIN_PATH):
        try:
            persons, roots = import_json(BUILTIN_PATH)
            return persons, roots, None, True
        except Exception as e:
            print("读取内置示例(明朝族谱)失败:", e)
    # 打包版未内置示例且无历史文件: 空白启动, 由用户打开/新建族谱
    if getattr(sys, 'frozen', False):
        return [], [], None, False
    # 兜底: 从 DOC 解析并生成到脚本目录
    persons, roots = export_json()
    return persons, roots, DATA_PATH_WRITABLE, False


def export_json(force_doc=False):
    """解析 TXT 并把结果写成 genealogy.json (--regen 用), 返回 (PERSONS, roots)。"""
    persons, roots = parse_text(load_text(force_doc=force_doc))
    existing_revisions = []
    if os.path.exists(DATA_PATH):
        try:
            with open(DATA_PATH, 'r', encoding='utf-8') as f:
                existing_revisions = json.load(f).get('revisions', [])
        except Exception:
            pass
    global CURRENT_TITLE
    CURRENT_TITLE = '温氏族谱-圣增公世系'
    data = {
        "version": 2,
        "title": CURRENT_TITLE,
        "help": ("通用族谱数据文件。可在软件内完成全部编辑; 也可手动修改字段: "
                 "gender(男/女)、birth_order(排行)、father(父亲姓名)、gen(世数)、"
                 "spouse(配偶)、terms(称呼表)。children 由软件按 father 自动识别。"),
        "terms": dict(KIN_TERMS),
        "people": [_person_to_dict(p) for p in persons],
        "revisions": existing_revisions,
    }
    with open(DATA_PATH_WRITABLE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return persons, roots


# ---------------------------------------------------------------------------
# 4. 文本展示
# ---------------------------------------------------------------------------
def ancestor_chain(p):
    chain, cur = [], p
    while cur is not None:
        chain.append(cur)
        cur = cur.father
    return list(reversed(chain))


def descendant_lines(p, depth=0, acc=None):
    if acc is None:
        acc = []
    pad = '    ' * depth
    tag = '└─ ' if depth else ''
    gen = f'{p.gen}世' if p.gen else '?世'
    label = f'（{p.gen_label}）' if p.gen_label else ''
    name_disp = (_child_prefix(p) + p.name) if depth > 0 else p.name
    if p.ru_zhui:
        spouse_str = f' · 赘婿(入赘{p.spouse})'
    elif p.ru_spouse:
        spouse_str = f' · 赘配{p.ru_spouse.name}(入赘夫)'
    elif getattr(p, 'married_in', False):
        spouse_str = f' · 嫁入(适{p.spouse})'
    elif p.spouse:
        spouse_str = f' · 配{p.spouse}'
    else:
        spouse_str = ''
    acc.append(f'{pad}{tag}{name_disp}{label}  [{gen}]'
               + (f' · 字{p.zi}' if p.zi else '')
               + spouse_str
               + (f' · 墓{p.burial}' if p.burial else ''))
    for c in p.children:
        descendant_lines(c, depth + 1, acc)
    return acc


def full_tree_lines(roots):
    acc = []
    for r in roots:
        descendant_lines(r, 0, acc)
        acc.append('')
    return acc


def _sib_disp_name(p, s):
    """兄弟姐妹显示名。赘婿挂到岳父名下后, 其妻子(也是岳父之女)会出现在
    兄弟姐妹列表里, 此时在该名字后加括号标注(妻子)/(丈夫)。"""
    spouse = ((p.spouse and s.name == p.spouse)
              or getattr(p, 'ru_spouse', None) is s
              or getattr(s, 'ru_spouse', None) is p)
    if not spouse:
        return s.name
    if getattr(s, 'ru_zhui', False) or s.gender == '男':
        return f'{s.name}（丈夫）'
    return f'{s.name}（妻子）'


def _child_disp_name(k):
    """子女显示名: 赘婿(挂在岳父名下的入赘夫)加(赘婿)标注。"""
    name = _child_prefix(k) + k.name
    if getattr(k, 'ru_zhui', False):
        name += '（赘婿）'
    return name


def person_detail(p):
    lines = []
    gen = f'{p.gen}世' if p.gen else '未知世'
    label = f'（{p.gen_label}）' if p.gen_label else ''
    lines.append(f'【{p.name}】{label}  {gen}')
    if p.zi:
        lines.append(f'字: {p.zi}')
    if p.aliases:
        lines.append(f'别名/又字: {"、".join(p.aliases)}')
    if p.ru_zhui:
        lines.append(f'身份: 赘婿（入赘{p.spouse}家）')
    elif p.ru_spouse:
        lines.append(f'配偶: {p.ru_spouse.name}（赘婿，入赘本家）')
    elif getattr(p, 'married_in', False):
        lines.append(f'身份: 嫁入（适{p.spouse}）')
    elif p.spouse:
        lines.append(f'配偶: {p.spouse}')
    if p.father:
        lines.append(f'父亲: {p.father.name}')
    elif p.ru_zhui:
        lines.append('父亲: (入赘，本宗未载)')
    elif getattr(p, 'married_in', False):
        lines.append('父亲: (嫁入，本宗未载)')
    else:
        lines.append('父亲: (未知/始祖)')
    sibs = p.siblings()
    if sibs:
        lines.append(f'兄弟姐妹: {"、".join(_sib_disp_name(p, s) for s in sibs)}')
    kids = p.children
    if kids:
        shown = [_child_disp_name(k) for k in kids]
        lines.append(f'子女: {"、".join(shown)}')
    if getattr(p, 'deceased', False):
        lines.append('状态: 已去世')
    if getattr(p, 'age_at_death', None):
        lines.append(f'享年: {p.age_at_death}岁')
    if p.burial:
        lines.append(f'葬地: {p.burial}')
    vs = getattr(p, 'voices', []) or []
    if vs:
        lines.append(f'语音信息: {len(vs)}条（编辑成员里可播放）')
    chain = ancestor_chain(p)
    if len(chain) > 1:
        lines.append('')
        lines.append('祖先链(上溯):')
        lines.append(' ← '.join(c.name for c in chain))
    lines.append('')
    lines.append('后裔(下延):')
    lines.extend(descendant_lines(p))
    return '\n'.join(lines)


# ---------------------------------------------------------------------------
# 5. 亲属称呼计算(称谓全部来自可自定义的 KIN_TERMS)
# ---------------------------------------------------------------------------
def ancestor_path(p):
    """返回从根到 p 的祖先链(含 p, 老祖在前)。
    - 若生父为入赘夫(外姓), 则改走母亲一侧以回溯本宗共祖;
    - 入赘夫自身在称呼上视为妻父之子(随妻族), 以便按其妻族定长幼称谓;
    - 嫁入女无本宗父时, 改走丈夫之父, 以便按夫家定长幼称谓。"""
    chain, cur = [], p
    while cur is not None:
        chain.append(cur)
        nxt = cur.father
        if (nxt is not None and getattr(nxt, 'ru_zhui', False)
                and getattr(cur, 'mother', None) is not None):
            nxt = cur.mother
        elif nxt is None and getattr(cur, 'ru_zhui', False):
            wife = getattr(cur, 'ru_spouse', None)
            if wife is not None and wife.father is not None:
                nxt = wife.father
        elif nxt is None and getattr(cur, 'married_in', False):
            husband = getattr(cur, 'spouse_ref', None)
            if husband is not None and husband.father is not None:
                nxt = husband.father
        cur = nxt
    return list(reversed(chain))


def lca(a, b):
    """最近公共祖先(无则返回 None)。"""
    pa = {id(x) for x in ancestor_path(a)}
    for x in reversed(ancestor_path(b)):
        if id(x) in pa:
            return x
    return None


def _up_term(n):
    table = {1: T('father'), 2: T('grandfather'), 3: T('up_3'),
             4: T('up_4'), 5: T('up_5')}
    if n in table:
        return table[n]
    try:
        return T('up_default').format(n=n)
    except Exception:
        return f'{n}世祖'


def _down_term(n):
    table = {1: T('son'), 2: T('grandson'), 3: T('down_3'),
             4: T('down_4'), 5: T('down_5')}
    if n in table:
        return table[n]
    try:
        return T('down_default').format(n=n)
    except Exception:
        return f'{n}世孙'


def _female_up_term(n):
    """女性长辈称谓(称呼嫁入成员的直系女祖时用)。"""
    table = {2: T('grandmother'), 3: T('up_3_female'),
             4: T('up_4_female'), 5: T('up_5_female')}
    if n in table:
        return table[n]
    try:
        return T('up_default_female').format(n=n)
    except Exception:
        return f'{n}世祖母'


def _sibling_term(a, b):
    """a 与 b 为兄弟姐妹, 返回 a 称 b。"""
    g = b.gender if b.gender else '男'
    elder = (b.birth_order or 0) < (a.birth_order or 0)
    if g == '女':
        return T('elder_sister') if elder else T('younger_sister')
    return T('elder_brother') if elder else T('younger_brother')


def _paternal_from(lca_node, node):
    """从 lca 到 node 的路径是否留在本宗(堂亲); 中途经外嫁女性则为表亲。"""
    path = ancestor_path(node)
    i = path.index(lca_node)
    for x in path[i + 1:-1]:
        if x.gender == '女' and getattr(x, 'ru_spouse', None) is not None:
            continue   # 招赘女: 子女随母姓留本宗
        if x.gender == '女':
            return False
    return True


def kinship_term(a, b):
    """返回 a 对 b 的称呼(中文亲属称谓, 称谓词来自称呼设置)。
    辈分以 gen(世) 为准: 同世=兄弟/堂表, 差一世=叔侄/父子..."""
    if a is b:
        return T('self')
    # 配偶最优先: 赘婿连入岳父家后夫妻会同父, 必须先于兄弟判断
    a_sp, b_sp = getattr(a, 'ru_spouse', None), getattr(b, 'ru_spouse', None)
    if (b_sp is a or a_sp is b
            or (a.spouse and b.clean == clean_name(a.spouse))
            or (b.spouse and a.clean == clean_name(b.spouse))):
        if getattr(b, 'ru_zhui', False) or b.gender == '男':
            return T('husband')
        return T('wife')
    # 嫁入者称呼他人: 借用丈夫的身份推算(随夫家定称谓);
    # 若配偶同为嫁入(异常数据)则跳过, 防止相互委托成死循环
    if (getattr(a, 'married_in', False)
            and getattr(a, 'spouse_ref', None) is not None
            and b is not a.spouse_ref
            and not getattr(a.spouse_ref, 'married_in', False)):
        return kinship_term(a.spouse_ref, b)
    ga, gb = a.gen, b.gen
    if ga is None or gb is None:
        return '（世信息缺失，无法判定）'
    up = ga - gb                      # b 比 a 高 up 辈(up>0: b 是 a 长辈)
    bg = b.gender if b.gender else '男'

    L = lca(a, b)
    prefix = T('tang_prefix')
    lineal = (L is not None) and (L is a or L is b)
    eff_prefix = '' if lineal else prefix

    # 嫁入成员: 他人称呼嫁入者, 通过其丈夫推算姻亲称谓
    if (getattr(b, 'married_in', False)
            and getattr(b, 'spouse_ref', None) is not None
            and a is not b.spouse_ref):
        h = b.spouse_ref
        if up == 0:
            # a 与嫁入者同辈: a 与其丈夫是兄弟/堂兄弟 -> 嫂子/弟媳
            if a.father is not None and a.father is h.father:
                a_is_elder = (a.birth_order is not None
                              and h.birth_order is not None
                              and a.birth_order < h.birth_order)
                return T('younger_brother_wife') if a_is_elder else T('elder_brother_wife')
            if a.birth_order is not None and h.birth_order is not None:
                a_is_elder = a.birth_order < h.birth_order
                return (eff_prefix
                        + (T('younger_brother_wife') if a_is_elder
                           else T('elder_brother_wife')))
            return eff_prefix + T('sisters')
        if up == 1:
            # a 比嫁入者晚一辈: 丈夫之子叫母亲, 丈夫之侄叫伯母/婶婶
            if a.father is h:
                return T('mother')
            if (a.father is not None and h.father is not None
                    and a.father.father is h.father):
                h_elder = (h.birth_order is not None
                           and a.father.birth_order is not None
                           and h.birth_order < a.father.birth_order)
                return T('wife_of_uncle_elder') if h_elder else T('wife_of_uncle_young')
            return eff_prefix + T('wife_of_uncle_young')
        if up >= 2:
            return _female_up_term(up)
        if up == -1:
            # a 比嫁入者长一辈: 丈夫之父叫儿媳, 丈夫之伯叔也叫伯母/婶婶
            if h.father is not None and a is h.father:
                return T('daughter_in_law')
            if (h.father is not None and a.father is not None
                    and a.father is h.father.father):
                a_elder = (a.birth_order is not None
                           and h.father.birth_order is not None
                           and a.birth_order < h.father.birth_order)
                return T('wife_of_uncle_elder') if a_elder else T('wife_of_uncle_young')
            return eff_prefix + T('wife_of_nephew')
        if up <= -2:
            return T('wife_of_grandson')

    # 直系: 父母(上一世且确为父/母)
    if up == 1 and a.father is b:
        return T('father')
    if up == 1 and a.mother is b:
        return T('mother')
    # 直系: 子女(下一世且确为子/女)
    if up == -1 and (b.father is a or b.mother is a):
        return T('daughter') if bg == '女' else T('son')

    if up == 0:                       # 同辈
        if a.father is not None and a.father is b.father:
            return _sibling_term(a, b)          # 同胞兄弟/姐妹
        if a.birth_year and b.birth_year and a.birth_year != b.birth_year:
            elder = b.birth_year < a.birth_year
            if bg == '女':
                return prefix + (T('tang_elder_sister') if elder
                                 else T('tang_younger_sister'))
            return prefix + (T('tang_elder_brother') if elder
                             else T('tang_younger_brother'))
        return prefix + (T('sisters') if bg == '女' else T('brothers'))

    if up == 1:                       # b 是 a 的长辈(上一世)
        if bg == '女':
            if a.father is not None and b.father is a.father:
                return T('aunt')                # 亲姑(同父)
            return prefix + T('tang_aunt')      # 堂姑
        parent = a.father
        # 亲伯叔: b 与 a 之父同父(即 b 是 a 之父的亲兄弟)
        if parent is not None and b.father is not None and b.father is parent.father:
            pb_ = b.birth_order or 0
            pa_ = parent.birth_order or 0
            if pb_ and pa_:
                if pb_ < pa_:
                    return T('uncle_elder')
                gf = parent.father
                if gf is not None:
                    sib_orders = [o for (nm, o, ad, gd) in gf._kids
                                  if o is not None and nm != parent.name and gd != '女']
                    if sib_orders and pb_ == max(sib_orders):
                        return T('uncle_young')
                return T('uncle')
            return T('uncle')
        pb_ = b.birth_order or 0
        pa_ = (parent.birth_order or 0) if parent else 0
        if pb_ and pa_:
            return prefix + (T('tang_uncle_elder') if pb_ < pa_
                             else T('tang_uncle_young'))
        return prefix + T('tang_uncle_young')   # 排行未知默认叔

    if up == 2:
        return eff_prefix + T('grandfather')
    if up > 2:
        if lineal:
            return _up_term(up)            # 直链条远祖: 曾祖父/高祖...
        return eff_prefix + _up_term(up)

    if up == -1:                      # b 是 a 的晚辈(下一世)
        if b.father is not None and a.father is not None and b.father.father is a.father:
            return T('niece') if bg == '女' else T('nephew')   # 亲侄
        return prefix + (T('niece') if bg == '女' else T('nephew'))
    if up == -2:
        if lineal:
            return T('grandson')
        return eff_prefix + T('grandnephew')
    if up < -2:
        if lineal:
            return _down_term(abs(up))     # 直链条远孙: 曾孙/玄孙...
        return eff_prefix + _down_term(abs(up))
    return '（关系较远）'


# ---------------------------------------------------------------------------
# 6. 图形界面
# ---------------------------------------------------------------------------
HELP_TEXT = """—————— 使用说明 ——————

■ 文件
  · 新建族谱: 输入族谱名称与始祖姓名, 从零建立自己的族谱
  · 打开族谱: 打开任意一份本软件保存的 .json 族谱文件
  · 保存(Ctrl+S): 保存修改到当前族谱文件
  · 另存为: 把当前族谱保存为一份新文件(发给别人的就是这份文件)
  · 导出全谱文本: 把整棵家族树导出为 .txt

■ 编辑成员
  · 左侧列表选中某人后, 点"编辑成员"修改资料(姓名/性别/排行/世数/
    父亲/配偶/别名/葬地/备注等)
  · "新增成员"添加一个人; 右键某人选"新增子女"可快速添加其子女
  · "删除成员"会连同其全部后裔一起删除(有确认提示)
  · 子女/兄弟姐妹关系由"父亲"字段自动推算, 改父亲即可改关系
  · 身份可选"入赘(随妻族)"或"嫁入": 嫁入＝女嫁入夫族, 本宗不载其父母,
    互称按丈夫一侧自动推算(嫂子/婶婶/儿媳/孙媳等)

■ 称呼设置
  · 各地称呼不同, 在"编辑 -> 称呼设置"里可自定义全部亲属称谓
  · 例: 父亲→阿爸、祖父→阿公、堂弟→堂舍弟、侄子→侄儿
  · 称呼随族谱文件保存, 发给别人的族谱自带这套称呼

■ 互称查询
  · 左侧双击人名填入甲、乙, 自动计算两人互称(按称呼设置输出)
"""


def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, simpledialog

    # ---- 启动加载数据 ----
    state = {'path': None, 'readonly': False, 'dirty': False}
    persons, roots, _p0, _ro0 = load_persons()
    state['path'] = _p0
    state['readonly'] = _ro0

    # ---- 配色(微信风格) ----
    BG = '#EDEDED'          # 应用背景(浅灰)
    PANEL = '#FFFFFF'       # 卡片/面板(白)
    PRIMARY = '#07C160'     # 微信绿(主色)
    PRIMARY_DK = '#06AD56'  # 绿(按下)
    TEXT = '#1F1F1F'        # 主文字
    SUBTEXT = '#8A8A8A'     # 次要文字
    BORDER = '#E3E3E3'      # 边框灰
    LIST_SEL = '#C9E7D4'    # 列表选中(浅绿)
    BUBBLE = '#F2F9F4'      # 结果气泡(极浅绿)
    FONT = ('Microsoft YaHei', 10)
    FONT_B = ('Microsoft YaHei', 10, 'bold')
    ORDER_ITEMS = [('', '（未知）')] + [(str(n), f'{ORDER_WORD[n]}({n})') for n in range(1, 11)]

    root = tk.Tk()
    root.geometry('920x700')
    root.minsize(820, 600)
    root.configure(bg=BG)

    ico_path = _res_path("app_icon.ico")
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass

    # ---- ttk 样式 ----
    LINE = '#D5DBE1'        # 控件描边(比 BORDER 深一档, 可见)
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TFrame', background=BG)
    style.configure('TLabel', background=BG, foreground=TEXT, font=FONT)
    style.configure('TButton', background=PANEL, foreground=TEXT,
                    borderwidth=1, relief='solid', padding=(10, 4),
                    font=FONT, bordercolor=LINE, lightcolor=LINE,
                    darkcolor=LINE)
    style.map('TButton', background=[('active', '#F2F4F6'),
                                     ('pressed', '#E8EBEE')],
              bordercolor=[('active', '#B8C0C8')])
    style.configure('Green.TButton', background=PRIMARY, foreground='white',
                    borderwidth=0, relief='flat', padding=(6, 3), width=-1,
                    font=FONT_B)
    style.map('Green.TButton', background=[('active', PRIMARY_DK),
                                           ('pressed', PRIMARY_DK)])
    # 编辑窗底部保存按钮: 度量(宽度/内边距/字体)继承 TButton 与"取消"一致,
    # 只改配色为绿色
    style.configure('SaveB.TButton', background=PRIMARY, foreground='white',
                    borderwidth=0, relief='flat')
    style.map('SaveB.TButton', background=[('active', PRIMARY_DK),
                                           ('pressed', PRIMARY_DK)])
    # 灰底标题栏上的绿色按钮（白字）
    style.configure('Header.TButton', background=PRIMARY, foreground='white',
                    borderwidth=0, relief='flat', padding=(6, 3), width=-1,
                    font=FONT_B)
    style.map('Header.TButton', background=[('active', PRIMARY_DK),
                                            ('pressed', PRIMARY_DK)])
    # 次级按钮: 白底 + 绿字 + 绿描边(描边可见)
    style.configure('Sec.TButton', background=PANEL, foreground=PRIMARY,
                    borderwidth=1, relief='solid', padding=(8, 3), width=-1,
                    font=FONT, bordercolor='#9ADFC0', lightcolor='#9ADFC0',
                    darkcolor='#9ADFC0')
    style.map('Sec.TButton', background=[('active', '#E9F9F1')],
              bordercolor=[('active', PRIMARY)],
              foreground=[('active', PRIMARY_DK)])
    # 分区卡片: 白底 + 实线可见描边 + 加粗标题
    style.configure('TLabelframe', background=PANEL, bordercolor=LINE,
                    relief='solid', borderwidth=1)
    style.configure('TLabelframe.Label', background=PANEL, foreground=TEXT,
                    font=('Microsoft YaHei', 9, 'bold'), padding=(2, 1))
    # 面板内部专用样式: 白底, 保证分区框内"铺满"不留灰底
    style.configure('Panel.TFrame', background=PANEL)
    style.configure('Panel.TLabel', background=PANEL, foreground=TEXT,
                    font=FONT)
    style.configure('Panel.TRadiobutton', background=PANEL, foreground=TEXT,
                    font=FONT, focuscolor=PANEL)
    # 置灰(禁用)态: 文字变浅灰, 底色不变(用于性别不符的身份选项)
    style.map('Panel.TRadiobutton', foreground=[('disabled', '#B8B8B8')],
              background=[('disabled', PANEL)])
    style.configure('TEntry', fieldbackground=PANEL, bordercolor=LINE,
                    lightcolor=LINE, darkcolor=LINE,
                    relief='solid', borderwidth=1, padding=(6, 4), font=FONT)
    style.map('TEntry', bordercolor=[('focus', PRIMARY)],
              lightcolor=[('focus', PRIMARY)],
              darkcolor=[('focus', PRIMARY)])
    style.configure('TCombobox', fieldbackground=PANEL, bordercolor=LINE,
                    lightcolor=LINE, darkcolor=LINE, arrowsize=12,
                    relief='solid', borderwidth=1)
    style.map('TCombobox', bordercolor=[('focus', PRIMARY)],
              lightcolor=[('focus', PRIMARY)],
              darkcolor=[('focus', PRIMARY)],
              fieldbackground=[('readonly', PANEL)])

    # ================= 状态与数据操作 =================
    def center_on_root(w):
        """把弹出窗口居中于主窗口, 定位好后才显示(配合 withdraw 防闪烁)。"""
        w.update_idletasks()
        _x = root.winfo_x() + max(0, (root.winfo_width()
                                      - w.winfo_reqwidth()) // 2)
        _y = root.winfo_y() + max(0, (root.winfo_height()
                                      - w.winfo_reqheight()) // 2)
        w.geometry(f'+{_x}+{_y}')
        w.deiconify()

    def app_title():
        if state['path']:
            name = os.path.basename(state['path'])
        elif state['readonly']:
            name = '内置示例族谱（只读，编辑后请另存为）'
        else:
            name = '未命名族谱'
        star = ' *' if state['dirty'] else ''
        ttl = CURRENT_TITLE or '族谱'
        return f'{ttl} - {name}{star}'

    def set_dirty(v=True):
        state['dirty'] = v
        root.title(app_title())

    def refresh_count():
        count_lbl.config(text=f'共 {len(persons)} 人')

    def sorted_persons():
        return sorted(persons,
                      key=lambda p: (p.gen if p.gen is not None else 999,
                                     p.birth_order if p.birth_order else 999,
                                     p.name))

    def rebuild_relations():
        """把当前 persons 序列化后重建关系引用(编辑/新增/删除后调用)。"""
        nonlocal persons, roots
        people = [_person_to_dict(p) for p in persons]
        persons, roots = rebuild_from_dicts(people)

    def do_save(save_as=False):
        """保存当前族谱。返回是否成功。"""
        path = state['path']
        need_as = save_as or (not path) or state['readonly']
        if need_as:
            init = os.path.basename(path) if path else (
                (CURRENT_TITLE or '我的族谱') + '.json')
            path = filedialog.asksaveasfilename(
                title='保存族谱', defaultextension='.json',
                initialfile=init,
                filetypes=[('族谱数据文件', '*.json'), ('所有文件', '*.*')])
            if not path:
                return False
        n = len(persons)
        try:
            save_json(path, note=f'{time.strftime("%Y-%m-%d %H:%M")} 编辑保存（共{n}人）')
        except Exception as e:
            messagebox.showerror('保存失败', str(e))
            return False
        state['path'] = path
        state['readonly'] = False
        set_last_file(path)
        set_dirty(False)
        refresh_count()
        return True

    def confirm_discard():
        """有未保存修改时询问。返回 True 表示可以继续(放弃或已保存)。"""
        if not state['dirty']:
            return True
        r = messagebox.askyesnocancel('未保存的修改',
                                      '当前族谱有未保存的修改，是否保存？')
        if r is None:
            return False
        if r:
            return do_save()
        return True

    # ---- 文件操作 ----
    def on_new():
        nonlocal persons, roots
        if not confirm_discard():
            return
        ttl = simpledialog.askstring('新建族谱', '族谱名称（如：张氏族谱）:',
                                     parent=root)
        if not ttl:
            return
        anc = simpledialog.askstring('新建族谱', '始祖姓名（第一世）:', parent=root)
        if not anc or not anc.strip():
            return
        anc = anc.strip()
        apply_terms(None)
        _apply_rev_info(None)
        global CURRENT_TITLE
        CURRENT_TITLE = ttl.strip()
        p = Person(anc, clean_name(anc), 1, '')
        persons, roots = [p], [p]
        state['path'] = None
        state['readonly'] = False
        set_dirty(True)
        qvar.set('')
        refresh_list()
        txt_set(f'已新建族谱「{CURRENT_TITLE}」，始祖：{anc}\n\n'
                '请通过左侧右键“新增子女”逐步添加成员，完成后记得保存。')
        refresh_count()

    def on_open():
        if not confirm_discard():
            return
        path = filedialog.askopenfilename(
            title='打开族谱', filetypes=[('族谱数据文件', '*.json'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            nonlocal persons, roots
            persons, roots = import_json(path)
        except Exception as e:
            messagebox.showerror('打开失败', f'无法读取该族谱文件：\n{e}')
            return
        state['path'] = path
        state['readonly'] = False
        set_last_file(path)
        set_dirty(False)
        qvar.set('')
        refresh_list()
        show_full()
        refresh_count()

    def on_open_builtin():
        """文件菜单: 打开内置示例族谱(明朝皇帝世系, 只读)。"""
        if not confirm_discard():
            return
        if not os.path.exists(BUILTIN_PATH):
            messagebox.showerror('打开失败', '未找到内置示例族谱文件。')
            return
        try:
            nonlocal persons, roots
            persons, roots = import_json(BUILTIN_PATH)
        except Exception as e:
            messagebox.showerror('打开失败', f'无法读取内置示例族谱：\n{e}')
            return
        state['path'] = None
        state['readonly'] = True
        set_dirty(False)
        qvar.set('')
        refresh_list()
        show_full()
        refresh_count()
        txt_set('当前显示的是内置示例族谱（只读）。\n\n'
                '明朝皇帝世系（1368-1644），共传十六帝，享国二百七十六年。\n\n'
                '修改后点“保存”会提示另存为一份属于你的族谱文件。\n\n'
                '—————— 全谱 ——————\n\n'
                + '\n'.join(full_tree_lines(
                    sorted(roots, key=lambda p: (p.gen if p.gen is not None else 999,
                                                 p.name)))))

    def on_save_as():
        do_save(save_as=True)

    def on_export_txt():
        lines = full_tree_lines(sorted(roots, key=lambda p: (p.gen if p.gen is not None else 999, p.name)))
        path = filedialog.asksaveasfilename(
            title='导出全谱文本', defaultextension='.txt',
            initialfile=(CURRENT_TITLE or '全谱') + '.txt',
            filetypes=[('文本文件', '*.txt')])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            messagebox.showinfo('导出成功', f'已导出到:\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    # ---- 成员编辑 ----
    def person_editor(person=None, default_father='', default_mother='',
                      default_gender='', default_name=''):
        """弹出成员编辑窗。person=None 表示新增。返回是否提交。"""
        is_new = person is None
        names_sorted = [p.name for p in sorted_persons()]
        females = [p.name for p in sorted_persons() if p.gender == '女']
        males = [p.name for p in sorted_persons() if p.gender == '男']

        win = tk.Toplevel(root)
        win.title('新增成员' if is_new else f'编辑成员 - {person.name}')
        win.configure(bg=PANEL)
        win.transient(root)
        win.grab_set()
        win.resizable(False, False)
        win.withdraw()      # 先隐藏, 居中定位好再显示, 避免左上角闪烁

        frm = ttk.Frame(win, style='Panel.TFrame')
        frm.pack(fill='both', expand=True, padx=10, pady=8)

        v_name = tk.StringVar(
            value=(default_name if is_new and default_name
                   else (person.name if person else '')))
        v_gender = tk.StringVar(
            value=(person.gender if person else '') or default_gender)
        # 排行存数字, 下拉框显示 '长子(1)' 等; 打开时转成显示值
        _order_disp = {int(k): v for k, v in ORDER_ITEMS if k}
        v_order = tk.StringVar(
            value=(_order_disp.get(person.birth_order, '')
                   if person and person.birth_order else ''))
        v_gen = tk.StringVar(value=str(person.gen) if person and person.gen else '')
        v_zi = tk.StringVar(value=person.zi if person else '')
        v_alias = tk.StringVar(
            value='、'.join(person.aliases) if person else '')
        v_spouse = tk.StringVar(value=person.spouse if person else '')
        v_father = tk.StringVar(
            value=(default_father if is_new and default_father
                   else (person.father.name if person and person.father else '')))
        v_mother = tk.StringVar(
            value=(default_mother if is_new and default_mother
                   else (person.mother.name if person and person.mother else '')))
        v_year = tk.StringVar(
            value=(person.birth_date if person and person.birth_date
                   else (str(person.birth_year) if person and person.birth_year else '')))
        v_death_year = tk.StringVar(
            value=(str(person.death_year)
                   if person and getattr(person, 'death_year', None) else ''))
        v_burial = tk.StringVar(value=person.burial if person else '')
        deceased_state = {'v': bool(getattr(person, 'deceased', False))
                          if person else False}
        v_identity = tk.StringVar(
            value=('养子/养女' if person and person.adopted
                   else '入赘(随妻族)' if person and person.ru_zhui
                   else '嫁入' if person and getattr(person, 'married_in', False)
                   else '普通成员'))
        photo_state = {'b64': (person.photo if person else '') or ''}

        # ---- 照片处理函数(先定义, 控件随后创建) ----
        def refresh_photo_preview():
            img = _b64_photo(photo_state['b64'], (108, 142))
            if img is not None:
                photo_lbl.config(image=img, text='', bg='#FFFFFF')
                photo_lbl.image = img
                x_lbl.place(relx=1.0, x=-1, y=1, anchor='ne')  # 右上角 X
            else:
                photo_lbl.config(image='', bg='#F0F0F0',
                                 text='单击\n上传照片')
                photo_lbl.image = None
                x_lbl.place_forget()

        def on_upload_photo():
            path = filedialog.askopenfilename(
                title='选择照片', parent=win,
                filetypes=[('图片文件', '*.jpg *.jpeg *.png *.bmp *.gif'),
                           ('所有文件', '*.*')])
            if not path:
                return
            try:
                photo_state['b64'] = _photo_to_b64(path)
            except ImportError:
                messagebox.showerror('错误', '照片功能需要 Pillow 库\n'
                                     '请运行: pip install Pillow', parent=win)
                return
            except Exception as e:
                messagebox.showerror('错误', f'无法读取图片:\n{e}', parent=win)
                return
            refresh_photo_preview()

        def on_clear_photo():
            photo_state['b64'] = ''
            refresh_photo_preview()

        # ================= 基本信息(含照片) =================
        sec_base = ttk.Labelframe(frm, text=' 基本信息 ')
        sec_base.grid(row=0, column=0, columnspan=2, sticky='new', pady=(0, 6))
        fg = ttk.Frame(sec_base, style='Panel.TFrame')
        fg.pack(side='left', fill='both', expand=True, padx=(8, 2), pady=4)

        def frow(r, text, widget, col=0):
            # 第二列标签更窄, 缩小姓名列与性别列之间的间距
            ttk.Label(fg, text=text, foreground=SUBTEXT,
                      width=7 if col == 0 else 4,
                      anchor='e', style='Panel.TLabel').grid(
                row=r, column=col * 2, sticky='e', padx=(0, 4), pady=3)
            widget.grid(row=r, column=col * 2 + 1, sticky='w', pady=3)

        # 出生日期: 输入框 + 内置日历选择器(📅), 仍支持农历自由格式手输
        datef = ttk.Frame(fg, style='Panel.TFrame')
        date_entry = ttk.Entry(datef, textvariable=v_year, width=11)
        date_entry.pack(side='left')
        ttk.Button(datef, text='📅', width=2,
                   command=lambda: _calendar_popup(
                       date_entry, v_year,
                       {'BG': BG, 'PANEL': PANEL, 'TEXT': TEXT,
                        'SUBTEXT': SUBTEXT, 'PRIMARY': PRIMARY})).pack(
            side='left', padx=(3, 0))

        def _update_age(*_):
            """出生日期行右侧: 在世显示现年, 已去世显示享年。"""
            mb = re.match(r'^(\d{4})', v_year.get().strip())
            if not mb:
                age_lbl.config(text='')
                return
            birth = int(mb.group(1))
            if deceased_state['v']:
                md = re.match(r'^(\d{4})$', v_death_year.get().strip())
                if md:
                    age = int(md.group(1)) - birth
                    age_lbl.config(text=f'享年{age}岁' if age >= 0 else '')
                else:
                    # 兼容旧数据: 只有享年记录没有去世年份
                    old = getattr(person, 'age_at_death', None) if person else None
                    age_lbl.config(text=f'享年{old}岁' if old else '')
            else:
                age = datetime.date.today().year - birth
                age_lbl.config(text=f'现年{age}岁' if age >= 0 else '')
        v_year.trace_add('write', _update_age)
        v_death_year.trace_add('write', _update_age)

        # 性别单选框
        gf = ttk.Frame(fg, style='Panel.TFrame')
        ttk.Radiobutton(gf, text='男', value='男', variable=v_gender,
                        style='Panel.TRadiobutton').pack(side='left')
        ttk.Radiobutton(gf, text='女', value='女', variable=v_gender,
                        style='Panel.TRadiobutton').pack(side='left', padx=(8, 0))
        if (person and person.gender) not in ('男', '女'):
            v_gender.set('')          # 未知则两个都不选

        # 年龄放在出生日期行右侧的空位(与性别同列), 不额外占宽度
        age_lbl = ttk.Label(fg, foreground=SUBTEXT, text='',
                            style='Panel.TLabel',
                            font=('Microsoft YaHei', 9))
        age_lbl.grid(row=3, column=3, sticky='w', pady=3)

        frow(0, '姓名 *', ttk.Entry(fg, textvariable=v_name, width=11))
        frow(0, '性别', gf, col=1)
        frow(1, '别名', ttk.Entry(fg, textvariable=v_alias, width=11))
        frow(1, '世数', ttk.Entry(fg, textvariable=v_gen, width=8), col=1)
        frow(2, '排行', ttk.Combobox(fg, textvariable=v_order, width=11,
                                     values=[x[1] for x in ORDER_ITEMS],
                                     state='readonly'))
        frow(2, '字', ttk.Entry(fg, textvariable=v_zi, width=8), col=1)
        frow(3, '出生日期', datef)

        # 照片(基本信息区右侧, 单击上传)
        pf = tk.Frame(sec_base, bg=PANEL)
        pf.pack(side='left', padx=(4, 8), pady=4)
        photo_holder = tk.Frame(pf, width=110, height=142, bg='#F0F0F0',
                                highlightbackground=LINE,
                                highlightthickness=1)
        photo_holder.pack_propagate(False)
        photo_holder.pack()
        photo_lbl = tk.Label(photo_holder, bg='#F0F0F0', fg=SUBTEXT,
                             text='单击\n上传照片', font=FONT, cursor='hand2')
        photo_lbl.pack(fill='both', expand=True)
        photo_lbl.bind('<Button-1>', lambda e: on_upload_photo())
        # 右上角小 X 号: 有照片时显示, 单击清除
        x_lbl = tk.Label(photo_holder, text='✕', bg='#DDDDDD', fg='#555555',
                         font=('Microsoft YaHei', 8), cursor='hand2', padx=2)
        x_lbl.bind('<Button-1>', lambda e: on_clear_photo())

        # ================= 亲属关系 =================
        sec_rel = ttk.Labelframe(frm, text=' 亲属关系 ')
        sec_rel.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        rg = ttk.Frame(sec_rel, style='Panel.TFrame')
        rg.pack(fill='x', padx=8, pady=4)

        def rrow(r, text, widget, col):
            """col: 0/1/2 —— 三个字段排同一行。"""
            ttk.Label(rg, text=text, foreground=SUBTEXT, width=5,
                      anchor='e', style='Panel.TLabel').grid(
                row=r, column=col * 2, sticky='e',
                padx=((14, 4) if col else (0, 4)), pady=3)
            widget.grid(row=r, column=col * 2 + 1, sticky='w', pady=3)

        father_cb = ttk.Combobox(rg, textvariable=v_father, width=9,
                                 values=[''] + names_sorted)
        mother_cb = ttk.Combobox(rg, textvariable=v_mother, width=9,
                                 values=[''] + females)
        rrow(0, '父亲', father_cb, 0)
        rrow(0, '母亲', mother_cb, 1)
        # 配偶: 普通成员自由输入; 嫁入时切换为丈夫下拉框(仅男成员, 选中即同步世数)
        spouse_entry = ttk.Entry(rg, textvariable=v_spouse, width=8)
        spouse_combo = ttk.Combobox(rg, textvariable=v_spouse, width=8,
                                    values=[''] + males)
        rrow(0, '配偶', spouse_entry, 2)
        spouse_combo.grid(row=0, column=5, sticky='w', pady=3)

        def _on_spouse_combo(_e=None):
            """嫁入选丈夫: 自动同步丈夫的世数。"""
            nm = v_spouse.get().strip()
            for q in persons:
                if q.name == nm and q.gen:
                    v_gen.set(str(q.gen))
                    break

        spouse_combo.bind('<<ComboboxSelected>>', _on_spouse_combo)

        def _update_identity_ui(*_):
            """嫁入: 隐藏父亲/母亲(本宗未载), 配偶切换为丈夫下拉框。"""
            jiaru = (v_identity.get() == '嫁入')
            if jiaru:
                father_cb.grid_remove()
                mother_cb.grid_remove()
                spouse_entry.grid_remove()
                spouse_combo.grid()
            else:
                father_cb.grid()
                mother_cb.grid()
                spouse_combo.grid_remove()
                spouse_entry.grid()

        v_identity.trace_add('write', _update_identity_ui)
        _update_identity_ui()

        # 第二行: 兄弟姐妹(按父亲/母亲自动计算, 括号内为长幼称呼)
        ttk.Label(rg, text='兄弟姐妹', foreground=SUBTEXT, width=5,
                  anchor='ne', style='Panel.TLabel').grid(
            row=1, column=0, sticky='e', padx=(0, 4), pady=3)
        sib_lbl = ttk.Label(rg, foreground=TEXT, style='Panel.TLabel',
                            wraplength=380, justify='left')
        sib_lbl.grid(row=1, column=1, columnspan=5, sticky='w', pady=3)

        # 第三行: 子女(按姓名自动关联, 括号内为称呼)
        ttk.Label(rg, text='子女', foreground=SUBTEXT, width=5,
                  anchor='ne', style='Panel.TLabel').grid(
            row=2, column=0, sticky='e', padx=(0, 4), pady=3)
        child_lbl = ttk.Label(rg, foreground=TEXT, style='Panel.TLabel',
                              wraplength=380, justify='left')
        child_lbl.grid(row=2, column=1, columnspan=5, sticky='w', pady=3)

        def _q_father_name(q):
            return (getattr(q, 'father_raw', '')
                    or (q.father.name if q.father else ''))

        def _q_mother_name(q):
            return (getattr(q, 'mother_raw', '')
                    or (q.mother.name if getattr(q, 'mother', None) else ''))

        def _parse_order(s):
            """排行值 -> 数字; 兼容显示值 '长子(1)' 与纯数字 '1'。"""
            s = (s or '').strip()
            mm = re.search(r'\((\d+)\)\s*$', s)
            if mm:
                return int(mm.group(1))
            if s.isdigit():
                return int(s)
            return None

        def _rel_tag(q):
            """兄弟姐妹称呼: 哥哥/姐姐/弟弟/妹妹; 判断不了返回 None。
            先比排行(数字小的在前), 排行缺失或相同再比出生年。"""
            g = getattr(q, 'gender', '')
            if g not in ('男', '女'):
                return None
            my_o = _parse_order(v_order.get())
            q_o = getattr(q, 'birth_order', None)
            older = None
            if my_o and q_o and my_o != q_o:
                older = q_o < my_o
            else:
                mm = re.match(r'^(\d{4})', v_year.get().strip())
                my_y = int(mm.group(1)) if mm else None
                q_y = getattr(q, 'birth_year', None)
                if my_y and q_y and q_y != my_y:
                    older = q_y < my_y
            if older is None:
                return None
            if g == '男':
                return '哥哥' if older else '弟弟'
            return '姐姐' if older else '妹妹'

        def _spouse_tag(q):
            """q 是否为本人配偶 -> 返回(丈夫)/(妻子)标注, 否则 None。
            编辑时配偶字段可能尚未保存, 优先用输入框实时值。"""
            sv = v_spouse.get().strip()
            is_sp = ((sv and q.name == sv)
                     or getattr(q, 'ru_spouse', None) is person
                     or (person is not None
                         and getattr(person, 'ru_spouse', None) is q))
            if not is_sp:
                return None
            if getattr(q, 'ru_zhui', False) or getattr(q, 'gender', '') == '男':
                return '丈夫'
            return '妻子'

        def _update_siblings(*_):
            fv = v_father.get().strip()
            mv = v_mother.get().strip()
            parts = []
            if fv or mv:
                for q in persons:
                    if person is not None and q is person:
                        continue
                    if (fv and _q_father_name(q) == fv) or \
                       (not fv and mv and _q_mother_name(q) == mv):
                        # 配偶标注(妻子/丈夫)优先于长幼称呼(哥哥/妹妹)
                        tag = _spouse_tag(q) or _rel_tag(q)
                        parts.append(q.name + (f'（{tag}）' if tag else ''))
            sib_lbl.config(
                text='、'.join(parts) if parts
                else ('（无）' if (fv or mv)
                      else '（填好父亲或母亲后自动关联）'))

        def _update_children(*_):
            """按姓名自动关联子女: 父亲或母亲填的是本人姓名即算。"""
            nm = v_name.get().strip()
            kids = []
            if nm:
                for q in persons:
                    if person is not None and q is person:
                        continue
                    if _q_father_name(q) != nm and _q_mother_name(q) != nm:
                        continue
                    g = getattr(q, 'gender', '')
                    if getattr(q, 'ru_zhui', False):
                        tag = '赘婿'   # 挂在岳父名下的入赘夫, 标注以免误认亲生
                    elif getattr(q, 'adopted', False):
                        tag = ('养子' if g == '男' else
                               '养女' if g == '女' else None)
                    elif g == '男':
                        tag = '儿子'
                    elif g == '女':
                        tag = '女儿'
                    else:
                        tag = None
                    kids.append(
                        ((1 if getattr(q, 'adopted', False) else 0,
                          getattr(q, 'birth_order', None) or 99,
                          q.birth_year or 9999, q.name),
                         q.name + (f'（{tag}）' if tag else '')))
            kids.sort(key=lambda x: x[0])
            child_lbl.config(
                text='、'.join(t for _, t in kids) if kids
                else '（填好姓名后自动关联子女）')

        v_father.trace_add('write', _update_siblings)
        v_mother.trace_add('write', _update_siblings)
        v_order.trace_add('write', _update_siblings)
        v_year.trace_add('write', _update_siblings)
        v_spouse.trace_add('write', _update_siblings)
        _update_siblings()
        v_name.trace_add('write', _update_children)
        _update_children()

        # ================= 身份与状态 =================
        sec_id = ttk.Labelframe(frm, text=' 身份与状态 ')
        sec_id.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        ig = ttk.Frame(sec_id, style='Panel.TFrame')
        ig.pack(fill='x', padx=8, pady=4)

        # 状态: 默认"在世"高亮, 单击切换为"已去世"
        ttk.Label(ig, text='状态', foreground=SUBTEXT, width=5,
                  anchor='e', style='Panel.TLabel').grid(
            row=0, column=0, sticky='e', padx=(0, 4), pady=3)
        deceased_btn = tk.Button(ig, font=FONT, cursor='hand2',
                                 relief='flat', borderwidth=0,
                                 padx=12, pady=1, width=6)
        deceased_btn.grid(row=0, column=1, sticky='w', pady=3)

        def _apply_deceased_style():
            if deceased_state['v']:
                deceased_btn.config(text='已去世', bg='#C9C9C9', fg=TEXT,
                                    activebackground='#B5B5B5')
            else:
                deceased_btn.config(text='在世', bg=PRIMARY, fg='white',
                                    activebackground=PRIMARY_DK)

        # 身份(同一行右列)
        ttk.Label(ig, text='身份', foreground=SUBTEXT, width=5,
                  anchor='e', style='Panel.TLabel').grid(
            row=0, column=2, sticky='e', padx=(18, 4), pady=3)
        idf = ttk.Frame(ig, style='Panel.TFrame')
        idf.grid(row=0, column=3, sticky='w', pady=3)
        # 身份四选一: 普通成员 / 养子养女 / 入赘 / 嫁入(单选框)
        zhui_rb = None
        jiaru_rb = None
        for _txt in ('普通成员', '养子/养女', '入赘(随妻族)', '嫁入'):
            _rb = ttk.Radiobutton(idf, text=_txt, value=_txt,
                                  variable=v_identity,
                                  style='Panel.TRadiobutton')
            _rb.pack(side='left')
            if _txt == '入赘(随妻族)':
                zhui_rb = _rb
            elif _txt == '嫁入':
                jiaru_rb = _rb

        # 悬停"入赘/嫁入"单选框时显示说明(替代原独立说明文字)
        tip_state = {'w': None}

        def _tip_hide(_e=None):
            if tip_state['w'] is not None:
                tip_state['w'].destroy()
                tip_state['w'] = None

        def _tip_show(w, text):
            _tip_hide()
            tw = tk.Toplevel(win)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f'+{w.winfo_rootx() + 10}'
                           f'+{w.winfo_rooty() + w.winfo_height() + 4}')
            tk.Label(tw, text=text,
                     bg='#FFFCEB', fg=TEXT, justify='left',
                     relief='solid', borderwidth=1,
                     font=('Microsoft YaHei', 9)).pack()
            tip_state['w'] = tw

        if zhui_rb is not None:
            zhui_rb.bind('<Enter>', lambda e: _tip_show(
                zhui_rb, '入赘＝男到女家（赘婿），其本人随妻族记世，\n'
                         '所生子女随母姓记世'))
            zhui_rb.bind('<Leave>', _tip_hide)
        if jiaru_rb is not None:
            jiaru_rb.bind('<Enter>', lambda e: _tip_show(
                jiaru_rb, '嫁入＝女到夫家，本宗不载其父母，\n'
                          '配偶选丈夫后，互称按丈夫一侧推算（嫂子/婶婶/儿媳等）'))
            jiaru_rb.bind('<Leave>', _tip_hide)

        # 性别限定身份: 男不可"嫁入"(男的对应"入赘"), 女不可"入赘";
        # 已选身份与新性别冲突时自动回到"普通成员"
        def _update_identity_enabled(*_):
            g = v_gender.get()
            if jiaru_rb is not None:
                jiaru_rb.state(['disabled'] if g == '男' else ['!disabled'])
            if zhui_rb is not None:
                zhui_rb.state(['disabled'] if g == '女' else ['!disabled'])
            if g == '男' and v_identity.get() == '嫁入':
                v_identity.set('普通成员')
            elif g == '女' and v_identity.get() == '入赘(随妻族)':
                v_identity.set('普通成员')

        v_gender.trace_add('write', _update_identity_enabled)
        _update_identity_enabled()

        # 去世年份 + 享年(自动计算) —— 仅"已去世"时显示, 标签与"状态"对齐
        dy_lbl = ttk.Label(ig, text='去世年份', foreground=SUBTEXT,
                           style='Panel.TLabel')
        dy_lbl.grid(row=1, column=0, sticky='e', padx=(0, 4), pady=3)
        dyf = ttk.Frame(ig, style='Panel.TFrame')
        dyf.grid(row=1, column=1, columnspan=3, sticky='w', pady=3)
        # 输入框容器宽度取"已去世"按钮实际像素宽, 两者严格对齐一致
        dy_holder = tk.Frame(dyf, bg=PANEL,
                             width=deceased_btn.winfo_reqwidth(),
                             height=deceased_btn.winfo_reqheight())
        dy_holder.pack(side='left')
        dy_holder.pack_propagate(False)
        ttk.Entry(dy_holder, textvariable=v_death_year,
                  width=10).pack(fill='both', expand=True)
        ttk.Label(dyf, text='享年', foreground=SUBTEXT,
                  style='Panel.TLabel').pack(side='left', padx=(12, 0))
        span_age_lbl = ttk.Label(dyf, text='', foreground=SUBTEXT,
                                 style='Panel.TLabel')
        span_age_lbl.pack(side='left', padx=(4, 0))

        def _update_span_age(*_):
            """填了去世年份后, 按出生年自动计算享年。"""
            md = re.match(r'^(\d{4})$', v_death_year.get().strip())
            mb = re.match(r'^(\d{4})', v_year.get().strip())
            if md and mb:
                age = int(md.group(1)) - int(mb.group(1))
                span_age_lbl.config(text=f'{age}岁' if age >= 0 else '')
            else:
                # 兼容旧数据: 只有享年记录没有去世年份
                old = getattr(person, 'age_at_death', None) if person else None
                span_age_lbl.config(text=f'{old}岁(原记录)' if old else '')
        v_death_year.trace_add('write', _update_span_age)
        v_year.trace_add('write', _update_span_age)

        # 葬地(单独一行, 输入框占满)
        burial_lbl = ttk.Label(ig, text='葬地', foreground=SUBTEXT,
                               style='Panel.TLabel')
        burial_lbl.grid(row=2, column=0, sticky='e', padx=(0, 4), pady=3)
        burial_entry = ttk.Entry(ig, textvariable=v_burial)
        burial_entry.grid(row=2, column=1, columnspan=3, sticky='ew', pady=3)
        ig.columnconfigure(3, weight=1)

        def _place_span_frame():
            """去世年份/享年/葬地连同标签, 仅当已去世时显示, 在世时全部隐藏。"""
            for w in (dy_lbl, dyf, burial_lbl, burial_entry):
                if deceased_state['v']:
                    w.grid()
                else:
                    w.grid_remove()

        def _toggle_deceased():
            deceased_state['v'] = not deceased_state['v']
            _apply_deceased_style()
            _place_span_frame()
            _update_age()

        deceased_btn.config(command=_toggle_deceased)
        _apply_deceased_style()
        _place_span_frame()

        # ================= 语音信息 =================
        sec_voice = ttk.Labelframe(frm, text=' 语音信息 ')
        sec_voice.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        voice_state = {'list': [dict(v) for v in
                                (getattr(person, 'voices', []) if person else [])]}

        vg = ttk.Frame(sec_voice, style='Panel.TFrame')
        vg.pack(fill='x', padx=8, pady=4)
        voice_lst = tk.Listbox(vg, height=3, font=FONT, bg=PANEL, fg=TEXT,
                               relief='flat', highlightthickness=1,
                               highlightbackground=LINE,
                               selectbackground=LIST_SEL,
                               selectforeground=TEXT, activestyle='none')
        voice_lst.pack(side='left', fill='both', expand=True)

        vbtns = ttk.Frame(vg, style='Panel.TFrame')
        vbtns.pack(side='left', padx=(8, 0))

        def _refresh_voices():
            voice_lst.delete(0, tk.END)
            for v in voice_state['list']:
                try:
                    kb = len(base64.b64decode(v['data'])) // 1024
                    size = f'{kb}KB'
                except Exception:
                    size = '?'
                voice_lst.insert(tk.END, f"▶ {v.get('name') or '语音'}  ({size})")
            if not voice_state['list']:
                voice_lst.insert(tk.END, '（还没有语音，点右侧"添加语音"）')

        def _sel_voice():
            sel = voice_lst.curselection()
            if not sel:
                return None
            idx = sel[0]
            if 0 <= idx < len(voice_state['list']):
                return idx
            return None

        def on_add_voice():
            path = filedialog.askopenfilename(
                title='选择语音/音频文件', parent=win,
                filetypes=[('音频文件', '*.wav *.mp3 *.wma *.m4a *.aac *.amr'),
                           ('所有文件', '*.*')])
            if not path:
                return
            try:
                with open(path, 'rb') as f:
                    raw = f.read()
            except Exception as e:
                messagebox.showerror('错误', f'无法读取文件:\n{e}', parent=win)
                return
            if len(raw) > 10 * 1024 * 1024:
                messagebox.showwarning('提示', '文件超过 10MB，太大会让族谱文件'
                                       '变得很臃肿，请先剪辑或压缩', parent=win)
                return
            ext = os.path.splitext(path)[1].lstrip('.').lower() or 'mp3'
            voice_state['list'].append({
                'name': os.path.basename(path),
                'ext': ext,
                'data': base64.b64encode(raw).decode('ascii'),
            })
            _refresh_voices()

        def on_play_voice():
            idx = _sel_voice()
            if idx is None:
                messagebox.showinfo('提示', '请先在列表中选中一条语音', parent=win)
                return
            v = voice_state['list'][idx]
            err = _voice_play(v['data'], v.get('ext'))
            if err:
                messagebox.showerror('无法播放', err, parent=win)

        def on_del_voice():
            idx = _sel_voice()
            if idx is None:
                messagebox.showinfo('提示', '请先在列表中选中要删除的语音', parent=win)
                return
            if messagebox.askyesno('删除语音',
                                   f"确定删除「{voice_state['list'][idx].get('name')}」？",
                                   parent=win):
                _voice_stop()
                voice_state['list'].pop(idx)
                _refresh_voices()

        ttk.Button(vbtns, text='添加语音', style='Sec.TButton',
                   command=on_add_voice).pack(fill='x')
        ttk.Button(vbtns, text='播放', style='Sec.TButton',
                   command=on_play_voice).pack(fill='x', pady=(4, 0))
        ttk.Button(vbtns, text='停止', style='Sec.TButton',
                   command=_voice_stop).pack(fill='x', pady=(4, 0))
        ttk.Button(vbtns, text='删除', style='Sec.TButton',
                   command=on_del_voice).pack(fill='x', pady=(4, 0))
        ttk.Label(sec_voice, foreground=SUBTEXT, style='Panel.TLabel',
                  text='双击列表项即可播放；语音随族谱文件一起保存，'
                       '微信语音请先转成 mp3/wav').pack(
            anchor='w', padx=8, pady=(0, 4))
        voice_lst.bind('<Double-Button-1>', lambda e: on_play_voice())
        _refresh_voices()

        # ================= 备注 =================
        sec_note = ttk.Labelframe(frm, text=' 备注 ')
        sec_note.grid(row=4, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        notes_txt = tk.Text(sec_note, width=40, height=3, font=FONT,
                            bg=PANEL, fg=TEXT, relief='flat',
                            borderwidth=0, highlightthickness=0)
        notes_txt.pack(fill='x', padx=8, pady=4)
        if person and person.notes:
            notes_txt.insert('1.0', person.notes)
        r = 5
        refresh_photo_preview()
        _update_age()

        # 排行显示值 <-> 数字 转换
        order_word2num = {v: int(k) for k, v in ORDER_ITEMS if k}

        def on_ok():
            name = v_name.get().strip()
            if not name:
                messagebox.showwarning('提示', '姓名不能为空', parent=win)
                return
            dup = next((x for x in persons if x.name == name), None)
            if is_new and dup:
                messagebox.showerror('错误', f'已存在同名成员「{name}」，请换个名字'
                                     '（或用别名/字区分）', parent=win)
                return
            if (not is_new) and dup is not None and dup is not person:
                messagebox.showerror('错误', f'已存在同名成员「{name}」，请换个名字',
                                     parent=win)
                return
            fv = v_father.get().strip()
            if fv and fv == name:
                messagebox.showerror('错误', '父亲不能是自己', parent=win)
                return
            gen = None
            if v_gen.get().strip():
                try:
                    gen = int(v_gen.get().strip())
                except ValueError:
                    messagebox.showerror('错误', '世数请填数字（如 1、2、25）', parent=win)
                    return
            year = None
            bdate = v_year.get().strip()
            # 出生日期: 兼容 1956 / 1956-03-21 / 1956年三月初五, 取前4位数字作出生年
            m_year = re.match(r'^(\d{4})(?!\d)', bdate)
            if m_year:
                year = int(m_year.group(1))
            death_year = None
            if v_death_year.get().strip():
                dy = v_death_year.get().strip()
                if not re.match(r'^\d{4}$', dy):
                    messagebox.showerror('错误', '去世年份请填4位数字（如 1998）',
                                         parent=win)
                    return
                death_year = int(dy)

            # 汇总字段
            ow = v_order.get()
            order_v = order_word2num.get(ow)
            if order_v is None and ow.strip().isdigit():
                order_v = int(ow.strip())      # 兼容纯数字形式
            alias_list = [a.strip() for a in re.split(r'[、,，;；/ ]', v_alias.get()) if a.strip()]
            notes = notes_txt.get('1.0', 'end').strip()

            if is_new:
                p = Person(name, clean_name(name))
                persons.append(p)
            else:
                p = person
            old_name = p.name
            p.name = name
            p.clean = clean_name(name)
            p.gender = v_gender.get()
            p.birth_order = order_v
            p.gen = gen
            p.zi = v_zi.get().strip()
            p.aliases = alias_list
            p.spouse = v_spouse.get().strip()
            p.father_raw = fv
            p.mother_raw = v_mother.get().strip()
            # 清空父亲/母亲时同时清掉旧对象引用, 交给 rebuild 按名字重建
            if not fv:
                p.father = None
            if not p.mother_raw:
                p.mother = None
            p.birth_date = bdate
            p.birth_year = year
            p.deceased = deceased_state['v']
            if not deceased_state['v']:
                p.death_year = None
                p.age_at_death = None
            else:
                p.death_year = death_year
                if death_year and year:
                    p.age_at_death = death_year - year   # 按出生/去世年自动计算
                # 未填去世年份则保留原享年记录
            p.burial = v_burial.get().strip() if deceased_state['v'] else p.burial
            p.adopted = (v_identity.get() == '养子/养女')
            p.ru_zhui = (v_identity.get() == '入赘(随妻族)')
            p.married_in = (v_identity.get() == '嫁入')
            # 配偶自动回填: 对方未填配偶时补上本人姓名, 建立双向数据链
            if p.spouse:
                sp = next((x for x in persons
                           if x.name == p.spouse and x is not p), None)
                if sp is not None and not sp.spouse:
                    sp.spouse = p.name
            p.photo = photo_state['b64']
            p.voices = voice_state['list']
            p.notes = notes
            # 若更名, 其他人的 father/mother 引用由 rebuild 按名字重建——
            # 旧名字引用会丢失, 因此先把旧引用改成新名
            if not is_new and old_name != name:
                for q in persons:
                    if q.father_raw == old_name:
                        q.father_raw = name
                    if getattr(q, 'mother_raw', '') == old_name:
                        q.mother_raw = name
            rebuild_relations()
            # 新成员: 若未填世数但有父亲, 按父亲+1 推
            if gen is None:
                me = next((x for x in persons if x.name == name), None)
                if me is not None and me.father is not None and me.father.gen:
                    me.gen = me.father.gen + 1
                    rebuild_relations()
            set_dirty(True)
            sel_name.set(name)
            refresh_list()
            refresh_count()
            win.destroy()

        def on_cancel():
            win.destroy()

        btns = ttk.Frame(frm, style='Panel.TFrame')
        btns.grid(row=r, column=0, columnspan=2, sticky='e', pady=(6, 0))
        cancel_btn = ttk.Button(btns, text='取消', command=on_cancel)
        # 保存按钮宽度取"取消"按钮实际像素宽, 两者严格一致
        save_holder = tk.Frame(btns, bg=PANEL, width=10, height=10)
        save_holder.pack(side='left', padx=6)
        save_holder.pack_propagate(False)
        save_btn = ttk.Button(save_holder, text='保存', style='SaveB.TButton',
                              command=on_ok)
        save_btn.pack(fill='both', expand=True)
        cancel_btn.pack(side='left', padx=6)
        win.bind('<Return>', lambda e: on_ok())
        win.bind('<Escape>', lambda e: on_cancel())
        win.protocol('WM_DELETE_WINDOW', on_cancel)
        win.bind('<Destroy>', lambda e: _voice_stop() if e.widget is win else None)
        # 居中定位(窗口仍隐藏): 用"取消"按钮请求宽度校正保存按钮, 定位好后统一显示
        win.update_idletasks()
        save_holder.config(width=cancel_btn.winfo_reqwidth(),
                           height=cancel_btn.winfo_reqheight())
        win.update_idletasks()
        center_on_root(win)
        win.wait_window()

    def get_selected():
        sel = lst.curselection()
        if not sel:
            return None
        idx = sel[0]
        if 0 <= idx < len(current_items):
            return current_items[idx]
        return None

    def on_add():
        person_editor(None)

    def on_edit():
        p = get_selected()
        if p is None:
            messagebox.showinfo('提示', '请先在左侧列表选中一名成员')
            return
        person_editor(p)

    def on_add_child():
        p = get_selected()
        person_editor(None, default_father=p.name if p else '')

    def on_delete():
        p = get_selected()
        if p is None:
            messagebox.showinfo('提示', '请先在左侧列表选中一名成员')
            return
        desc = []

        def collect(x):
            desc.append(x)
            for c in list(x.children):
                collect(c)
        collect(p)
        if len(desc) > 1:
            msg = (f'「{p.name}」有后裔 {len(desc) - 1} 人。\n'
                   f'删除将连同全部后裔一起删除，确定吗？')
        else:
            msg = f'确定删除成员「{p.name}」吗？'
        if not messagebox.askyesno('删除成员', msg):
            return
        remove_names = {x.name for x in desc}
        nonlocal persons, roots
        persons = [x for x in persons if x.name not in remove_names]
        # 清掉指向被删者的 mother 引用
        for x in persons:
            if getattr(x, 'mother_raw', '') in remove_names:
                x.mother_raw = ''
        rebuild_relations()
        set_dirty(True)
        refresh_list()
        refresh_count()
        txt_set(f'已删除「{p.name}」及后裔 {len(desc)} 人。\n记得保存（Ctrl+S）。')

    # ---- 称呼设置 ----
    def terms_dialog():
        win = tk.Toplevel(root)
        win.title('称呼设置（各地称呼不同，可自定义）')
        win.configure(bg=BG)
        win.transient(root)
        win.grab_set()
        win.withdraw()      # 先隐藏, 居中定位好再显示

        tip = ttk.Label(win, foreground=SUBTEXT,
                        text='修改后点“保存并应用”。称谓随族谱文件一起保存，'
                             '发给别人的族谱自带这套称呼。')
        tip.pack(anchor='w', padx=14, pady=(10, 2))

        outer = tk.Frame(win, bg=BG)
        outer.pack(fill='both', expand=True, padx=14, pady=6)
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw', tags='IF')
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        canvas.bind_all('<MouseWheel>',
                        lambda e: canvas.yview_scroll(-e.delta // 120, 'units'))
        canvas.itemconfigure('IF', width=canvas.winfo_reqwidth() - 4)

        def _unbind(_e=None):
            try:
                canvas.unbind_all('<MouseWheel>')
            except Exception:
                pass
        win.bind('<Destroy>', _unbind)

        entries = {}
        row = 0
        for gname, items in TERM_GROUPS:
            ttk.Label(inner, text=f'◆ {gname}', foreground=PRIMARY,
                      font=FONT_B).grid(row=row, column=0, columnspan=2,
                                        sticky='w', pady=(8, 2))
            row += 1
            for key, label in items:
                ttk.Label(inner, text=label, foreground=TEXT).grid(
                    row=row, column=0, sticky='e', padx=(0, 8), pady=2)
                var = tk.StringVar(value=T(key))
                e = ttk.Entry(inner, textvariable=var, width=22)
                e.grid(row=row, column=1, sticky='w', pady=2)
                entries[key] = var
                row += 1

        def on_apply():
            for key, var in entries.items():
                v = var.get().strip()
                if not v:
                    messagebox.showwarning('提示', '称谓不能为空', parent=win)
                    return
                if key in ('up_default', 'down_default') and '{n}' not in v:
                    messagebox.showwarning(
                        '提示', f'“{v}”需包含 {{n}} 占位符（代表第几世）\n'
                        '例如: 第{n}世祖', parent=win)
                    return
            for key, var in entries.items():
                KIN_TERMS[key] = var.get().strip()
            set_dirty(True)
            win.destroy()
            # 让当前互称结果立即按新称呼刷新
            if a_var.get() and b_var.get():
                compute_kin()
            messagebox.showinfo('完成', '称呼已应用并将在下次保存时写入族谱文件。\n'
                                '互称查询会立即使用新称呼。')

        def on_reset():
            for key, var in entries.items():
                var.set(DEFAULT_TERMS[key])

        btns = ttk.Frame(win)
        btns.pack(fill='x', padx=14, pady=(2, 12))
        ttk.Button(btns, text='保存并应用', style='Green.TButton',
                   command=on_apply).pack(side='left', padx=6)
        ttk.Button(btns, text='恢复默认', command=on_reset).pack(side='left', padx=6)
        ttk.Button(btns, text='取消', command=win.destroy).pack(side='left', padx=6)
        center_on_root(win)          # 弹在主窗口中间
        win.wait_window()

    def revision_dialog():
        """填写族谱基本信息(名称/地址)与本次修订信息(人员/日期/说明), 随族谱文件保存。"""
        win = tk.Toplevel(root)
        win.title('修订信息')
        win.configure(bg=BG)
        win.transient(root)
        win.grab_set()
        win.resizable(False, False)

        tip = ttk.Label(win, foreground=SUBTEXT,
                        text='族谱基本信息与本次修订的人员、日期、说明'
                             '都随族谱文件一起保存，方便后人了解族谱来历与由谁修订。')
        tip.pack(anchor='w', padx=14, pady=(10, 2))

        frm = ttk.Frame(win)
        frm.pack(fill='both', expand=True, padx=14, pady=6)

        # ---- 族谱基本信息 ----
        ttk.Label(frm, text='◆ 族谱基本信息', foreground=PRIMARY,
                  font=FONT_B).grid(row=0, column=0, sticky='w', pady=(0, 2))
        ttk.Label(frm, text='族谱名称:', foreground=TEXT).grid(
            row=1, column=0, sticky='w')
        e_title = ttk.Entry(frm, width=44)
        e_title.grid(row=2, column=0, sticky='ew', pady=(2, 8))
        e_title.insert(0, CURRENT_TITLE or '')

        ttk.Label(frm, text='地址（祖籍/居住地/编修地等）:',
                  foreground=TEXT).grid(row=3, column=0, sticky='w')
        e_addr = ttk.Entry(frm, width=44)
        e_addr.grid(row=4, column=0, sticky='ew', pady=(2, 8))
        e_addr.insert(0, REV_INFO.get('address', ''))

        # ---- 修订信息 ----
        ttk.Label(frm, text='◆ 修订信息', foreground=PRIMARY,
                  font=FONT_B).grid(row=5, column=0, sticky='w', pady=(4, 2))
        ttk.Label(frm, text='修订人员（多人用顿号或逗号分隔）:',
                  foreground=TEXT).grid(row=6, column=0, sticky='w')
        e_who = ttk.Entry(frm, width=44)
        e_who.grid(row=7, column=0, sticky='ew', pady=(2, 8))
        e_who.insert(0, REV_INFO.get('revisers', ''))

        ttk.Label(frm, text='修订日期:', foreground=TEXT).grid(
            row=8, column=0, sticky='w')
        e_date = ttk.Entry(frm, width=44)
        e_date.grid(row=9, column=0, sticky='ew', pady=(2, 8))
        e_date.insert(0, REV_INFO.get('date') or time.strftime('%Y-%m-%d'))

        ttk.Label(frm, text='修订说明（修改了什么、为什么修订等）:',
                  foreground=TEXT).grid(row=10, column=0, sticky='w')
        t_note = tk.Text(frm, width=46, height=7, wrap='word', font=FONT,
                         bg=PANEL, fg=TEXT, relief='solid', borderwidth=1)
        t_note.grid(row=11, column=0, sticky='ew', pady=(2, 10))
        t_note.insert('1.0', REV_INFO.get('notes', ''))

        def on_ok():
            global CURRENT_TITLE
            CURRENT_TITLE = e_title.get().strip()
            REV_INFO['address'] = e_addr.get().strip()
            REV_INFO['revisers'] = e_who.get().strip()
            REV_INFO['date'] = e_date.get().strip()
            REV_INFO['notes'] = t_note.get('1.0', 'end').strip()
            root.title(app_title())
            set_dirty(True)
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(fill='x', padx=14, pady=(2, 12))
        ttk.Button(btns, text='保存', style='Green.TButton',
                   command=on_ok).pack(side='left', padx=6)
        ttk.Button(btns, text='取消', command=win.destroy).pack(side='left', padx=6)
        e_who.focus_set()
        center_on_root(win)
        win.wait_window()

    def show_help():
        txt_set(HELP_TEXT)

    def show_about():
        revs = []
        if state['path'] and os.path.exists(state['path']):
            try:
                with open(state['path'], encoding='utf-8') as f:
                    revs = json.load(f).get('revisions', [])
            except Exception:
                revs = []
        buf = '—————— 修订记录 ——————\n\n'
        buf += f'当前族谱: {state["path"] or "（未保存到文件）"}\n\n'
        if CURRENT_TITLE or REV_INFO.get('address'):
            buf += '族谱基本信息:\n'
            if CURRENT_TITLE:
                buf += f'  族谱名称: {CURRENT_TITLE}\n'
            if REV_INFO.get('address'):
                buf += f'  地址: {REV_INFO["address"]}\n'
            buf += '\n'
        if any(REV_INFO.values()):
            buf += '本次修订信息:\n'
            if REV_INFO.get('revisers'):
                buf += f'  修订人员: {REV_INFO["revisers"]}\n'
            if REV_INFO.get('date'):
                buf += f'  修订日期: {REV_INFO["date"]}\n'
            if REV_INFO.get('notes'):
                buf += f'  修订说明: {REV_INFO["notes"]}\n'
            buf += '\n'
        if revs:
            for i, r in enumerate(revs[-20:], 1):
                buf += f'{i}. {r}\n'
            if len(revs) > 20:
                buf += f'... 共 {len(revs)} 条\n'
        else:
            buf += '（暂无修订记录）'
        txt_set(buf)

    # ================= 界面布局 =================
    qvar = tk.StringVar()
    # 菜单栏
    menubar = tk.Menu(root)
    m_file = tk.Menu(menubar, tearoff=0)
    m_file.add_command(label='新建族谱', command=on_new)
    m_file.add_command(label='打开族谱...', command=on_open)
    m_file.add_command(label='打开示例族谱（明朝）', command=on_open_builtin)
    m_file.add_separator()
    m_file.add_command(label='保存', accelerator='Ctrl+S', command=lambda: do_save())
    m_file.add_command(label='另存为...', command=on_save_as)
    m_file.add_command(label='导出全谱文本...', command=on_export_txt)
    m_file.add_separator()
    m_file.add_command(label='退出', command=root.destroy)
    menubar.add_cascade(label='文件', menu=m_file)

    m_edit = tk.Menu(menubar, tearoff=0)
    m_edit.add_command(label='新增成员', command=on_add)
    m_edit.add_command(label='编辑成员', command=on_edit)
    m_edit.add_command(label='新增子女（选中者）', command=on_add_child)
    m_edit.add_command(label='删除成员', command=on_delete)
    m_edit.add_separator()
    m_edit.add_command(label='称呼设置...', command=terms_dialog)
    m_edit.add_command(label='修订信息...', command=revision_dialog)
    menubar.add_cascade(label='编辑', menu=m_edit)

    m_help = tk.Menu(menubar, tearoff=0)
    m_help.add_command(label='使用说明', command=show_help)
    m_help.add_command(label='关于/修订记录', command=show_about)
    menubar.add_cascade(label='帮助', menu=m_help)
    root.config(menu=menubar)

    # 左右分栏
    main = ttk.Frame(root)
    main.pack(fill='both', expand=True, padx=8, pady=0)

    left = tk.Frame(main, width=170, bg=PANEL, highlightbackground=BORDER,
                    highlightthickness=1)
    left.pack(side='left', fill='y', padx=(0, 8), pady=8)
    left.pack_propagate(False)
    lst = tk.Listbox(left, font=('Microsoft YaHei', 11), width=10,
                     bg=PANEL, fg=TEXT, selectbackground=LIST_SEL,
                     selectforeground=TEXT, borderwidth=0,
                     highlightthickness=0, activestyle='none')
    scr = ttk.Scrollbar(left, command=lst.yview)
    lst.config(yscrollcommand=scr.set)
    lst.pack(side='left', fill='both', expand=True)
    scr.pack(side='right', fill='y')

    right = ttk.Frame(main)
    right.pack(side='left', fill='both', expand=True, padx=(0, 8), pady=8)

    # ---- 显示区: 普通文本框(无特效, 稳定流畅) ----
    txtbox = tk.Frame(right, bg=PANEL, highlightbackground=BORDER,
                      highlightthickness=1)
    txt = tk.Text(txtbox, wrap='word', font=('Microsoft YaHei', 11),
                  bg=PANEL, fg=TEXT, relief='flat', borderwidth=0,
                  highlightthickness=0, padx=12, pady=8, state='disabled')
    vbar = ttk.Scrollbar(txtbox, orient='vertical', command=txt.yview)
    txt.config(yscrollcommand=vbar.set)
    txt.pack(side='left', fill='both', expand=True)
    vbar.pack(side='right', fill='y')

    # 照片条: 选中成员有照片时显示在文本区上方
    photo_bar = tk.Frame(right, bg=PANEL, highlightbackground=BORDER,
                         highlightthickness=1)
    photo_lbl = tk.Label(photo_bar, bg=PANEL)
    photo_lbl.pack(anchor='w', padx=10, pady=8)

    def show_photo(p):
        img = _b64_photo(getattr(p, 'photo', '') if p else '', (96, 120))
        if img is not None:
            photo_lbl.config(image=img)
            photo_lbl.image = img
            if not photo_bar.winfo_ismapped():
                photo_bar.pack(before=txtbox, fill='x', pady=(0, 3))
        else:
            photo_bar.pack_forget()

    def txt_set(*chunks):
        txt.config(state='normal')
        txt.delete('1.0', 'end')
        txt.insert('1.0', ''.join(chunks))
        txt.config(state='disabled')

    # ---- 列表/查询函数(需在工具栏之前定义) ----
    current_items = []
    sel_name = tk.StringVar()   # 记住最近操作/选中的成员名

    def search_list(q):
        q = q.strip()
        items = []
        for p in persons:
            hay = p.name + p.clean + ''.join(p.aliases)
            if not q or q in hay:
                items.append(p)
        return items

    def refresh_list():
        nonlocal current_items
        prev = get_selected().name if get_selected() else sel_name.get()
        items = search_list(qvar.get())
        items.sort(key=lambda p: (p.gen if p.gen is not None else 999,
                                  p.birth_order if p.birth_order else 999,
                                  p.name))
        current_items = items
        lst.delete(0, tk.END)
        for p in items:
            tag = f'{p.gen}世' if p.gen else '?'
            mark = ''
            if getattr(p, 'ru_zhui', False):
                mark = '·赘婿'
            elif getattr(p, 'married_in', False):
                mark = '·嫁入'
            elif getattr(p, 'adopted', False):
                mark = '·养'
            lst.insert(tk.END, f'{p.name}  ({tag}{mark})')
        # 恢复选中
        for i, p in enumerate(items):
            if p.name == prev:
                lst.selection_clear(0, tk.END)
                lst.selection_set(i)
                lst.see(i)
                break
        refresh_count()
        show_photo(get_selected())

    def show_full():
        ordered_roots = sorted(roots, key=lambda p: (p.gen if p.gen is not None else 999, p.name))
        txt_set('—————— 全谱 ——————\n\n', '\n'.join(full_tree_lines(ordered_roots)))

    # ---- 图形全谱(世系图: 世代分层卡片+吊线, 可缩放滚动) ----
    graph_state = {'win': None}

    def show_graph():
        old = graph_state['win']
        if old is not None and old.winfo_exists():
            old.lift()
            old.focus_force()
            return
        win = tk.Toplevel(root)
        graph_state['win'] = win
        win.withdraw()               # 先隐藏, 居中定位后再显示, 防闪烁
        win.title('世系图（图形全谱） ｜ 单击卡片：主窗口显示详情 ｜ '
                  '双击卡片：只看此人分支 ｜ Ctrl+滚轮缩放')
        win.geometry('1000x680')
        win.configure(bg=BG)
        win.transient(root)

        NODE_H = 50                       # 上下式卡片高
        HGAP, VGAP = 12, 30
        MARGIN = 24
        GEN_COLORS = ['#EAF6EE', '#E7F0FB', '#FDF3E0', '#F4EDFA', '#FDECEC']
        LINE_C = '#9CCFAC'
        COUPLE_C = '#E09595'              # 夫妻连线(粉红双线, 含入赘/标准婚姻)
        NAME_FS, SUB_FS = 11, 9           # 卡片字体(已加大)

        def _sub_of(p, limit=12):
            """卡片第二行: 世代 + 配偶, 超长截断。"""
            sub = f'{p.gen}世' if p.gen else ''
            if p.spouse:
                sub += (' ' if sub else '') + '配' + p.spouse
            if len(sub) > limit:
                sub = sub[:limit - 1] + '…'
            return sub

        # 卡片宽度自适应内容: 取名字/第二行实测较大者 + 内边距, 并封顶
        from tkinter import font as _tkfont
        _fname = _tkfont.Font(family='Microsoft YaHei',
                              size=NAME_FS, weight='bold')
        _fsub = _tkfont.Font(family='Microsoft YaHei', size=SUB_FS)
        w_card = {}
        for q in persons:
            w_card[q] = min(
                max(_fname.measure(q.name), _fsub.measure(_sub_of(q))) + 18,
                210)

        # 深度(沿父链上溯)
        dep = {}

        def depth_of(p):
            chain, cur = [], p
            while cur is not None and cur not in dep:
                chain.append(cur)
                cur = cur.father
            base = dep.get(cur, -1)
            for q in reversed(chain):
                base += 1
                dep[q] = base
            return dep[p]

        def kkey(p):
            return (p.gen if p.gen is not None else 999, p.name)

        # ---- 左右式布局(紧凑): 世代从左到右, 成员逐行排列 ----
        CH = 38                           # 左右式卡片高
        COL_W = max(w_card.values()) + 28  # 列宽=最宽卡片+间距
        ROW_H = 46
        pos_lr = {}
        row_i = {'n': 0}
        nodes_set = set(persons)          # 当前视图包含的人(分支视图会缩小)

        def colx(p):
            return MARGIN + depth_of(p) * COL_W + w_card[p] / 2.0

        def layout_lr(p):
            s_ = getattr(p, 'ru_spouse', None)
            if getattr(p, 'ru_zhui', False) and s_ is not None \
                    and s_ in nodes_set:
                # 赘婿: 妻子若未布局则先布(夫妻单位一起布), 自己跟随妻子
                if s_ not in pos_lr:
                    layout_lr(s_)
                return pos_lr[p][1]
            if (s_ is not None and getattr(s_, 'ru_zhui', False)
                    and s_ in nodes_set and s_ not in pos_lr):
                # 夫妻单位: 妻子p在上, 赘婿s_紧贴其下一行;
                # 子女取两人并集, 从夫妻中点引出
                union, ids = [], set()
                for k in list(p.children) + list(s_.children):
                    if id(k) not in ids:
                        ids.add(id(k))
                        union.append(k)
                if union:
                    cs = [layout_lr(k) for k in union]
                    cy = (min(cs) + max(cs)) / 2.0
                    yw, ym = cy - ROW_H / 2.0, cy + ROW_H / 2.0
                else:
                    yw = MARGIN + row_i['n'] * ROW_H + CH / 2.0
                    row_i['n'] += 1
                    ym = MARGIN + row_i['n'] * ROW_H + CH / 2.0
                    row_i['n'] += 1
                    cy = (yw + ym) / 2.0
                pos_lr[p] = (colx(p), yw)
                pos_lr[s_] = (colx(s_), ym)
                return cy
            kids = list(p.children)
            if kids:
                cs = [layout_lr(k) for k in kids]
                cy = (min(cs) + max(cs)) / 2.0
            else:
                cy = MARGIN + row_i['n'] * ROW_H + CH / 2.0
                row_i['n'] += 1
            pos_lr[p] = (colx(p), cy)
            return cy

        # ---- 上下式布局(传统吊线图): 子树占横向区间, 父居中 ----
        pos_tb = {}
        cursor = {'x': MARGIN}

        def layout_tb(p):
            y = depth_of(p) * (NODE_H + VGAP)
            s_ = getattr(p, 'ru_spouse', None)
            if getattr(p, 'ru_zhui', False) and s_ is not None \
                    and s_ in nodes_set:
                if s_ not in pos_tb:
                    layout_tb(s_)
                return pos_tb[p][0]
            if (s_ is not None and getattr(s_, 'ru_zhui', False)
                    and s_ in nodes_set and s_ not in pos_tb):
                # 夫妻单位: 同一行, 妻子与赘婿水平紧邻
                union, ids = [], set()
                for k in list(p.children) + list(s_.children):
                    if id(k) not in ids:
                        ids.add(id(k))
                        union.append(k)
                if union:
                    cs = [layout_tb(k) for k in union]
                    cx = (min(cs) + max(cs)) / 2.0
                    xw = cx - (w_card[p] / 2 + HGAP / 2.0)
                    xm = cx + (w_card[s_] / 2 + HGAP / 2.0)
                else:
                    xw = cursor['x'] + w_card[p] / 2
                    cursor['x'] += w_card[p] + HGAP
                    xm = cursor['x'] + w_card[s_] / 2
                    cursor['x'] += w_card[s_] + HGAP
                    cx = (xw + xm) / 2.0
                pos_tb[p] = (xw, y)
                pos_tb[s_] = (xm, y)
                return cx
            kids = list(p.children)
            if kids:
                cs = [layout_tb(k) for k in kids]
                cx = (min(cs) + max(cs)) / 2.0
            else:
                cx = cursor['x'] + w_card[p] / 2.0
                cursor['x'] += w_card[p] + HGAP
            pos_tb[p] = (cx, y)
            return cx

        roots_all_cache = sorted((q for q in persons if q.father is None), key=kkey)
        view_root = {'p': None}      # 双击某人 -> 只看此人分支; None=全谱
        total = {'lr': (0, 0), 'tb': (0, 0)}

        def _fix_tb_overlap():
            """上下式同深度卡片重叠修复: 逐层从左到右扫描,
            重叠者连同整棵子树(及其配偶)右移让位。"""
            def subtree_shift(start, dx):
                seen, stack = set(), [start]
                while stack:
                    q = stack.pop()
                    if id(q) in seen or q not in pos_tb:
                        continue
                    seen.add(id(q))
                    x, y = pos_tb[q]
                    pos_tb[q] = (x + dx, y)
                    sp = getattr(q, 'ru_spouse', None)
                    if sp is not None and sp in pos_tb and id(sp) not in seen:
                        sx, sy = pos_tb[sp]
                        pos_tb[sp] = (sx + dx, sy)
                        seen.add(id(sp))
                        stack.extend(sp.children)
                    stack.extend(q.children)
            by_dep = {}
            for q in pos_tb:
                by_dep.setdefault(dep[q], []).append(q)
            for d in sorted(by_dep):
                ps = sorted(by_dep[d], key=lambda q: pos_tb[q][0])
                for a, b in zip(ps, ps[1:]):
                    need = pos_tb[a][0] + w_card[a] / 2 + HGAP + w_card[b] / 2
                    if pos_tb[b][0] < need - 0.5:
                        subtree_shift(b, need - pos_tb[b][0])

        def rebuild():
            """按当前视图根(全谱或某人分支)重新计算 dep/布局/尺寸。"""
            R = view_root['p']
            dep.clear()
            nodes_set.clear()
            if R is None:
                nodes_set.update(persons)
                for q in persons:
                    depth_of(q)
                tree_roots = roots_all_cache
            else:
                nodes_set.add(R)
                dep[R] = 0           # 分支视图: 深度相对此人重新计
                stack = [R]
                while stack:
                    q = stack.pop()
                    for k in q.children:
                        nodes_set.add(k)
                        dep[k] = dep[q] + 1
                        stack.append(k)
                tree_roots = [R]
            pos_lr.clear()
            row_i['n'] = 0
            pos_tb.clear()
            cursor['x'] = MARGIN
            for r in tree_roots:
                layout_lr(r)
            if R is None:
                for q in sorted(persons, key=kkey):
                    if q not in pos_lr and q.father not in pos_lr:
                        row_i['n'] += 1     # 独立树之间空一行
                        layout_lr(q)
            for r in tree_roots:
                cursor['x'] += HGAP
                layout_tb(r)
            if R is None:
                for q in sorted(persons, key=kkey):
                    if q not in pos_tb and q.father not in pos_tb:
                        cursor['x'] += HGAP * 2
                        layout_tb(q)
            _fix_tb_overlap()
            max_dep = max(dep.values())
            total['lr'] = (MARGIN * 2 + max_dep * COL_W + max(w_card.values()),
                           MARGIN * 2 + row_i['n'] * ROW_H)
            tb_max_x = max((pos_tb[q][0] + w_card[q] / 2 for q in pos_tb),
                           default=MARGIN)
            total['tb'] = (tb_max_x + MARGIN,
                           (max_dep + 1) * (NODE_H + VGAP) + MARGIN)

        rebuild()

        # 顶部工具条
        bar = tk.Frame(win, bg=BG)
        bar.pack(fill='x', padx=8, pady=4)
        view_var = tk.StringVar(value='当前: 全谱')
        tk.Label(bar, textvariable=view_var, bg=BG, fg=PRIMARY,
                 font=(FONT[0], 9, 'bold')).pack(side='left')
        scale = {'v': 1.0}
        mode = {'v': 'lr'}   # lr=左右式(紧凑, 默认) / tb=上下式(传统吊线)
        mode_btns = {}

        def set_mode(m):
            mode['v'] = m
            mode_btns['lr'].config(text='左右式 ●' if m == 'lr' else '左右式')
            mode_btns['tb'].config(text='上下式 ●' if m == 'tb' else '上下式')
            redraw()

        def set_scale(v):
            scale['v'] = max(0.4, min(2.5, v))
            redraw()

        def fit_width():
            vw = canvas.winfo_width() or 1000
            set_scale(vw / max(total[mode['v']][0], 1))

        def back_to_full():
            view_root['p'] = None
            view_var.set('当前: 全谱')
            rebuild()
            redraw()
            fit_width()

        def export_pdf():
            """把当前视图(全谱或分支, 左右式/上下式)渲染成 PDF。"""
            import tkinter.filedialog as _fd
            from tkinter import messagebox as _mb
            try:
                from PIL import Image, ImageDraw, ImageFont
            except ImportError:
                _mb.showerror('导出失败', '缺少 Pillow 库, 无法导出 PDF。',
                              parent=win)
                return
            path = _fd.asksaveasfilename(
                parent=win, title='导出世系图为 PDF',
                defaultextension='.pdf', initialfile='世系图.pdf',
                filetypes=[('PDF 文件', '*.pdf')])
            if not path:
                return
            try:
                SS = 2                       # 超采样倍数: 2倍渲染+等效300DPI
                mv = mode['v']
                pos = pos_lr if mv == 'lr' else pos_tb
                tw, th = int(total[mv][0]) + 1, int(total[mv][1]) + 1
                img = Image.new('RGB', (tw * SS, th * SS), 'white')
                dr = ImageDraw.Draw(img)

                def X(v):
                    return v * SS

                try:
                    f_name = ImageFont.truetype(
                        'C:/Windows/Fonts/msyhbd.ttc',
                        round(NAME_FS * 4 / 3 * SS))
                    f_sub = ImageFont.truetype(
                        'C:/Windows/Fonts/msyh.ttc',
                        round(SUB_FS * 4 / 3 * SS))
                except OSError:
                    f_name = f_sub = ImageFont.load_default()

                # 夫妻连线(与屏幕绘制同一套规则, 含入赘/标准婚姻)
                couple_mid = {}
                done_pair = set()
                for m, (cx, cy) in pos.items():
                    s_ = getattr(m, 'ru_spouse', None) or getattr(m, 'spouse_ref', None)
                    if s_ not in pos or m in done_pair:
                        continue
                    done_pair.update((id(m), id(s_)))
                    wx, wy = pos[s_]
                    my = (cy + wy) / 2.0
                    if mv == 'lr':
                        same = dep[m] == dep[s_]
                        if same:
                            mx = MARGIN + dep[m] * COL_W + COL_W / 2.0
                            ya = min(cy, wy) + CH / 2
                            yb = max(cy, wy) - CH / 2
                            for dx in (-3, 3):
                                dr.line((X(mx + dx), X(ya),
                                         X(mx + dx), X(yb)),
                                        fill=COUPLE_C, width=SS)
                        else:
                            mx = (cx + wx) / 2.0
                            dr.line((X(cx), X(cy), X(wx), X(wy)),
                                    fill=COUPLE_C, width=SS)
                    else:
                        same = dep[m] == dep[s_]
                        if same:
                            lp, rp = (m, s_) if cx < wx else (s_, m)
                            xa = pos[lp][0] + w_card[lp] / 2
                            xb = pos[rp][0] - w_card[rp] / 2
                            mx = (xa + xb) / 2.0
                            my = cy + NODE_H
                            for dy in (-3, 3):
                                dr.line((X(xa), X(my + dy),
                                         X(xb), X(my + dy)),
                                        fill=COUPLE_C, width=SS)
                        else:
                            mx = (cx + wx) / 2.0
                            dr.line((X(cx), X(cy), X(wx), X(wy)),
                                    fill=COUPLE_C, width=SS)
                    couple_mid[m] = (mx, my)
                    couple_mid[s_] = (mx, my)

                # 父子吊线
                for p, (cx, cy) in pos.items():
                    f = p.father
                    if f not in pos:
                        continue
                    fx, fy = pos[f]
                    if f in couple_mid:
                        xa, ya = couple_mid[f]
                    elif mv == 'lr':
                        xa, ya = fx + w_card[f] / 2, fy
                    else:
                        xa, ya = fx, fy + NODE_H
                    if mv == 'lr':
                        xb, yb = cx - w_card[p] / 2, cy
                        xm = (xa + xb) / 2.0
                        dr.line((X(xa), X(ya), X(xm), X(ya),
                                 X(xm), X(yb), X(xb), X(yb)),
                                fill=LINE_C, width=SS)
                    else:
                        xb, yb = cx, cy
                        ym = (ya + yb) / 2.0
                        dr.line((X(xa), X(ya), X(xa), X(ym),
                                 X(xb), X(ym), X(xb), X(yb)),
                                fill=LINE_C, width=SS)

                # 卡片
                for p, (cx, cy) in pos.items():
                    color = GEN_COLORS[dep[p] % len(GEN_COLORS)]
                    sub = _sub_of(p)
                    if mv == 'lr':
                        hw = w_card[p] / 2
                        x0, y0, x1, y1 = cx - hw, cy - CH / 2, cx + hw, cy + CH / 2
                        dr.rectangle((X(x0), X(y0), X(x1), X(y1)),
                                     fill=color, outline=BORDER, width=SS)
                        ny = y0 + (y1 - y0) * 0.30
                        sy = y0 + (y1 - y0) * 0.74
                        dr.text((X(x0 + 7), X(ny)), p.name, font=f_name,
                                fill=TEXT, anchor='lm')
                        dr.text((X(x0 + 7), X(sy)), sub, font=f_sub,
                                fill=SUBTEXT, anchor='lm')
                    else:
                        hw = w_card[p] / 2
                        x0, y0, x1, y1 = (cx - hw, cy, cx + hw, cy + NODE_H)
                        dr.rectangle((X(x0), X(y0), X(x1), X(y1)),
                                     fill=color, outline=BORDER, width=SS)
                        dr.text((X((x0 + x1) / 2), X(y0 + (y1 - y0) * 0.32)),
                                p.name, font=f_name, fill=TEXT, anchor='mm')
                        dr.text((X((x0 + x1) / 2), X(y0 + (y1 - y0) * 0.72)),
                                sub, font=f_sub, fill=SUBTEXT, anchor='mm')

                img.save(path, 'PDF', resolution=150 * SS)
                _mb.showinfo('导出成功', '已导出到:\n%s' % path, parent=win)
            except Exception as e:
                _mb.showerror('导出失败', str(e), parent=win)

        ttk.Button(bar, text='返回全谱', width=7,
                   command=back_to_full).pack(side='right', padx=2)

        ttk.Button(bar, text='＋', width=3,
                   command=lambda: set_scale(scale['v'] * 1.2)).pack(side='right', padx=2)
        ttk.Button(bar, text='－', width=3,
                   command=lambda: set_scale(scale['v'] / 1.2)).pack(side='right', padx=2)
        ttk.Button(bar, text='导出PDF', width=7,
                   command=export_pdf).pack(side='right', padx=2)
        ttk.Button(bar, text='适宽', width=5,
                   command=fit_width).pack(side='right', padx=2)
        mode_btns['tb'] = ttk.Button(bar, text='上下式', width=7,
                                     command=lambda: set_mode('tb'))
        mode_btns['tb'].pack(side='right', padx=2)
        mode_btns['lr'] = ttk.Button(bar, text='左右式 ●', width=7,
                                     command=lambda: set_mode('lr'))
        mode_btns['lr'].pack(side='right', padx=2)

        # 画布 + 滚动条
        cf = tk.Frame(win, bg=BG)
        cf.pack(fill='both', expand=True)
        canvas = tk.Canvas(cf, bg='white', highlightthickness=0)
        vsb = ttk.Scrollbar(cf, orient='vertical', command=canvas.yview)
        hsb = ttk.Scrollbar(cf, orient='horizontal', command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        cf.grid_rowconfigure(0, weight=1)
        cf.grid_columnconfigure(0, weight=1)

        cur = {'tag': None, 'rect': None}
        rect_ids = {}   # tag -> rectangle item id(仅矩形可设 outline)

        def redraw():
            s = scale['v']
            canvas.delete('all')
            rect_ids.clear()
            cur['tag'] = None
            cur['rect'] = None

            def _card(p, x0, y0, x1, y1, sub, centered):
                color = GEN_COLORS[dep[p] % len(GEN_COLORS)]
                tag = 'n%d' % id(p)
                rect_id = canvas.create_rectangle(x0, y0, x1, y1, fill=color,
                                                  outline=BORDER, width=1,
                                                  tags=(tag, 'rect'))
                rect_ids[tag] = rect_id
                fs = max(7, round((NAME_FS if centered else NAME_FS - 1) * s))
                fs2 = max(6, round(SUB_FS * s))
                if centered:
                    nx, sx = (x0 + x1) / 2, (x0 + x1) / 2
                    ny = y0 + (y1 - y0) * 0.32
                    sy = y0 + (y1 - y0) * 0.72
                    anchor = 'center'
                else:
                    nx = sx = x0 + 7 * s
                    ny = y0 + (y1 - y0) * 0.30
                    sy = y0 + (y1 - y0) * 0.74
                    anchor = 'w'
                canvas.create_text(nx, ny, text=p.name, anchor=anchor,
                                   font=('Microsoft YaHei', fs, 'bold'),
                                   fill=TEXT, tags=tag)
                canvas.create_text(sx, sy, text=sub, anchor=anchor,
                                   font=('Microsoft YaHei', fs2),
                                   fill=SUBTEXT, tags=tag)

                def on_click(e, p=p, tag=tag):
                    if cur['rect'] is not None:
                        canvas.itemconfig(cur['rect'], outline=BORDER, width=1)
                    rect_id = rect_ids.get(tag)
                    if rect_id is not None:
                        canvas.itemconfig(rect_id, outline=PRIMARY, width=2)
                        cur['rect'] = rect_id
                    cur['tag'] = tag
                    try:
                        idx = current_items.index(p)
                        lst.selection_clear(0, tk.END)
                        lst.selection_set(idx)
                        lst.see(idx)
                    except ValueError:
                        pass
                    sel_name.set(p.name)
                    txt_set(person_detail(p))
                    show_photo(p)

                def on_enter(e):
                    canvas.config(cursor='hand2')

                def on_leave(e):
                    canvas.config(cursor='')

                def on_dbl(e, p=p):
                    # 双击: 以此人为根, 只显示其分支
                    view_root['p'] = p
                    view_var.set('当前: %s 的分支' % p.name)
                    rebuild()
                    redraw()
                    fit_width()

                canvas.tag_bind(tag, '<Button-1>', on_click)
                canvas.tag_bind(tag, '<Double-Button-1>', on_dbl)
                canvas.tag_bind(tag, '<Enter>', on_enter)
                canvas.tag_bind(tag, '<Leave>', on_leave)

            if mode['v'] == 'lr':
                pos = pos_lr
                # 夫妻连线: 赘婿与妻子/标准婚姻配偶(同一世代列)之间画双线
                couple_mid = {}      # 丈夫/妻子 -> 两人连线中点(引出子女)
                done_pair = set()
                for m, (cx, cy) in pos.items():
                    s_ = getattr(m, 'ru_spouse', None) or getattr(m, 'spouse_ref', None)
                    if s_ not in pos or m in done_pair:
                        continue
                    done_pair.add(m)
                    done_pair.add(s_)
                    wx, wy = pos[s_]
                    same_col = dep[m] == dep[s_]   # 同一世代列
                    my = (cy + wy) / 2.0
                    if same_col:
                        mx = MARGIN + dep[m] * COL_W + COL_W / 2.0
                        ya = min(cy, wy) + CH / 2
                        yb = max(cy, wy) - CH / 2
                        for dx in (-3, 3):
                            canvas.create_line(mx * s + dx, ya * s,
                                               mx * s + dx, yb * s,
                                               fill=COUPLE_C, width=1)
                    else:
                        mx = (cx + wx) / 2.0
                        canvas.create_line(cx * s, cy * s, wx * s, wy * s,
                                           fill=COUPLE_C, width=1)
                    couple_mid[m] = (mx, my)
                    couple_mid[s_] = (mx, my)
                # 吊线: 父右边缘中点 -> 横线 -> 竖线 -> 子左边缘中点
                # (入赘夫妻的子女改从夫妻连线中点引出)
                for p, (cx, cy) in pos.items():
                    f = p.father
                    if f not in pos:
                        continue
                    fx, fy = pos[f]
                    if f in couple_mid:
                        xa, ya = couple_mid[f]
                        xa, ya = xa * s, ya * s
                    else:
                        xa, ya = (fx + w_card[f] / 2) * s, fy * s
                    xb, yb = (cx - w_card[p] / 2) * s, cy * s
                    xm = (xa + xb) / 2.0
                    canvas.create_line(xa, ya, xm, ya, xm, yb, xb, yb,
                                       fill=LINE_C, width=1)
                # 卡片: 名字左对齐, 第二行 世代+配偶
                for p, (cx, cy) in pos.items():
                    hw = w_card[p] / 2.0
                    x0 = (cx - hw) * s
                    y0 = (cy - CH / 2) * s
                    x1 = (cx + hw) * s
                    y1 = (cy + CH / 2) * s
                    _card(p, x0, y0, x1, y1, _sub_of(p), centered=False)
            else:
                pos = pos_tb
                # 夫妻连线: 同一世代的夫妻卡之间画双线(入赘/标准婚姻)
                couple_mid = {}
                done_pair = set()
                for m, (cx, cy) in pos.items():
                    s_ = getattr(m, 'ru_spouse', None) or getattr(m, 'spouse_ref', None)
                    if s_ not in pos or m in done_pair:
                        continue
                    done_pair.add(m)
                    done_pair.add(s_)
                    wx, wy = pos[s_]
                    same_row = dep[m] == dep[s_]   # 同一世代行
                    if same_row:
                        lp, rp = (m, s_) if cx < wx else (s_, m)
                        xa = pos[lp][0] + w_card[lp] / 2
                        xb = pos[rp][0] - w_card[rp] / 2
                        mx = (xa + xb) / 2.0
                        my = cy + NODE_H
                        for dy in (-3, 3):
                            canvas.create_line(xa * s, my * s + dy,
                                               xb * s, my * s + dy,
                                               fill=COUPLE_C, width=1)
                    else:
                        mx = (cx + wx) / 2.0
                        my = (cy + wy) / 2.0
                        canvas.create_line(cx * s, cy * s, wx * s, wy * s,
                                           fill=COUPLE_C, width=1)
                    couple_mid[m] = (mx, my)
                    couple_mid[s_] = (mx, my)
                # 吊线: 父底边中点 -> 竖线 -> 横线 -> 子顶边中点
                # (入赘夫妻的子女改从夫妻连线中点引出)
                for p, (fx, fy) in pos.items():
                    if p.father not in pos:
                        continue
                    px, py = pos[p.father]
                    if p.father in couple_mid:
                        xa, ya = couple_mid[p.father]
                        xa, ya = xa * s, ya * s
                    else:
                        xa, ya = px * s, (py + NODE_H) * s
                    xb, yb = fx * s, fy * s
                    ym = (ya + yb) / 2.0
                    canvas.create_line(xa, ya, xa, ym, xb, ym, xb, yb,
                                       fill=LINE_C, width=1)
                for p, (cx, y) in pos.items():
                    hw = w_card[p] / 2.0
                    x0 = (cx - hw) * s
                    y0 = y * s
                    x1 = (cx + hw) * s
                    y1 = (y + NODE_H) * s
                    _card(p, x0, y0, x1, y1, _sub_of(p), centered=True)
            tw, th = total[mode['v']]
            canvas.configure(scrollregion=(0, 0, tw * s, th * s))

        def on_wheel(e):
            canvas.yview_scroll(-1 if e.delta > 0 else 1, 'units')

        def on_shift_wheel(e):
            canvas.xview_scroll(-1 if e.delta > 0 else 1, 'units')

        def on_ctrl_wheel(e):
            set_scale(scale['v'] * (1.15 if e.delta > 0 else 1 / 1.15))
            return 'break'

        canvas.bind('<MouseWheel>', on_wheel)
        canvas.bind('<Shift-MouseWheel>', on_shift_wheel)
        canvas.bind('<Control-MouseWheel>', on_ctrl_wheel)
        redraw()

        # 打开时默认适宽(等画布有实际尺寸后再算, 只执行一次)
        _fitted = {'done': False}

        def _fit_on_open(e=None):
            if _fitted['done'] or canvas.winfo_width() < 200:
                return
            _fitted['done'] = True
            fit_width()
            canvas.unbind('<Configure>')

        canvas.bind('<Configure>', _fit_on_open)
        win.after(200, _fit_on_open)
        # 弹出在主窗口中间(窗口固定 1000x680, 不能用内容请求尺寸算)
        win.update_idletasks()
        _gx = root.winfo_x() + max(0, (root.winfo_width() - 1000) // 2)
        _gy = root.winfo_y() + max(0, (root.winfo_height() - 680) // 2)
        win.geometry('1000x680+%d+%d' % (_gx, _gy))
        win.deiconify()

    # 顶部工具栏(before=main 保持其在窗口顶部; 函数此时均已定义)
    top = tk.Frame(root, bg=BG)
    top.pack(fill='x', before=main)
    # 查询输入框: 宽度加大使"查询"按钮左缘与下方列表滚动条右缘对齐
    qentry = ttk.Entry(top, textvariable=qvar, width=22)
    qentry.pack(side='left', padx=(8, 2), pady=1)
    qentry.bind('<Return>', lambda e: refresh_list())

    ttk.Button(top, text='查询', style='Header.TButton',
               command=lambda: refresh_list()).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='显示全谱', style='Header.TButton',
               command=show_full).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='图形全谱', style='Header.TButton',
               command=show_graph).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='清空', style='Header.TButton',
               command=lambda: (qvar.set(''), txt_set(), refresh_list())).pack(
        side='left', padx=1, pady=1)
    # 编辑区按钮
    ttk.Separator(top, orient='vertical').pack(side='left', fill='y', padx=4, pady=3)
    ttk.Button(top, text='新增成员', style='Green.TButton',
               command=on_add).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='编辑成员', style='Green.TButton',
               command=on_edit).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='删除成员', style='Green.TButton',
               command=on_delete).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='保存', style='Green.TButton',
               command=lambda: do_save()).pack(side='left', padx=(5, 1), pady=1)
    ttk.Button(top, text='称呼设置', style='Green.TButton',
               command=terms_dialog).pack(side='left', padx=1, pady=1)
    ttk.Button(top, text='修订信息', style='Green.TButton',
               command=revision_dialog).pack(side='left', padx=1, pady=1)
    count_lbl = ttk.Label(top, text='', background=BG, foreground=SUBTEXT)
    count_lbl.pack(side='right', padx=8, pady=1)

    # ---- 互称面板 ----
    # 布局: 上两行(标题行/甲乙行)为左列, 右列放互称结果(与上两行同高);
    #       下方左右分栏: 左=甲信息, 右=乙信息
    a_var = tk.StringVar()
    b_var = tk.StringVar()
    # 互称面板顶部分割线(与上方内容分隔)
    ttk.Separator(right, orient='horizontal').pack(fill='x')
    kin_frame = tk.Frame(right, bg=BG)
    kin_frame.pack(fill='x', expand=False, padx=0, pady=0)
    # 正文显示区在互称面板下方占满剩余空间
    # (修复: txtbox 此前丢失 pack, 导致"显示全谱"等内容无处显示)
    txtbox.pack(fill='both', expand=True, pady=(6, 0))
    # 三列: 甲信息 | 竖分割线 | 乙信息(甲乙列等宽)
    kin_frame.grid_columnconfigure(0, weight=1, uniform='kin')
    kin_frame.grid_columnconfigure(2, weight=1, uniform='kin')

    lock_status = tk.StringVar(value='未锁定')
    header_row = tk.Frame(kin_frame, bg=BG)
    header_row.grid(row=0, column=0, columnspan=3, sticky='ew',
                    padx=(8, 0), pady=(2, 0))
    ttk.Label(header_row, text='互称（左侧双击人名填入甲、乙）',
              background=BG, foreground=SUBTEXT).pack(side='left')
    tk.Label(header_row, textvariable=lock_status, fg=SUBTEXT, bg=BG,
             font=(FONT[0], 8)).pack(side='left', padx=6)

    name_row = tk.Frame(kin_frame, bg=BG)
    name_row.grid(row=1, column=0, columnspan=3, sticky='ew',
                  padx=(8, 0), pady=4)
    btn_a = ttk.Button(name_row, text='甲', width=3, style='Green.TButton',
                       command=lambda: set_lock('甲'))
    btn_a.pack(side='left', padx=4)
    tk.Label(name_row, textvariable=a_var, width=12, bg=PANEL, fg=TEXT,
             relief='sunken', borderwidth=1, anchor='center',
             font=FONT).pack(side='left', padx=4)
    btn_b = ttk.Button(name_row, text='乙', width=3, style='Green.TButton',
                       command=lambda: set_lock('乙'))
    btn_b.pack(side='left', padx=(12, 4))
    tk.Label(name_row, textvariable=b_var, width=12, bg=PANEL, fg=TEXT,
             relief='sunken', borderwidth=1, anchor='center',
             font=FONT).pack(side='left', padx=4)

    # 右列: 互称结果(两行, 与标题行/甲乙行同高, 节省一整行空间)
    kin_card = tk.Frame(kin_frame, bg=BUBBLE)
    kin_card.grid(row=0, column=2, rowspan=2, sticky='nsew',
                  padx=(2, 8), pady=(2, 4))
    kin_lbl1 = tk.Label(kin_card, text='左侧双击人名', bg=BUBBLE, fg=TEXT,
                        anchor='w', font=(FONT[0], 10, 'bold'))
    kin_lbl1.pack(fill='x', padx=8, pady=(3, 0))
    kin_lbl2 = tk.Label(kin_card, text='依次填入甲、乙', bg=BUBBLE, fg=TEXT,
                        anchor='w', font=(FONT[0], 10))
    kin_lbl2.pack(fill='x', padx=8, pady=(0, 3))

    # 上半区与信息栏之间的水平分割线
    ttk.Separator(kin_frame, orient='horizontal').grid(
        row=2, column=0, columnspan=3, sticky='ew', pady=(2, 2))

    # 下半区: 左=甲信息, 右=乙信息, 中间竖分割线
    def _info_text():
        return tk.Text(kin_frame, height=7, wrap='word',
                       font=(FONT[0], 9), relief='flat',
                       bg=PANEL, fg=TEXT, borderwidth=0,
                       highlightthickness=1, highlightbackground=BORDER,
                       state='disabled', padx=8, pady=4)

    info_a = _info_text()
    info_a.grid(row=3, column=0, sticky='nsew', padx=(8, 4), pady=(0, 2))
    ttk.Separator(kin_frame, orient='vertical').grid(
        row=3, column=1, sticky='ns', padx=2, pady=(0, 2))
    info_b = _info_text()
    info_b.grid(row=3, column=2, sticky='nsew', padx=(4, 8), pady=(0, 2))

    # 面板底部分割线
    ttk.Separator(kin_frame, orient='horizontal').grid(
        row=4, column=0, columnspan=3, sticky='ew', pady=(2, 4))

    def _set_info(w, text):
        w.config(state='normal')
        w.delete('1.0', tk.END)
        w.insert('1.0', text)
        w.config(state='disabled')

    def brief_info(p):
        """互称面板用的人物摘要: 有什么显示什么, 没有的不显示"""
        if p is None:
            return '（未选择）'
        lines = []
        gen = f'{p.gen}世' if p.gen else '世代未知'
        label = f'（{p.gen_label}）' if p.gen_label else ''
        lines.append(f'{p.name}{label} {gen}')
        if p.zi:
            lines.append(f'字: {p.zi}')
        if p.aliases:
            lines.append(f'别名: {"、".join(p.aliases)}')
        if p.birth_order:
            lines.append(f'排行: {ORDER_WORD.get(p.birth_order, p.birth_order)}')
        if p.birth_date:
            lines.append(f'出生: {p.birth_date}')
        # 年龄/享年
        if getattr(p, 'deceased', False):
            if getattr(p, 'age_at_death', None):
                lines.append(f'享年: {p.age_at_death}岁')
            elif getattr(p, 'death_year', None) and p.birth_year:
                lines.append(f'享年: 约{p.death_year - p.birth_year}岁')
            else:
                lines.append('已去世')
        elif p.birth_year:
            age = datetime.date.today().year - p.birth_year
            lines.append(f'年龄: 约{age}岁')
        if p.adopted:
            lines.append('身份: 养子/养女')
        elif p.ru_zhui:
            lines.append(f'身份: 赘婿（入赘{p.spouse}家）')
        elif getattr(p, 'married_in', False):
            lines.append(f'身份: 嫁入（适{p.spouse}）')
        if p.spouse and not p.ru_zhui and not getattr(p, 'married_in', False):
            lines.append(f'配偶: {p.spouse}')
        if p.father:
            lines.append(f'父亲: {p.father.name}')
        if getattr(p, 'mother', None):
            lines.append(f'母亲: {p.mother.name}')
        sibs = p.siblings()
        if sibs:
            lines.append(f'兄弟姐妹: {"、".join(_sib_disp_name(p, s) for s in sibs)}')
        kids = p.children
        if kids:
            shown = [_child_disp_name(k) for k in kids]
            lines.append(f'子女: {"、".join(shown)}')
        if p.burial:
            lines.append(f'葬地: {p.burial}')
        return '\n'.join(lines)

    lock_mode = {'who': None}

    def set_lock(who):
        lock_mode['who'] = who if lock_mode['who'] != who else None
        if lock_mode['who'] is None:
            lock_status.set('未锁定：左侧点名依次填甲、乙')
        elif lock_mode['who'] == '甲':
            lock_status.set('已锁定甲，左侧点名只改乙')
        else:
            lock_status.set('已锁定乙，左侧点名只改甲')
        btn_a.config(text='甲 ✓' if lock_mode['who'] == '甲' else '甲')
        btn_b.config(text='乙 ✓' if lock_mode['who'] == '乙' else '乙')

    # ---- 列表事件(函数已在显示区前定义) ----
    def on_select(ev):
        p = get_selected()
        if p:
            sel_name.set(p.name)
            txt_set(person_detail(p))
            show_photo(p)

    def find_person(name):
        return next((x for x in persons if x.name == name), None)

    def compute_kin():
        na, nb = a_var.get(), b_var.get()
        A = find_person(na) if na else None
        B = find_person(nb) if nb else None
        # 下方左右信息栏: 填了谁就显示谁的信息
        _set_info(info_a, brief_info(A))
        _set_info(info_b, brief_info(B))
        # 右上结果区: 两行互称结果
        if not na or not nb:
            return
        if not A or not B:
            kin_lbl1.config(text='未找到对应人物')
            kin_lbl2.config(text='')
            return
        t1 = kinship_term(A, B)
        t2 = kinship_term(B, A)
        kin_lbl1.config(text=f'{na} 称 {nb}：{t1}')
        kin_lbl2.config(text=f'{nb} 称 {na}：{t2}')

    def on_double(ev):
        sel = lst.nearest(ev.y)
        if sel < 0 or sel >= len(current_items):
            return
        lst.selection_clear(0, tk.END)
        lst.selection_set(sel)
        name = current_items[sel].name
        who = lock_mode['who']
        if who == '甲':
            b_var.set(name)
        elif who == '乙':
            a_var.set(name)
        else:
            if not a_var.get() or (a_var.get() and b_var.get()):
                a_var.set(name)
                b_var.set('')
            elif name != a_var.get():
                b_var.set(name)
        compute_kin()

    # 右键菜单
    def _ctx_view():
        p = get_selected()
        if p:
            txt_set(person_detail(p))
            show_photo(p)

    ctx = tk.Menu(root, tearoff=0)
    ctx.add_command(label='查看', command=_ctx_view)
    ctx.add_command(label='编辑成员', command=on_edit)
    ctx.add_command(label='新增子女', command=on_add_child)
    ctx.add_separator()
    ctx.add_command(label='设为甲', command=lambda: _pick('甲'))
    ctx.add_command(label='设为乙', command=lambda: _pick('乙'))
    ctx.add_separator()
    ctx.add_command(label='删除成员', command=on_delete)

    def _pick(who):
        p = get_selected()
        if not p:
            return
        if who == '甲':
            a_var.set(p.name)
        else:
            b_var.set(p.name)
        compute_kin()

    def on_right_click(ev):
        idx = lst.nearest(ev.y)
        if idx < 0 or idx >= len(current_items):
            return
        lst.selection_clear(0, tk.END)
        lst.selection_set(idx)
        p = current_items[idx]
        sel_name.set(p.name)
        txt_set(person_detail(p))
        ctx.tk_popup(ev.x_root, ev.y_root)

    # 互称面板变量(上方已创建, 此处只挂联动)
    a_var.trace_add('write', lambda *_: compute_kin())
    b_var.trace_add('write', lambda *_: compute_kin())
    compute_kin()

    lst.bind('<<ListboxSelect>>', on_select)
    lst.bind('<Double-Button-1>', on_double)
    lst.bind('<Button-3>', on_right_click)
    lst.bind('<Delete>', lambda e: on_delete())

    # 快捷键
    root.bind_all('<Control-s>', lambda e: do_save() or 'break')
    root.bind_all('<Control-S>', lambda e: do_save() or 'break')
    root.bind_all('<Control-n>', lambda e: on_add())
    root.bind_all('<Control-e>', lambda e: on_edit())

    def on_close():
        if confirm_discard():
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', on_close)

    # ---- 启动 ----
    root.title(app_title())
    refresh_list()
    show_full()
    if state['readonly']:
        txt_set('当前显示的是内置示例族谱（只读）。\n\n'
                '要编辑自己的族谱：\n'
                '1. 文件 -> 新建族谱，从零建立；或\n'
                '2. 文件 -> 打开族谱，载入别人发你的 .json 文件。\n\n'
                '任何修改后点“保存”会提示另存为一份属于你的族谱文件。\n\n'
                '—————— 全谱 ——————\n\n'
                + '\n'.join(full_tree_lines(
                    sorted(roots, key=lambda p: (p.gen if p.gen is not None else 999,
                                                 p.name)))))
    root.mainloop()


def run_test():
    persons, roots, path, ro = load_persons()
    print(f'解析完成: 共 {len(persons)} 人, 根节点 {len(roots)} 个')
    for r in roots:
        print('根:', r.name, 'gen=', r.gen)
    for p in persons[:5]:
        print('-' * 40)
        print(person_detail(p))
    print('...')
    print(f'\n当前数据文件: {path or "(内置示例/未保存)"}')
    if not os.path.exists(DATA_PATH) and not getattr(sys, 'frozen', False):
        export_json()
        print(f'已生成可编辑文件: {DATA_PATH_WRITABLE}')


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if '--regen' in sys.argv:
        try:
            persons, _ = export_json(force_doc=True)
            print('已重新解析并生成 genealogy.json（%d 人）' % len(persons))
        except Exception as e:
            print('重新生成失败:', e)
    if '--test' in sys.argv:
        run_test()
    else:
        if args and os.path.exists(args[0]):
            set_last_file(os.path.abspath(args[0]))   # 双击 json / 命令行指定文件
        try:
            run_gui()
        except ImportError:
            print('当前环境无 tkinter，已退回文本模式:')
            run_test()
