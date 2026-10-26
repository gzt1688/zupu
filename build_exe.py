"""
构建脚本 v4 - 简化为仅 PyInstaller + 缓存清理
"""
import subprocess
import sys
import os
import shutil
import struct
import io
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICO_PATH = os.path.join(BASE_DIR, "app_icon.ico")
SRC_ICO = os.path.join(BASE_DIR, "favicon.ico")
if not os.path.exists(SRC_ICO):
    SRC_ICO = os.path.join(BASE_DIR, "app_icon.ico")   # favicon 缺失时直接以现有 app_icon 重建
# 源图标缺失(如 favicon.ico 已被改名/删除)时, 回退到现有 app_icon.ico, 保证打包不因图标源缺失而失败
if not os.path.exists(SRC_ICO):
    SRC_ICO = ICO_PATH
PY_SCRIPT = os.path.join(BASE_DIR, "族谱查看器.py")
DIST_DIR = os.path.join(BASE_DIR, "dist_new")

SIZES = [16, 32, 48, 64, 128, 256]

def build_ico():
    """手动构建多分辨率 ICO"""
    from PIL import Image
    
    print("[1/3] 构建多分辨率 ICO...")
    src = Image.open(SRC_ICO).convert("RGBA")
    if src.size != (256, 256):
        src = src.resize((256, 256), Image.LANCZOS)
    
    png_list = []
    for s in SIZES:
        img = src.copy() if s == 256 else src.resize((s, s), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        png_list.append(buf.getvalue())
    
    num = len(SIZES)
    buf = bytearray()
    buf += struct.pack("<HHH", 0, 1, num)
    
    data_offset = 6 + num * 16
    entries = bytearray()
    image_data = bytearray()
    
    for s, png_bytes in zip(SIZES, png_list):
        w = 0 if s >= 256 else s
        h = 0 if s >= 256 else s
        entries += struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, len(png_bytes), data_offset)
        image_data += png_bytes
        data_offset += len(png_bytes)
    
    buf += entries + image_data
    with open(ICO_PATH, "wb") as f:
        f.write(bytes(buf))
    
    print(f"  {ICO_PATH}: {os.path.getsize(ICO_PATH):,} bytes, {num} 帧 ({' '.join(str(s) for s in SIZES)})")

def build_exe():
    """PyInstaller 打包"""
    print("\n[2/3] PyInstaller 打包...")
    
    # 清理
    for d in [DIST_DIR, os.path.join(BASE_DIR, "build")]:
        if os.path.exists(d):
            shutil.rmtree(d)
    spec = os.path.join(BASE_DIR, "族谱查看器.spec")
    if os.path.exists(spec):
        os.remove(spec)
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile", "--windowed",
        "--name", "族谱查看器",
        f"--icon={ICO_PATH}",
        f"--distpath={DIST_DIR}",
        # 打包内置示例(明朝皇帝世系), 无历史文件/数据时作为只读示例展示
        "--add-data", f"明朝族谱.json;.",
        "--add-data", f"app_icon.ico;.",
        # 照片功能需要 Pillow(编辑窗/主窗缩略图)
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageTk",
        PY_SCRIPT,
        "--clean", "--noupx",
    ]
    
    result = subprocess.run(cmd, cwd=BASE_DIR, capture_output=True, text=True, timeout=600)
    if result.returncode != 0:
        print("  PyInstaller 失败!")
        print(result.stderr[-2000:])
        sys.exit(1)
    
    exe = os.path.join(DIST_DIR, "族谱查看器.exe")
    sz = os.path.getsize(exe)
    print(f"  打包完成: {exe} ({sz:,} bytes)")
    
    # 验证 EXE 中是否有 PNG 图标数据
    with open(exe, "rb") as f:
        data = f.read()
    png_count = data.count(b'\x89PNG')
    print(f"  EXE 中 PNG 数据块: {png_count} 个")
    return sz

def clear_cache():
    """彻底清除 Windows 图标缓存"""
    print("\n[3/3] 清除 Windows 图标缓存...")
    print("  注意: 此操作会重启 Explorer!")
    
    # 1. 结束 Explorer
    subprocess.run(["taskkill", "/f", "/im", "explorer.exe"],
                   capture_output=True, timeout=15)
    time.sleep(2)
    
    # 2. 删除图标缓存文件
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    
    # IconCache.db
    icon_db = os.path.join(local_appdata, "IconCache.db")
    if os.path.exists(icon_db):
        try:
            os.remove(icon_db)
            print(f"  已删除 IconCache.db")
        except Exception as e:
            print(f"  删除 IconCache.db 失败: {e}")
    
    # thumbcache_*.db
    thumb_dir = os.path.join(local_appdata, "Microsoft", "Windows", "Explorer")
    if os.path.exists(thumb_dir):
        for fn in os.listdir(thumb_dir):
            if fn.startswith("thumbcache_") or fn.startswith("iconcache_"):
                try:
                    os.remove(os.path.join(thumb_dir, fn))
                except:
                    pass
        print("  已清理缩略图缓存")
    
    # 3. 用 ie4uinit 刷新
    try:
        subprocess.run(["ie4uinit.exe", "-ClearIconCache"], capture_output=True, timeout=10)
        print("  执行 ie4uinit -ClearIconCache")
    except:
        pass
    
    # 4. 重启 Explorer
    time.sleep(1)
    subprocess.Popen("explorer.exe")
    time.sleep(3)
    print("  已重启 Explorer")

def main():
    print("=" * 60)
    print("  族谱查看器 - 构建脚本 v4")
    print("=" * 60)

    build_ico()
    build_exe()
    print("\n完成! EXE 位置: %s" % os.path.join(DIST_DIR, "族谱查看器.exe"))

if __name__ == "__main__":
    main()
