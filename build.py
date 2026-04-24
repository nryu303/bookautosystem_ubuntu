"""BookAutoSystem ビルドスクリプト

PyInstaller でビルドし、config・data・Playwrightブラウザを
出力フォルダにコピーする。

使い方:
  python build.py
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(ROOT_DIR, "dist", "BookAutoSystem")

# Windows端末のエンコード問題を回避（Linux/macOSでは既にUTF-8）
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass


def _playwright_cache_dir() -> str:
    """Playwrightブラウザキャッシュのデフォルトパスをプラットフォーム別に返す。"""
    env = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if env:
        return env
    if sys.platform == "win32":
        return os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")
    if sys.platform == "darwin":
        return str(Path.home() / "Library" / "Caches" / "ms-playwright")
    return str(Path.home() / ".cache" / "ms-playwright")


def step(msg):
    print(f"\n{'='*60}")
    print(f"  {msg}")
    print(f"{'='*60}")


def main():
    step("1. PyInstaller ビルド実行")
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "BookAutoSystem.spec", "--noconfirm"],
        cwd=ROOT_DIR,
    )
    if result.returncode != 0:
        print("ERROR: PyInstaller ビルド失敗")
        sys.exit(1)

    step("2. config/ フォルダをコピー")
    src_config = os.path.join(ROOT_DIR, "config")
    dst_config = os.path.join(DIST_DIR, "config")
    if os.path.exists(dst_config):
        shutil.rmtree(dst_config)
    shutil.copytree(src_config, dst_config)
    print(f"  → {dst_config}")

    step("3. data/ フォルダ構成を作成")
    for subdir in ["db", "screenshots", "html_snapshots", "csv", "output"]:
        path = os.path.join(DIST_DIR, "data", subdir)
        os.makedirs(path, exist_ok=True)
        print(f"  → {path}")

    # 自家在庫ファイル / URL指定型ファイルのサンプルをコピー
    for fname in ("self_stock.xlsx", "self_stock.csv",
                  "url_specified.xlsx", "url_specified.tsv"):
        src = os.path.join(ROOT_DIR, "data", "csv", fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(DIST_DIR, "data", "csv", fname))
            print(f"  → {fname} コピー")

    step("4. logs/ フォルダを作成")
    os.makedirs(os.path.join(DIST_DIR, "logs"), exist_ok=True)

    step("5. Playwright ブラウザをコピー")
    pw_src = _playwright_cache_dir()
    pw_dst = os.path.join(DIST_DIR, "playwright-browsers")

    if os.path.exists(pw_src):
        if os.path.exists(pw_dst):
            shutil.rmtree(pw_dst)
        print(f"  コピー中... (これには数分かかることがあります)")
        shutil.copytree(pw_src, pw_dst)
        # サイズ確認
        total_size = sum(
            os.path.getsize(os.path.join(dp, f))
            for dp, _, fns in os.walk(pw_dst)
            for f in fns
        )
        print(f"  → {pw_dst} ({total_size // (1024*1024)} MB)")
    else:
        print(f"  WARNING: Playwright ブラウザが見つかりません: {pw_src}")
        print(f"  実行前に playwright install chromium を実行してください")

    step("6. launcher.py をコピー")
    launcher_src = os.path.join(ROOT_DIR, "launcher.py")
    if os.path.exists(launcher_src):
        shutil.copy2(launcher_src, os.path.join(DIST_DIR, "launcher.py"))

    step("ビルド完了")
    exe_name = "BookAutoSystem.exe" if sys.platform == "win32" else "./BookAutoSystem"
    dist_rel = os.path.join("dist", "BookAutoSystem")
    print(f"""
  出力先: {DIST_DIR}

  実行方法:
    cd {dist_rel}
    {exe_name} --once     (1回実行)
    {exe_name} --web      (管理画面のみ)
    {exe_name}            (常駐モード)

  配布方法:
    {dist_rel} フォルダを圧縮して配布
    受取側はフォルダを展開して {exe_name} を実行するだけ
""")


if __name__ == "__main__":
    main()
