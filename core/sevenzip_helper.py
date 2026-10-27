import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import List, Optional, Tuple


DEFAULT_7Z_PATHS = [
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe",
    r"D:\Program Files\7-Zip\7z.exe",
    r"D:\7-Zip\7z.exe",
]


def find_7z_executable(custom_path: str = "") -> Optional[str]:
    """
    自动查找 7z.exe 路径：
    1. 优先使用用户在界面指定的路径
    2. 检查 PyInstaller 打包内置的 7z_bin/7z.exe (_MEIPASS) 或程序同级目录
    3. 检查系统常见安装位置与 PATH 环境变量
    """
    if custom_path and os.path.isfile(custom_path):
        return str(Path(custom_path).resolve())

    # 检查 PyInstaller 单文件打包解压目录 (_MEIPASS)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled_7z = Path(meipass) / "7z_bin" / "7z.exe"
        if bundled_7z.is_file():
            return str(bundled_7z)

    # 检查当前程序所在目录
    app_dir = (
        Path(sys.executable).resolve().parent
        if getattr(sys, "frozen", False)
        else Path(__file__).resolve().parent.parent
    )
    for rel in ("7z_bin/7z.exe", "7z.exe"):
        local_7z = app_dir / rel
        if local_7z.is_file():
            return str(local_7z)

    for p in DEFAULT_7Z_PATHS:
        if os.path.isfile(p):
            return p

    which_7z = shutil.which("7z") or shutil.which("7za")
    if which_7z and os.path.isfile(which_7z):
        return which_7z

    return None


def create_encrypted_7z(
    input_files: List[Path],
    output_7z_path: Path,
    password: str,
    sevenzip_exe: str = "",
    compression_level: int = 1,
) -> Tuple[bool, str, Path]:
    """
    调用 7z.exe 将一个或多个本地文件打包为 AES-256 加密的 .7z 压缩包，
    并强制开启 -mhe=on（加密文件名与目录头，未输入密码无法看到内部文件名）。
    """
    exe_path = find_7z_executable(sevenzip_exe)
    if not exe_path:
        return (
            False,
            "未找到 7-Zip (7z.exe) 程序！请确认已安装 7-Zip 或在界面上指定 7z.exe 路径。",
            output_7z_path,
        )

    if not password:
        return False, "开启 7z 加密时必须设置加密密码！", output_7z_path

    if not input_files:
        return False, "没有提供待加密的源文件！", output_7z_path

    output_7z_path = Path(output_7z_path).resolve()
    output_7z_path.parent.mkdir(parents=True, exist_ok=True)

    if output_7z_path.exists():
        try:
            output_7z_path.unlink()
        except Exception as e:
            return (
                False,
                f"无法清理旧的临时 7z 文件 ({output_7z_path}): {e}",
                output_7z_path,
            )

    cmd = [
        exe_path,
        "a",
        "-t7z",
        f"-mx={int(compression_level)}",
        "-mhe=on",
        f"-p{password}",
        "-y",
        str(output_7z_path),
    ] + [str(Path(f).resolve()) for f in input_files]

    creation_flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creation_flags,
            timeout=3600,
        )
        if proc.returncode != 0:
            err_detail = (proc.stderr or proc.stdout or "").strip()
            return (
                False,
                f"7z 加密打包失败 (返回码 {proc.returncode}): {err_detail[:300]}",
                output_7z_path,
            )

        if not output_7z_path.is_file() or output_7z_path.stat().st_size == 0:
            return (
                False,
                "7z 命令已执行，但未生成有效的 .7z 压缩文件！",
                output_7z_path,
            )

        return (
            True,
            "7z 加密成功（AES-256 数据加密 + -mhe=on 文件名加密）",
            output_7z_path,
        )

    except subprocess.TimeoutExpired:
        return False, "7z 压缩加密超时（超过 3600 秒）！", output_7z_path
    except Exception as e:
        return False, f"调用 7z.exe 发生异常: {str(e)}", output_7z_path
