"""Create the local environment and repair incomplete pip from Python's bundled wheel."""
from __future__ import annotations
import os
from pathlib import Path
import subprocess
import sys
import venv

PIP_PROBE = 'import pip._internal.cli.main; import pip._internal.utils'


def pip_is_working(python: Path) -> bool:
    for args in [('-m', 'pip', '--version'), ('-c', PIP_PROBE)]:
        result = subprocess.run([str(python), '-I', *args], capture_output=True, text=True)
        if result.returncode:
            return False
    return True


def ensure_pip(python: Path, report=print) -> bool:
    if pip_is_working(python):
        return False
    report('检测到 pip 缺失或文件不完整，正在使用 Python 自带安装包离线修复……')
    try:
        import ensurepip
        bundled = Path(ensurepip.__file__).parent / '_bundled'
        wheel = next(bundled.glob('pip-*.whl'))
    except (ImportError, StopIteration) as exc:
        raise RuntimeError('当前 Python 不含 pip 离线安装包，请修复或重新安装 Python 官方完整版后重试。') from exc
    # Running pip directly from its trusted bundled wheel avoids importing damaged files.
    # Force replacement even when metadata says the same pip version is installed.
    code = (
        "import runpy,sys; sys.path.insert(0,sys.argv[1]); "
        "sys.argv=['pip','install','--no-index','--no-deps','--force-reinstall',sys.argv[1]]; "
        "runpy.run_module('pip',run_name='__main__')"
    )
    result = subprocess.run([str(python), '-I', '-c', code, str(wheel)], capture_output=True, text=True)
    if result.returncode or not pip_is_working(python):
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError('pip 离线修复失败。请检查文件权限、磁盘空间或安全软件拦截记录。\n' + detail)
    report('pip 离线修复成功。')
    return True


def prepare_environment(root: Path) -> Path:
    env = root / '.venv'
    python = env / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')
    if not python.exists():
        print('正在创建程序专用 Python 环境……', flush=True)
        venv.EnvBuilder(with_pip=False).create(env)
    ensure_pip(python)
    result = subprocess.run([str(python), '-I', '-c', 'import docx, openpyxl'], capture_output=True)
    if result.returncode:
        print('正在安装程序依赖；此步骤需要联网……', flush=True)
        result = subprocess.run([str(python), '-I', '-m', 'pip', 'install', '-r', str(root / 'requirements.txt')], cwd=root)
        if result.returncode:
            raise RuntimeError('依赖安装失败，请查看上方具体错误。网络、证书、写入权限或安装包异常均可能导致失败。')
    return python


def main() -> int:
    if sys.version_info < (3, 10):
        print('需要 Python 3.10 或以上版本，请升级 Python 后重试。')
        return 1
    root = Path(__file__).resolve().parent
    try:
        python = prepare_environment(root)
        return subprocess.call([str(python), str(root / 'app.py'), *sys.argv[1:]], cwd=root)
    except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
        print('\n启动失败：' + str(exc), file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        return 130


if __name__ == '__main__':
    raise SystemExit(main())
