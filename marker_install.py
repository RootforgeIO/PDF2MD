"""Install the optional Marker engine for the Windows setup build."""
from __future__ import annotations

from i18n import tr, get_language

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def installed_build() -> bool:
    return Path(__file__).with_name("installed.flag").is_file()


def user_data_dir() -> Path:
    return Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "PDF2MD"


def default_download_dir() -> Path:
    return Path(os.environ.get("ProgramData", r"C:\ProgramData")) / "PDF2MD"


def marker_command(root: Path, device: str) -> list[str]:
    return [
        str(Path(sys.executable).with_name("python.exe")), "-E", "-s", "-B", "-u", "-X", "utf8",
        str(Path(__file__).with_name("marker_single_portable.py")),
        "--managed-install", str(root), device,
    ]


def installed_marker_command() -> list[str] | None:
    try:
        settings = json.loads((user_data_dir() / "marker.json").read_text(encoding="utf-8"))
        root, device = Path(settings["directory"]), settings["device"]
        if not root.is_absolute() or device not in ("cpu", "cuda"):
            return None
        site = root / "site"
        # Marker uses a namespace package and therefore has no top-level __init__.py.
        if (site / "marker/scripts/convert_single.py").is_file() and (site / "torch/__init__.py").is_file():
            return marker_command(root, device)
    except (OSError, ValueError, KeyError, TypeError):
        pass
    return None


def _pip_command(root: Path, device: str) -> list[str]:
    build = "cu126" if device == "cuda" else "cpu"
    wheel_base = f"https://download.pytorch.org/whl/{build}"
    return [
        str(Path(sys.executable).with_name("python.exe")), "-I", "-B", "-X", "utf8", "-m", "pip", "install",
        "--index-url", "https://pypi.org/simple", "--disable-pip-version-check", "--no-input",
        "--no-cache-dir", "--no-compile", "--ignore-installed", "--only-binary=:all:",
        "--target", str(root / "site"),
        f"{wheel_base}/torch-2.7.1%2B{build}-cp312-cp312-win_amd64.whl",
        f"{wheel_base}/torchvision-0.22.1%2B{build}-cp312-cp312-win_amd64.whl",
        "marker-pdf==1.10.2", "surya-ocr==0.17.1", "transformers==4.57.6", "pdftext==0.6.3",
    ]


def _run(command: list[str], temp: Path, cancel_event) -> None:
    from pdf2md import ConversionCancelled, _kill_process_tree

    if cancel_event.is_set():
        raise ConversionCancelled()
    env = {key: value for key, value in os.environ.items() if not key.upper().startswith("PIP_")}
    env.update(PIP_CONFIG_FILE=os.devnull, TEMP=str(temp), TMP=str(temp), PDF2MD_LANGUAGE=get_language())
    # A file drains output even while pip or the model downloader is busy. Only errors reach the GUI.
    with tempfile.TemporaryFile(dir=temp) as output:
        with subprocess.Popen(
            command, cwd=temp, env=env, stdout=output, stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW,
        ) as proc:
            while proc.poll() is None:
                if cancel_event.wait(0.2):
                    _kill_process_tree(proc)
                    raise ConversionCancelled()
            if proc.returncode:
                length = output.seek(0, os.SEEK_END)
                output.seek(max(0, length - 3500))
                detail = output.read().decode("utf-8", errors="replace").strip()
                raise RuntimeError(detail or tr('Установка Marker завершилась с кодом {v0}.').format(v0=proc.returncode))


def _register_marker(directory: Path, device: str) -> None:
    state = user_data_dir() / "marker.json"
    fd, name = tempfile.mkstemp(dir=state.parent)
    pending = Path(name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump({"directory": str(directory), "device": device}, stream, ensure_ascii=False)
        pending.replace(state)
    finally:
        pending.unlink(missing_ok=True)


def install_marker(folder: Path, device: str, cancel_event, progress) -> None:
    import msvcrt

    if device not in ("cpu", "cuda"):
        raise ValueError(tr('Выберите CPU или NVIDIA CUDA.'))
    data = user_data_dir()
    data.mkdir(parents=True, exist_ok=True)
    with (data / "marker-install.lock").open("a+b") as lock:
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise RuntimeError(tr('Marker уже устанавливается в другом окне PDF2MD.')) from exc
        if installed_marker_command() is not None:
            return

        folder = folder.resolve()
        folder.mkdir(parents=True, exist_ok=True)
        # mkdtemp uses mode 0700. Python 3.12.4+ applies a private ACL even under ProgramData.
        directory = Path(tempfile.mkdtemp(prefix="Marker-", dir=folder))
        registered = False
        try:
            with tempfile.TemporaryDirectory(dir=directory) as temporary:
                temp = Path(temporary)
                progress(5, tr('Загружаю Marker и PyTorch…'))
                _run(_pip_command(directory, device), temp, cancel_event)
                progress(55, tr('Загружаю и проверяю модели…'))
                _run([*marker_command(directory, device), "--prepare-models"], temp, cancel_event)
            if cancel_event.is_set():
                from pdf2md import ConversionCancelled

                raise ConversionCancelled()
            _register_marker(directory, device)
            registered = True
        finally:
            if not registered and directory.resolve().parent == folder:
                shutil.rmtree(directory)
