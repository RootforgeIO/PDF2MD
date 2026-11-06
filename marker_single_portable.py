from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from i18n import tr, set_language


def _configure_cache(models_dir: Path, *, managed: bool = False) -> None:
    datalab_models_dir = models_dir / "datalab" / "models"
    models_dir.mkdir(parents=True, exist_ok=True)
    datalab_models_dir.mkdir(parents=True, exist_ok=True)

    defaults = {
        "MODEL_CACHE_DIR": datalab_models_dir,
        "HF_HOME": models_dir / "huggingface",
        "HUGGINGFACE_HUB_CACHE": models_dir / "huggingface" / "hub",
        "TRANSFORMERS_CACHE": models_dir / "huggingface" / "transformers",
        "TORCH_HOME": models_dir / "torch",
        "XDG_CACHE_HOME": models_dir / "cache",
    }
    for key, value in defaults.items():
        if managed:
            os.environ[key] = str(value)
        else:
            os.environ.setdefault(key, str(value))


def _patch_surya_downloader() -> None:
    import surya.common.s3 as s3

    original_download_directory = s3.download_directory

    def download_directory(remote_path: str, local_dir: str):
        local_path = Path(local_dir)
        if local_path.exists() and not s3.check_manifest(local_dir):
            shutil.rmtree(local_path)
        local_path.mkdir(parents=True, exist_ok=True)
        return original_download_directory(remote_path, local_dir)

    s3.download_directory = download_directory


def main() -> int:
    set_language("en" if os.environ.get("PDF2MD_LANGUAGE") == "en" else "ru")
    managed = sys.argv[1:2] == ["--managed-install"]
    if managed:
        root, device = Path(sys.argv[2]), sys.argv[3]
        del sys.argv[1:4]
        sys.path.insert(0, str(root / "site"))
        os.chdir(root)
        os.environ.update(TORCH_DEVICE=device, S3_BASE_URL="https://models.datalab.to")
        _configure_cache(root / "models", managed=True)
    else:
        root = os.environ.get("PDF2MD_PORTABLE_ROOT")
        if root:
            _configure_cache(Path(root) / "models")
    _patch_surya_downloader()
    if managed and sys.argv[1:] == ["--prepare-models"]:
        import torch
        from marker.models import create_model_dict

        if device == "cuda" and not torch.cuda.is_available():
            raise RuntimeError(tr("CUDA недоступна. Проверьте драйвер NVIDIA или установите вариант CPU."))
        create_model_dict()
        return 0
    from marker.scripts.convert_single import convert_single_cli

    convert_single_cli(prog_name="marker_single")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
