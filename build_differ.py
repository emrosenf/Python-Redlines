import os
import shlex
import shutil
import subprocess
import tarfile
import zipfile
from pathlib import Path
from typing import Optional

ROOT = Path(__file__).resolve().parent
SRC_DIR = ROOT / "src" / "python_redlines"
DIST_DIR = SRC_DIR / "dist"
CSPROJ_DIR = ROOT / "csproj"
FRAMEWORK = "net10.0"
DEFAULT_RID = "linux-arm64"


def get_version() -> str:
    """Extract the package version from __about__.py."""
    about: dict[str, str] = {}
    with (SRC_DIR / "__about__.py").open() as f:
        exec(f.read(), about)
    return about["__version__"]


def run_command(command: str, cwd: Optional[Path] = None) -> int:
    """Run a shell command, stream output, and return the exit code."""
    print(f"Running: {command}")
    process = subprocess.Popen(
        command,
        shell=True,
        cwd=str(cwd) if cwd else None,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    if process.stdout is not None:
        for line in process.stdout:
            print(line.rstrip())
    return process.wait()


def compress_files(source_dir: Path, target_file: Path, arcname: Optional[str] = None) -> None:
    """Compress files in ``source_dir`` into ``target_file``."""
    arcname = arcname or source_dir.name
    if target_file.name.endswith(".tar.gz"):
        with tarfile.open(target_file, "w:gz") as tar:
            tar.add(str(source_dir), arcname=arcname)
    elif target_file.suffix == ".zip":
        with zipfile.ZipFile(target_file, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(source_dir):
                for file in files:
                    file_path = Path(root) / file
                    zipf.write(file_path, arcname=str(file_path.relative_to(source_dir.parent)))
    else:
        raise ValueError(f"Unsupported archive format: {target_file}")


def ensure_submodules() -> None:
    if (ROOT / "Open-Xml-PowerTools" / "OpenXmlPowerTools").is_dir():
        return
    if not (ROOT / ".git").is_dir():
        return
    print("Initializing submodules...")
    exit_code = run_command("git submodule update --init --recursive", cwd=ROOT)
    if exit_code != 0:
        raise SystemExit("Failed to initialize submodules")


def cleanup_dist(dist_dir: Path) -> None:
    """Remove all build artifacts except .gitignore."""
    if not dist_dir.exists():
        return
    for entry in dist_dir.iterdir():
        if entry.name == ".gitignore":
            continue
        if entry.is_dir():
            shutil.rmtree(entry)
        else:
            entry.unlink()
        print(f"Deleted old build file: {entry.name}")


def clean_build_artifacts() -> None:
    for path in [CSPROJ_DIR / "bin", CSPROJ_DIR / "obj"]:
        if path.exists():
            shutil.rmtree(path)
            print(f"Removed {path}")


def find_publish_dir(rid: str) -> Optional[Path]:
    base = CSPROJ_DIR / "bin" / "Release" / FRAMEWORK
    candidates: list[Path] = [
        base / rid / "publish",
        base / rid / "self-contained" / "publish",
    ]
    candidates.extend(path for path in base.glob("**/publish") if rid in str(path))
    for candidate in candidates:
        if candidate.is_dir():
            return candidate
    return None


def publish_binary(rid: str) -> Path:
    for self_contained in (True, False):
        sc_value = "true" if self_contained else "false"
        exit_code = run_command(
            f"dotnet publish {shlex.quote(str(CSPROJ_DIR))} -c Release -r {rid} --self-contained {sc_value}",
            cwd=ROOT,
        )
        if exit_code != 0:
            raise SystemExit(f"dotnet publish failed (self-contained={sc_value}); see output above.")
        publish_dir = find_publish_dir(rid)
        if publish_dir:
            print(f"Publish output: {publish_dir}")
            return publish_dir
        if self_contained:
            print("Publish directory missing for self-contained build; retrying without self-contained...")
    raise SystemExit(f"Publish directory missing after dotnet publish for rid={rid}.")


def main() -> None:
    ensure_submodules()
    version = get_version()
    rid = os.environ.get("REDLINES_RID", DEFAULT_RID)
    print(f"Version: {version}")
    print(f"Runtime identifier: {rid}")

    cleanup_dist(DIST_DIR)
    publish_dir = publish_binary(rid)

    DIST_DIR.mkdir(parents=True, exist_ok=True)
    archive_path = DIST_DIR / f"{rid}-{version}.tar.gz"
    compress_files(publish_dir, archive_path, arcname=rid)

    clean_build_artifacts()
    print("Build and compression complete.")


if __name__ == "__main__":
    main()

