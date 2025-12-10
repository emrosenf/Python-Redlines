import subprocess
import os
import tarfile
import zipfile


def get_version():
    """
    Extracts the version from the specified __about__.py file.
    """
    about = {}
    with open('./src/python_redlines/__about__.py') as f:
        exec(f.read(), about)
    return about['__version__']


def run_command(command):
    """
    Runs a shell command and prints its output.
    """
    process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if process.stdout is None:
        return
    for line in process.stdout:
        print(line.decode().strip())


def compress_files(source_dir, target_file, arcname=None):
    """
    Compresses files in the specified directory into a tar.gz or zip file.
    """
    arcname = arcname or os.path.basename(source_dir)
    if target_file.endswith('.tar.gz'):
        with tarfile.open(target_file, "w:gz") as tar:
            tar.add(source_dir, arcname=arcname)
    elif target_file.endswith('.zip'):
        with zipfile.ZipFile(target_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    zipf.write(os.path.join(root, file),
                               os.path.relpath(os.path.join(root, file),
                                               os.path.join(source_dir, '..')))


def ensure_submodules():
    if os.path.isdir('./Open-Xml-PowerTools/OpenXmlPowerTools'):
        return
    if not os.path.isdir('.git'):
        return
    print("Initializing submodules...")
    run_command('git submodule update --init --recursive')


def cleanup_dist(dist_dir):
    """Remove all build artifacts except .gitignore."""
    if not os.path.isdir(dist_dir):
        return
    for file in os.listdir(dist_dir):
        if file == '.gitignore':
            continue
        file_path = os.path.join(dist_dir, file)
        os.remove(file_path)
        print(f"Deleted old build file: {file}")


def main():
    ensure_submodules()
    version = get_version()
    print(f"Version: {version}")

    dist_dir = "./src/python_redlines/dist/"
    cleanup_dist(dist_dir)

    rid = 'linux-arm64'
    print(f"Building for {rid}...")
    run_command(f'dotnet publish ./csproj -c Release -r {rid} --self-contained')

    publish_dir = f'./csproj/bin/Release/net8.0/{rid}/publish'
    compress_files(publish_dir, f"{dist_dir}/{rid}-{version}.tar.gz", arcname=rid)

    # Clean build artifacts to keep tree small
    for path in ["./csproj/bin", "./csproj/obj"]:
        if os.path.exists(path):
            run_command(f'rm -rf {path}')

    print("Build and compression complete.")


if __name__ == "__main__":
    main()
