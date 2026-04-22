import subprocess
import sys
import os

try:
    # Python 3.8+
    from importlib.metadata import distributions
except ImportError:
    # Fallback (very old Python)
    print("❌ Python version too old. Please use Python 3.8+")
    sys.exit(1)


def get_installed_packages():
    """Return a set of installed package names (lowercase)."""
    installed = set()
    for dist in distributions():
        try:
            name = dist.metadata["Name"]
            if name:
                installed.add(name.lower())
        except:
            pass
    return installed


def check_and_install(file_path):
    if not os.path.exists(file_path):
        print(f"❌ Error: {file_path} not found.")
        return

    installed_packages = get_installed_packages()

    with open(file_path, 'r') as f:
        modules = [line.strip() for line in f if line.strip() and not line.startswith('#')]

    for module in modules:
        # Extract clean package name
        clean_name = (
            module.split('==')[0]
            .split('>=')[0]
            .split('<=')[0]
            .split('>')[0]
            .split('<')[0]
            .strip()
            .lower()
        )

        if clean_name in installed_packages:
            print(f"✅ {module} is already installed.")
        else:
            print(f"⚠️ {module} not found. Installing...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", module])
                print(f"🚀 Successfully installed {module}.")
            except Exception as e:
                print(f"❌ Failed to install {module}: {e}")


if __name__ == "__main__":
    check_and_install('requirements.txt')