import PyInstaller.__main__
from pyinstaller_versionfile import create_versionfile

# --- Step 1: Read SemVer from version.txt and convert to a Windows-compatible format ---
with open("version.txt", "r") as f:
    semver = f.read().strip()

# Windows executables require a four-part version (e.g., 1.2.3.0).
# We append '.0' to the SemVer string from your file to satisfy this requirement.
win_version = f"{semver}.0"

# --- Step 2: Generate the version file using the four-part version string ---
create_versionfile(
    "version_info.py",
    version=win_version,
    file_description="StocksRPA Excel Processor",
    internal_name="StocksRPA",
    original_filename="StocksRPA.exe",
    product_name="StocksRPA",
)

# --- Step 3: Define the PyInstaller arguments ---
pyinstaller_args = [
    'main.py',
    '--onefile',
    '--windowed',
    '--name=StocksRPA',
    '--add-data=version.txt;.',
    '--version-file=version_info.py',
]

# --- Step 4: Run PyInstaller ---
PyInstaller.__main__.run(pyinstaller_args)

print("\nBuild complete. The executable is in the 'dist' folder.")
