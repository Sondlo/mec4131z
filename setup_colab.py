import sys
import subprocess
import pkg_resources
import os

# List of required packages
REQUIRED_PACKAGES = [
    "pypsa",
    "numpy",
    "pandas",
    "matplotlib",
    "scipy",
    "seaborn",
    "networkx",
    "tqdm",
    "pyomo",
    "cvxpy",
    "cplex",
    "gurobipy",
    "openpyxl",
    "xlrd"
]

def is_colab():
    """Check if the script is running inside Google Colab."""
    try:
        import google.colab
        return True
    except ImportError:
        return False

def is_installed(package_name):
    """Check if a package is already installed."""
    try:
        pkg_resources.get_distribution(package_name)
        return True
    except pkg_resources.DistributionNotFound:
        return False

def install_colab_dependencies():
    """Install required packages only if running in Google Colab and not already installed."""
    if not is_colab():
        print("Not running in Google Colab. No installation needed.")
        return

    # Create a flag file to check if the script was already run
    flag_file = "/content/.colab_packages_installed"

    if os.path.exists(flag_file):
        print("Dependencies already installed. Skipping installation.")
        return
    
    print("Installing required packages in Google Colab...")

    for package in REQUIRED_PACKAGES:
        if not is_installed(package):
            subprocess.call([sys.executable, "-m", "pip", "install", package])

    # Create the flag file after successful installation
    with open(flag_file, "w") as f:
        f.write("installed")

    print("Installation complete.")

    drive.mount('/content/drive')
    folder_path = f'/content/drive/My Drive/{folder_name}'
    if not os.path.exists(folder_path): os.makedirs(folder_path)
    os.chdir(folder_path)

    markdown_content = f"""
    Folder '{folder_path}' is ready to use. <br><br>

    Open Google Drive and upload the excel model file to:

    [{folder_name}](https://drive.google.com/drive/my-drive)

    (this link will open in a new tab)
    """

    # Display the markdown content
    display(Markdown(markdown_content))