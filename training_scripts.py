import os
import pandas as pd
import logging
import pypsa
import numpy as np
import plotly.express as px
import shutil

pd.options.plotting.backend = "plotly"
import warnings
warnings.filterwarnings(action='ignore', message='Data Validation extension is not supported and will be removed')
pd.options.mode.chained_assignment = None  # Suppress the warning


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




def convert_selected_sheets_to_csv(excel_file_path, csv_folder_path):
    """
    Reads an Excel file, checks if any sheets match a predefined list, and converts them into CSV files.

    Parameters:
    excel_file_path (str): The file path of the Excel file.
    csv_folder_path (str): The directory where CSV files should be saved.

    Returns:
    list: List of paths to the created CSV files.
    """
    logging.basicConfig(level=logging.INFO)

    if os.path.exists(csv_folder_path):
        shutil.rmtree(csv_folder_path)

    # Recreate the folder
    os.makedirs(csv_folder_path)

    created_csv_files = []

    # Initialize network
    n = pypsa.Network()
    all_variables = ['snapshots']
    for key in n.all_components:
        list_name = n.components[key]['list_name']
        var_names = (n.component_attrs[key][n.component_attrs[key]["varying"]== True].index).to_list()
        all_variables.append(list_name)
        for var in var_names:
            all_variables.append(f"{list_name}-{var}")

    xls = None  # Initialize xls variable

    try:
        # Load Excel file
        xls = pd.ExcelFile(excel_file_path)

        # Iterate through sheets in the predefined list
        for sheet_name in xls.sheet_names:
            if sheet_name in all_variables:
                # Read sheet into DataFrame
                df = xls.parse(sheet_name)

                # Define the output CSV file path
                csv_file_path = os.path.join(csv_folder_path, f"{sheet_name}.csv")

                # Save DataFrame as CSV
                df.to_csv(csv_file_path, index=False)

                logging.info(f"Converted {sheet_name} to CSV.")
                created_csv_files.append(csv_file_path)

        logging.info(f"Conversion complete. CSV files are saved in '{csv_folder_path}'")
        return csv_folder_path

    except Exception as e:
        logging.error(f"Error processing Excel file: {e}")
        return []

    finally:
        # Explicitly close the Excel file
        if xls is not None:
            xls.close()
            logging.info("Excel file closed successfully.")

