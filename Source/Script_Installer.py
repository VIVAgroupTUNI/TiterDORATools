"""
A script designed to automate the installation and updating of Python packages, 
specifically "pip", "xlrd", and "openpyxl". It provides a structured and user-friendly 
output to guide users through the process, including error handling and status updates.

The script performs the following actions:
1. Ensures "pip" is installed, and if not, installs it.
2. Updates "pip" to the latest version.
3. Ensures "xlrd" (a library for reading .xls files) is installed, and if not, installs it.
4. Updates "xlrd" to the latest version.
5. Ensures "openpyxl" (a library for reading and writing .xlsx files) is installed, and if not, installs it.
6. Updates "openpyxl" to the latest version.

The script uses various helper functions to print status messages, actions, tasks, progress, success, and errors.
It also includes exit functions to gracefully handle errors or user interruptions.

Usage:
    python Script_Installer.py

Prerequisites:
    - Python must be installed on the system.
    - Internet access is required to install and update packages.

Example:
    python Script_Installer.py
"""




########## IMPORTS ##########


import sys
import time
import subprocess




########## ATTRIBUTES ##########

SCRIPT_VERSION = "0.1.0"    # The public version number of the script

action_iterator = 1
DELAY = 0.05




########## EXIT ROUTES ##########


def exit_by_error(message):
    """Exits the script with a given error message."""
    try:
        time.sleep(DELAY)
        print_error(message)
        time.sleep(DELAY)
        print_status("Failed")
    except KeyboardInterrupt:
        exit_by_interruption()
    exit("")


def exit_by_interruption():
    """Exits the script due to an interruption (e.g., KeyboardInterrupt, CTRL+C)."""
    try:
        time.sleep(DELAY)
        print()
        time.sleep(DELAY)
        print_status("Stopped")
    except KeyboardInterrupt:
        exit_by_interruption()
    exit("")




########## PRINT FUNCTIONS ##########


def print_status(status_string, leading_line_break=True, tailing_line_break=False):
    """Prints a status message with version in the top border."""
    try:
        time.sleep(DELAY)
        text = f"SCRIPT INSTALLER {status_string.upper()}"
        # Middle line padding
        asterisks = "*" * 3
        spaces = " " * 5
        middle_line = f"{asterisks}{spaces}{text}{spaces}{asterisks}"
        # Top border with version
        top_border_text = f" v{SCRIPT_VERSION} "
        total_width = len(middle_line)
        top_line = "*" * ((total_width - len(top_border_text)) // 2) + top_border_text
        if len(top_line) < total_width:
            top_line += "*" * (total_width - len(top_line))
        bottom_line = "*" * len(middle_line)

        if leading_line_break:
            time.sleep(DELAY)
            print()
        time.sleep(DELAY)
        print(top_line)
        time.sleep(DELAY)
        print(middle_line)
        time.sleep(DELAY)
        print(bottom_line)
        if tailing_line_break or status_string.strip().lower() == "completed":
            time.sleep(DELAY)
            print()
    except KeyboardInterrupt:
        exit_by_interruption()


def print_action(action_str):
    """Prints a separator line for actions."""
    try:
        global action_iterator
        time.sleep(DELAY)
        print()
        time.sleep(DELAY)
        print(f"========== Action {action_iterator}: {action_str.title()} ==========")
        action_iterator += 1
    except KeyboardInterrupt:
        exit_by_interruption()


def print_task(task_str):
    """Prints a task message."""
    try:
        time.sleep(DELAY)
        print(f"> Task: {task_str}..")
    except KeyboardInterrupt:
        exit_by_interruption()


def print_progress(progress_str):
    """Prints a progress message."""
    try:
        time.sleep(DELAY)
        print(f"> Progress: {progress_str}")
    except KeyboardInterrupt:
        exit_by_interruption()


def print_success(success_str):
    """Prints a success message."""
    try:
        time.sleep(DELAY)
        print(f"> Success: {success_str}")
    except KeyboardInterrupt:
        exit_by_interruption()


def print_error(error_str):
    """Prints an error message."""
    try:
        time.sleep(DELAY)
        print(f"> Error: {error_str}")
    except KeyboardInterrupt:
        exit_by_interruption()




########## SCRIPT INITIALIZATION ##########


def ensure_pip_installed():
    """Ensure pip is installed."""
    try:
        import pip
        print_success('"pip" is already installed.')
    except ImportError:
        print_progress('Could not find "pip", attempting to install using ensurepip.')
        try:
            subprocess.check_call([sys.executable, '-m', 'ensurepip', '--default-pip'])
            print_success('"pip" installed successfully.')
        except Exception as e:
            exit_by_error(f'Failed to install "pip".  > REASON: {e}.')


def ensure_library_installed(library_name: str):
    """Ensure a given library is installed."""
    try:
        __import__(library_name)
        print_success(f'"{library_name}" is already installed.')
    except ImportError:
        print_progress(f'Could not find "{library_name}", attempting to install.')
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", library_name])
            print_success(f'"{library_name}" installed successfully.')
        except Exception as e:
            print_error(f'Failed to install "{library_name}".  > REASON: {e}.')


def update_library(library_name: str):
    """Update a given library to the latest version."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", library_name])
        print_success(f'"{library_name}" version is up to date.')
    except Exception as e:
        print_error(f'Failed to update "{library_name}".  > REASON: {e}.')




########## ACTUAL SCRIPT ##########


def run_script():
    print_action('Install PIP')
    print_task('Ensuring that the main python installer tool "pip" is installed.')
    ensure_pip_installed()
    # PIP HAS BEEN INSTALLED

    print_action('Update PIP')
    print_task('Attempting to update "pip".')
    update_library("pip")
    # PIP HAS BEEN UPDATED

    print_action('Install XLRD')
    print_task('Ensuring that the .xls file reader library "xlrd" is installed.')
    ensure_library_installed("xlrd")
    # XLRD HAS BEEN INSTALLED

    print_action('Update XLRD')
    print_task('Attempting to update "xlrd".')
    update_library("xlrd")
    # XLRD HAS BEEN UPDATED

    print_action('Install OPENPYXL')
    print_task('Ensuring that the .xlsx file reader library "openpyxl" is installed.')
    ensure_library_installed("openpyxl")
    # OPENPYXL HAS BEEN INSTALLED

    print_action('Update OPENPYXL')
    print_task('Attempting to update "openpyxl".')
    update_library("openpyxl")
    # OPENPYXL HAS BEEN UPDATED

    print_action('Install NUMPY')
    print_task('Ensuring that the numerical operations library "numpy" is installed.')
    ensure_library_installed("numpy")
    # NUMPY HAS BEEN INSTALLED

    print_action('Update NUMPY')
    print_task('Attempting to update "numpy".')
    update_library("numpy")
    # NUMPY HAS BEEN UPDATED

    print_action('Install MATPLOTLIB')
    print_task('Ensuring that the visual plotting library "matplotlib" is installed.')
    ensure_library_installed("matplotlib")
    # MATPLOTLIB HAS BEEN INSTALLED

    print_action('Update MATPLOTLIB')
    print_task('Attempting to update "matplotlib".')
    update_library("matplotlib")
    # MATPLOTLIB HAS BEEN UPDATED

    print_action('Install SCIPY')
    print_task('Ensuring that the curve fitting library "scipy" is installed.')
    ensure_library_installed("scipy")
    # SCIPY HAS BEEN INSTALLED

    print_action('Update SCIPY')
    print_task('Attempting to update "scipy".')
    update_library("scipy")
    # SCIPY HAS BEEN UPDATED

    print_action('Install PANDAS')
    print_task('Ensuring that the data handling library "pandas" is installed.')
    ensure_library_installed("pandas")
    # PANDAS HAS BEEN INSTALLED

    print_action('Update PANDAS')
    print_task('Attempting to update "pandas".')
    update_library("pandas")
    # PANDAS HAS BEEN UPDATED

    print_action('Install SCIKIT-POSTHOCS')
    print_task('Ensuring that the statistical test library "scikit-posthocs" is installed.')
    ensure_library_installed("scikit-posthocs")
    # SCIKIT-POSTHOCS HAS BEEN INSTALLED

    print_action('Update SCIKIT-POSTHOCS')
    print_task('Attempting to update "scikit-posthocs".')
    update_library("scikit-posthocs")
    # SCIKIT-POSTHOCS HAS BEEN UPDATED




########## MAIN ENTRY POINT ##########


if __name__ == "__main__":
    print_status("Started")
    run_script()
    print_status("Completed")
