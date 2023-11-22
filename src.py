import os
import shutil
import zipfile
import win32com.client

def create_shortcut(dest_folder, shortcut_folder, file_name):
    dest_file = os.path.join(dest_folder, file_name)
    shortcut_name = os.path.splitext(file_name)[0] + ".lnk"
    shortcut_path = os.path.join(shortcut_folder, shortcut_name)

    try:
        # Create a shortcut in the shortcut folder
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = dest_file
        shortcut.WorkingDirectory = dest_folder
        shortcut.IconLocation = dest_file
        shortcut.save()

        print(f"Created shortcut for {file_name} in {shortcut_folder}")
    except Exception as e:
        print(f"Error creating shortcut for {file_name}: {e}")

def extract_and_copy(src_folder, dest_folder, shortcut_folder):
    # Ensure destination folders exist
    os.makedirs(dest_folder, exist_ok=True)

    # Ensure shortcut folder exists or create it
    os.makedirs(shortcut_folder, exist_ok=True)

    for file_name in os.listdir(src_folder):
        src_path = os.path.join(src_folder, file_name)

        # Check if the file is a zip file
        if file_name.endswith('.zip'):
            with zipfile.ZipFile(src_path, 'r') as zip_ref:
                # Extract all files from the zip archive
                zip_ref.extractall(dest_folder)
                print(f"Extracted files from {file_name} to {dest_folder}")

                # Iterate over extracted files and create shortcuts
                for extracted_file in zip_ref.namelist():
                    create_shortcut(dest_folder, shortcut_folder, extracted_file)
        elif file_name.endswith('.exe'):
            # Check if the file already exists in the destination folder
            dest_file = os.path.join(dest_folder, file_name)
            if os.path.exists(dest_file):
                print(f"{file_name} already exists in the destination folder. Skipping.")
                continue

            # Copy the .exe file to the destination folder
            shutil.copy(src_path, dest_file)
            print(f"Copied {file_name} to {dest_folder}")

            # Create a shortcut in the shortcut folder
            create_shortcut(dest_folder, shortcut_folder, file_name)

if __name__ == "__main__":
    # Specify your source, destination, and shortcut folders
    source_folder = "test/download"
    destination_folder = "test/software"
    shortcut_folder = "test/installedSoftware"

    extract_and_copy(source_folder, destination_folder, shortcut_folder)