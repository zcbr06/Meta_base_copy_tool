FILE COPY TOOL - README

DESCRIPTION:

This PowerShell script allows you to copy files from a selected source directory to a destination directory based on the file metadata (specifically, the "Author" field). The script provides an interactive menu for easy navigation and selection of required parameters.


REQUIREMENTS:

- Windows 10 or later.

- PowerShell 5.1 or newer.

- Administrator privileges.

USAGE INSTRUCTIONS:

- Run the Script

- Right-click on the script file and select "Run as Administrator".

- If prompted by UAC, click "Yes" to allow execution.

- Menu Navigation

- The script presents an interactive menu with five options:

- Select source folder: Choose the directory where files are located.

- Select destination folder: Choose the directory where files will be copied.

- Enter author name: Specify the author's name from metadata.

- Start copy process: Begin copying files.

- Exit: Close the script.

- Select an option by pressing the corresponding number key.
The selected paths and author name are displayed in the menu.

Copy Process

The script scans all files in the source directory.
It retrieves the author from metadata and compares it with the entered name.
Matching files are copied to the destination folder.
A progress bar displays the percentage of completion.
Successfully copied files are listed in the terminal.

Completion

Once the process is complete, the script will display a summary of copied files.

Press any key to exit.

NOTES:

If no files match the specified author, nothing will be copied.
Ensure that the metadata of the target files includes the correct author name.
The script only copies files (does not move or delete original files).
The destination folder is created automatically if it does not exist.

TROUBLESHOOTING:

- Script does not start: Ensure you run it as Administrator.

- Files are not copied: Check if metadata contains the correct author name.

- PowerShell Execution Policy error: Run     Set-ExecutionPolicy Unrestricted -Scope Process in PowerShell before executing the script.