# Excel Date Checker - Background Process

This project automates the process of checking for dates coming up in Excel files, specifically identifying dates that are two days ahead. It runs as a background process when Windows starts, providing notifications without needing to open the Excel file.

## Usage

1. **Insert VBA Code:** Copy and paste the provided VBA code into the Excel file that you want to check for upcoming dates and check the date column.

2. **Modify File Paths:** Open the shell script and VBS file and update the directory paths as needed to point to the Excel file you want to monitor.

3. **Set Up Startup Program:** Create a shortcut to the VBS file and place it in the Windows startup program folder. This will ensure that the background process runs automatically when Windows starts.

    - **Alternative:** Alternatively, you can use Task Scheduler to schedule the batch file to run at system startup.

## Contributions

Contributions, suggestions, and feedback are welcome! If you have ideas for improvements or additional features, feel free to fork the repository, make your changes, and submit a pull request.
