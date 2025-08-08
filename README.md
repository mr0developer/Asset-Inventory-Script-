# Asset Inventory Script (Windows 10/11)

This PowerShell script collects basic hardware and software information from a Windows 10 or 11 machine and saves it as a text file on the root of the USB drive where the script is run.  
It is designed to run without administrator privileges and is useful for quick IT asset inventory.

## Features
- Detects **Windows version** and **activation status**  
- Detects **Microsoft Office version** and **activation status**  
- Retrieves the **machine serial number**  
- Checks if **OneDrive** is installed  
- Detects installed **antivirus software** and attempts to retrieve its **installation date**  
- Reports **installed RAM size**  
- Works without admin rights  

## Usage
1. Copy the script to a USB flash drive.  
2. Insert the USB into a Windows 10/11 machine.  
3. Run the script manually from the USB.  
4. The results will be saved in `asset_inventory.txt` at the root of the USB drive.  
