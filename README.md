# OptionsLib 
## what is this
this is a pretty simple excel plugin that I wrote in python to get better at python and help out a friend.
it definitely needs some optimization, commenting, etc, but its a solid start and it's pretty functional.
if you have any questions feel free to email me!

---

## INSTALLATION
You'll need all of the .py files in the same directory as your excel spreadsheet file.
In addition, your workbook needs to be saved as a macro enabled workbook, so make sure you do that through excel by using "**Save as**" and selecting "**Macro Enabled Workbook**", the extension should be **.xlsm**
Now, you'll need to make sure that python is installed, and I'd recommend creating a venv using a terminal.

#### 1. Navigate to your project directory
cd \path\to\your\project

#### 2. Create a virtual environment (replace 'venv' with your preferred name)
python -m venv venv

#### 3. Activate the virtual environment
venv\Scripts\activate
(activate.ps1 if using Powershell)

#### 4. Install required packages
pip install xlwings yfinance diskcache

#### 5. Deactivate the environment when done
deactivate

### Installing the xlwings Excel Add-in on Windows

Follow these steps to install the xlwings add-in for Excel:

#### 1. Install the Excel Add-in

xlwings addin install

 This command copies the xlwings add-in to Excel’s XLSTART folder, making it available when you launch Excel.

#### 3. Set VBA Reference for Advanced Features

##### If you want to use RunPython or user-defined functions (UDFs) from VBA:
##### - Open Excel and press Alt + F11 to open the VBA editor.
##### - Go to Tools > References...
##### - Check the box for xlwings in the list.

#### 4. Keep Versions in Sync
#####  Close Excel before running the install command to avoid errors.
pip install --upgrade xlwings
xlwings addin install

#### 5. Uninstalling

xlwings addin remove
pip uninstall xlwings

## Notes:
 - Excel must be installed on your system.
 - The add-in version should always match the Python package version.
