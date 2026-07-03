# ETABS Load Combination Manager & Transfer Tool (VBA)

An automation tool built with **Excel VBA** that interfaces with the **ETABS OAPI (Open Application Programming Interface)**. This tool allows structural engineers to easily export load combinations to Excel, modify names and scale factors in a user-friendly spreadsheet environment, and import/transfer them back into the same model or a completely new ETABS structure.

## 🎯 Project Objective
Managing and editing dozens or hundreds of load combinations inside ETABS can be tedious and time-consuming. This tool streamlines the workflow by:
*   **Rapid Editing:** Allowing fast, bulk modifications of combination names and coefficients using standard Excel features.
*   **Cross-Model Transfer:** Seamlessly copying standard or custom combination matrices from a template/reference file and applying them to a newly modeled structure.
*   **Documentation:** Generating a clean, structured layout of your combinations for calculation reports.

---

## 🚀 Key Features
1.  **Dynamic Connection:** Can attach to an already running instance of ETABS or launch a fresh `.edb` file directly using a built-in file picker.
2.  **Smart Matrix Export:** Creates a beautifully mapped grid where rows represent combination names and columns represent Load Cases. 
3.  **Conflict Prevention:** When importing, the script cross-checks existing combinations in the target model. If a duplicate name is found, it prompts the user with options to **Overwrite (Yes)**, **Skip (No)**, or **Abort (Cancel)**.
4.  **Auto-Unlock:** Automatically unlocks the ETABS model prior to editing combinations to prevent OAPI modification errors.

---

## 💻 Prerequisites & Setup

To deploy this code in your Excel workbook, follow these configuration steps:

1.  **Enable ETABS Type Library:**
    *   Open Excel and press `Alt + F11` to enter the VBA IDE.
    *   Go to **Tools** > **References** from the top menu.
    *   Scroll down, locate **`ETABSv1`** (or your specific installed version, e.g., ETABS 22), check the box, and click OK.
2.  **Workbook Layout:**
    *   Ensure your workbook contains a worksheet named exactly **`Sheet1`**. The script uses this sheet explicitly to read and write data.
3.  **Insert the Code:**
    *   Right-click your VBA project explorer, choose `Insert` > `Module`, and paste the provided code inside.

> ⚠️ **Important Path Note:** In the `OpenModel_From_Selector` subroutine, the executable path is hardcoded for Version 22 (`C:\Program Files\Computers and Structures\ETABS 22\ETABS.exe`). If you are running a different version, update this string to your local installation directory.

---

## 🛠 How to Use (Workflow)

### Step 1: Exporting Combinations
1. Open your source ETABS model containing the combinations you want to extract.
2. Run the `Export_Load_Combinations_To_Excel` macro.
3. The script populates `Sheet1` with a complete matrix of your load combinations and their corresponding factors.

### Step 2: Excel Editing
* Feel free to rename combinations in the first column, insert new rows, or adjust the scale factors under the specific Load Case columns.

### Step 3: Importing / Transferring
1. Open the target ETABS model (the new structure where you want these combinations applied).
2. Run the `Import_Load_Combinations_From_Excel` macro.
3. The tool safely pushes the spreadsheet definitions directly into the active ETABS database.

---

## 📝 Excel Matrix Structure Reference
Once exported, your data will align in `Sheet1` in the following matrix format:

| Combo Name | DEAD | LIVE | EX | EY |
| :--- | :---: | :---: | :---: | :---: |
| **Combo1** | 1.4 | | | |
| **Combo2** | 1.2 | 1.6 | | |
| **Combo3** | 1.2 | 1.0 | 1.0 | |

*Empty cells or cells with a value of `0` are automatically ignored during the import process to keep your ETABS database clean.*

---

## 🤝 Contributing
Contributions, bug reports, and feature requests (such as automatic ETABS version detection or support for multiple combo types like Envelope/Absolute) are welcome! Feel free to open an issue or submit a pull request.
