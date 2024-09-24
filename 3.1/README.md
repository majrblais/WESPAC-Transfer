# WESP-AC Version 3.1 to 3.4 Data Transfer

This repository contains a Python script to automate the transfer of data from WESP-AC version 3.1 files to version 3.4. The script reads data from version 3.1, adjusts for structural differences between the two versions, and copies the values into a version 3.4 template, preserving the formatting in the new file.

## Script Functionality

The script performs the following operations:

1. **File Handling**:
   - The script reads WESP-AC 3.1 files from a specified folder and saves the converted 3.4 files into a new folder named `<version>_transferred`.
   - For each file, the data from sheets `OF`, `F`, and `S` is read from version 3.1 and pasted into the corresponding locations in version 3.4.

2. **Sheet Adjustments**:
   - **OF Sheet**:
     - Data from D5 in version 3.1 is shifted by 1 row when copying to D6 in version 3.4 until D108.
     - D110 to D111 in version 3.1 map directly to D110 to D111 in version 3.4.
     - Rows D112 to D150 in version 3.1 are shifted up by 1 in version 3.4.
     - Special case: Rows D151 to D154 in version 3.1 are merged into D149 and D150 in version 3.4 based on their values.
     - Rows D156 to D169 are copied from D152 to D165 in version 3.4, with an empty cell at D169 in 3.1.
     - D170 to D179 in version 3.1 map to D165 to D174 in version 3.4.

   - **F Sheet**:
     - Data up to row D192 is copied directly between the two versions.
     - Row D193 in version 3.1 is shifted to D195 in version 3.4.
     - Rows D222 to D226 in version 3.1 map to D225 to D229 in version 3.4.
     - Rows D228 to D239 in version 3.1 map to D232 to D243 in version 3.4.
     - Special handling for D240 and D241 in version 3.1, which map to D231 in version 3.4.
     - Rows D248 to D253 in version 3.1 map to D337 to D341 in version 3.4.
     - Special case: If either D289 or D290 in version 3.1 is 1, then D284 in version 3.4 is set to 1.
     - Rows D291 to D332 in version 3.1 map to D285 to D326 in version 3.4.
     - Rows D333 to D341 in version 3.1 map to D328 to D336 in version 3.4.

   - **S Sheet**:
     - The data is identical between both versions from F6 to F88.
     - Extra rows in version 3.4 from F89 to F101 are not filled and are left as is.

3. **Special Case Handling**:
   - For specific rows, such as D194 in the F sheet, a condition is applied: if D194 in version 3.1 is 1, then D194 in version 3.4 is set to 0, and vice versa.

## How to Use the Script

1. Place the WESP-AC 3.1 files in a folder named `3.1` in the same directory as the script.
2. Ensure the WESP-AC 3.4 template file is named `wespac_3.4.xlsx` and located in the same directory.
3. Run the script. It will generate a new folder named `3.1_transferred`, where the updated files will be saved with a `_3.4` suffix.
4. The script automatically adjusts the data based on the specific rules defined for each sheet.

## Dependencies

- `xlwings`: Used to read and manipulate Excel files.
- Ensure that you have Excel installed on your system for `xlwings` to function properly.

You can install the required dependencies using:

```bash
pip install xlwings
