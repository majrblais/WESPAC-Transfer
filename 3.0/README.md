# WESP-AC Version 3.0 to 3.4 Data Transfer

This repository contains a Python script that automates the transfer of data from WESP-AC version 3.0 files to version 3.4. The script reads the data from version 3.0, accounts for structural changes in version 3.4, and pastes the values into a 3.4 template, applying necessary adjustments for missing rows, additional rows, and special cases where inferred data is required.

## Script Functionality

The script performs the following operations:

### 1. **OF Sheet (Unchanged from 3.1 to 3.4)**:
   - The data structure in the OF sheet remains identical to the 3.1 to 3.4 transition, with no modifications required.

### 2. **F Sheet Adjustments**:
   - **Starting at different rows**:
     - In version 3.0, data starts at row D3, while in version 3.4, it starts at D5.
   
   - **Row Mapping**:
     - D3 to D52 (3.0) is equal to D5 to D54 (3.4).
     - **Skipping row D55**: The script skips row D55 in version 3.4.
     - D53 to D93 (3.0) is equal to D56 to D96 (3.4).
     - **Skipping row D97**: The script skips row D97 in version 3.4.
     - D94 to D118 (3.0) is equal to D98 to D122 (3.4).
     - **Skipping row D123**: The script skips row D123 in version 3.4.
     - D119 to D129 (3.0) is equal to D124 to D134 (3.4).
     - **Extra row in 3.0 (D130)**: Row D130 in version 3.0 is skipped.
     - D131 to D189 (3.0) is equal to D135 to D193 (3.4).
     - **Special Case for D194 (3.4)**: If D191 in version 3.0 is `0`, then D194 in 3.4 is set to `1`. Otherwise, D194 in 3.4 is set to `0`.
     - D190 to D218 (3.0) is equal to D195 to D223 (3.4).
     - **Special Row Mapping**:
       - D224 (3.0) is equal to D224 (3.4).
       - D219 to D223 (3.0) is equal to D225 to D229 (3.4).
       - D237 and D238 (3.0) are equal to D230 and D231 (3.4).
       - D225 to D236 (3.0) are equal to D232 to D243 (3.4).
     - D239 to D244 (3.0) is equal to D244 to D249 (3.4).
     - **Extra Row in 3.4 (D250)**: Row D250 in version 3.4 is set to `0`.
     - D252 to D284 (3.0) is equal to D251 to D283 (3.4).
     - **Special Case for D285 and D286**: If either D285 or D286 in version 3.0 is `1`, then D284 in version 3.4 is set to `1`.
     - D287 to D328 (3.0) is equal to D285 to D326 (3.4).
     - **Extra Row in 3.4 (D327)**: Row D327 in version 3.4 is set to `0`.
     - D329 to D337 (3.0) is equal to D328 to D336 (3.4).
     - D250 and D251 (3.0) are equal to D342 and D343 (3.4).
     - **Extra Rows in 3.4 (D344 and D345)**: Rows D344 and D345 in version 3.4 are set to `0`.

   - **Additional Special Mapping**:
     - D66 to D72 (3.0) is equal to D72 to D78 (3.4).
     - D73 to D75 (3.0) is equal to D69 to D71 (3.4).
     - D345 to D351 (3.0) is equal to D238 to D343 (3.4).

### 3. **S Sheet Adjustments**:
   - In version 3.0, the data starts at F3, while in version 3.4, it starts at F6.
   - Data from F3 to F85 (3.0) is equal to F6 to F88 (3.4).
   - Any rows in version 3.4 beyond F88 (up to F101) are left as they are, since they do not exist in version 3.0.

## Missing Data and Inferred Values

For missing rows in version 3.0 or extra rows in version 3.4, the following inferred values are used:
- **Row D250 in 3.4**: Set to `0`.
- **Row D327 in 3.4**: Set to `0`.
- **Rows D344 and D345 in 3.4**: Set to `0`.
- **Special Case for D194 in 3.4**: If D191 in version 3.0 is `0`, then D194 in version 3.4 is set to `1`. Otherwise, D194 in 3.4 is set to `0`.
- **Special Case for D284 in 3.4**: If either D285 or D286 in version 3.0 is `1`, then D284 in version 3.4 is set to `1`.

## How to Use the Script

1. Place the WESP-AC 3.0 files in a folder named `3.0` in the same directory as the script.
2. Ensure the WESP-AC 3.4 template file is named `wespac_3.4.xlsx` and located in the same directory.
3. Run the script. It will generate a new folder named `3.0_transferred`, where the updated files will be saved with a `_3.4` suffix.
4. The script automatically adjusts the data based on the specific rules defined for each sheet.

## Dependencies

- `xlwings`: Used to read and manipulate Excel files.
- Ensure that you have Excel installed on your system for `xlwings` to function properly.

You can install the required dependencies using:

```bash
pip install xlwings
