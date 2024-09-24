# WESP-AC Version 3.1 to 3.4 Data Transfer


## Missing/Extra Data from 3.1 to 3.3/3.4
- In version 3.2, **OF28** (Fish Access or Use) spans rows 151–154. In version 3.4, **OF28** (Anadromous Species Access or Use) is consolidated into rows 149–150. If either value from the consecutive rows in version 3.2 (e.g., rows 151 or 152) is 1, the corresponding row in version 3.4 (row 149) is set to 1. If neither value is 1, the row in version 3.4 is set to 0. This same logic applies to rows 153 and 154 in version 3.2, which are combined into row 150 in version 3.4.
- In version 3.4, **F19** (Shallow Open Ponded Water + Bare Saturated Substrate) has an additional option (>100,000 sq.m) compared to version 3.1, which we ignored.
- In version 3.4, **F32** (Ponded Open Water - Minimum Size) is new so it is ignored when transfering.






## Script Functionality

The script performs the following operations:

1. **File Handling**:
   - The script reads WESP-AC 3.1 files from a specified folder and saves the converted 3.4 files into a new folder named `<version>_transferred`.
   - For each file, the data from sheets `OF`, `F`, and `S` is read from version 3.1 and pasted into the corresponding locations in version 3.4.

2. **Sheet Adjustments**:
   - **OF Sheet** (same as 3.0):
		- **OF1-OF18**: Data from D5 in version 3.1 is shifted by 1 row when copying to D6 in version 3.4 until D108.
		- **OF19**: D110 to D111 in version 3.1 map directly to D110 to D111 in version 3.4.
		- **OF20-27**: Rows D112 to D150 in version 3.1 are shifted up by 1 in version 3.4.
		- **OF28**: Special case: Rows D151 to D154 in version 3.1 are merged into D149 and D150 in version 3.4 based on their values.
		- **OF29-OF33**: Rows D156 to D169 are copied from D152 to D165 in version 3.4, with an empty cell at D169 in 3.1.
		- **OF33-OF38**: D170 to D179 in version 3.1 map to D165 to D174 in version 3.4.

   - **F Sheet**:
		- **F1-F19**: Data from row D5 to D123 is directly copied between both versions (3.1 and 3.4). (handles F19 Issue)
		- **F19**: Row D123 is set to 0 in version 3.4. 
		- **F20-F39**: Rows from D124 to D191 in 3.1 are shifted up by one row to D125 to D192 in 3.4. (handles F32)
		- **F41**: For rows D194 to D222 in 3.1, they are shifted by two rows in 3.4 (D196 to D224).
		- **F40**: For rows D222 to D227 in 3.1, they are shifted by three rows in 3.4 (D225 to D230).
		- **F43-F44**: For rows D228 to D239 in 3.1, they are shifted by four rows in 3.4 (D232 to D243).
		- **F42**: Row D240 in 3.1 is shifted to D231 in 3.4.
		- **F45**: Rows D242 to D247 in 3.1 are shifted to D244 to D249 in 3.4.
		- **F65-F66**: Rows D249 to D253 in 3.1 are shifted to D337 to D341 in 3.4.
		- **F53**: Special case: If either D289 or D290 in 3.1 equals 1, then D284 in 3.4 is set to 1.
		- **F54-F62**: Rows from D291 to D332 in 3.1 are shifted to D285 to D326 in 3.4.
		- **F62-64**: Rows from D333 to D341 in 3.1 are shifted to D328 to D336 in 3.4.


   - **S Sheet**:
     - The data is identical between both versions from F6 to F88.
     - Extra rows in version 3.4 from F89 to F101 are not filled and are left as is.

3. **Special Case Handling**:
   - For specific rows, such as D194 in the F sheet, a condition is applied: if D194 in version 3.1 is 1, then D194 in version 3.4 is set to 0, and vice versa.

## How to Use the Script

1. Place the WESP-AC 3.1 files in a folder named `3.1` in the same directory as the script.
2. Ensure the WESP-AC 3.4 template file is named `wespac_3.4.xlsx` and located in the same directory.
3. Run the script (`wespac_3.1_to_3.4.py`). It will generate a new folder named `3.1_transferred`, where the updated files will be saved with a `_3.4` suffix.
4. The script automatically adjusts the data based on the specific rules defined for each sheet.

## Dependencies

- `xlwings`: Used to read and manipulate Excel files.
- Ensure that you have Excel installed on your system for `xlwings` to function properly.

You can install the required dependencies using:

```bash
pip install xlwings
