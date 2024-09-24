# WESP-AC Version 3.0 to 3.4 Data Transfer


## Missing/Extra Data from 3.0 to 3.3/3.4
- In version 3.2, **OF28** (Fish Access or Use) spans rows 151–154. In version 3.4, **OF28** (Anadromous Species Access or Use) is consolidated into rows 149–150. If either value from the consecutive rows in version 3.2 (e.g., rows 151 or 152) is 1, the corresponding row in version 3.4 (row 149) is set to 1. If neither value is 1, the row in version 3.4 is set to 0. This same logic applies to rows 153 and 154 in version 3.2, which are combined into row 150 in version 3.4.
- In version 3.4, **F19** (Shallow Open Ponded Water + Bare Saturated Substrate) has an additional option (>100,000 sq.m) compared to version 3.0, which we ignored.
- In version 3.4, **F32** (Ponded Open Water - Minimum Size) is new so it is ignored when transfering.
- In version 3.4, **F53** has one less option than in version 3.0, however, this fix was easy since we just combined them
- In version 3.4, **F62** has an extra option (moose hunting) which was not in earlier versions, it was just set at 0.




## Script Functionality

The script performs the following operations:

1. **File Handling**:
   - The script reads WESP-AC 3.0 files from a specified folder and saves the converted 3.4 files into a new folder named `<version>_transferred`.
   - For each file, the data from sheets `OF`, `F`, and `S` is read from version 3.0 and pasted into the corresponding locations in version 3.4.

2. **Sheet Adjustments**:
   - **OF Sheet** (same as 3.1):
     - **OF1-OF18**: Data from D5 in version 3.0 is shifted by 1 row when copying to D6 in version 3.4 until D108.
     - **OF19**: D110 to D111 in version 3.0 map directly to D110 to D111 in version 3.4.
     - **OF20-27**: Rows D112 to D150 in version 3.0 are shifted up by 1 in version 3.4.
     - **OF28**: Special case: Rows D151 to D154 in version 3.0 are merged into D149 and D150 in version 3.4 based on their values.
     - **OF29-OF33**: Rows D156 to D169 are copied from D152 to D165 in version 3.4, with an empty cell at D169 in 3.0.
     - **OF33-OF38**: D170 to D179 in version 3.0 map to D165 to D174 in version 3.4.

   - **F Sheet**:
     - **F1-F7**: Data from row D3 to D52 in version 3.0 is shifted by 2 rows when copying to D5 in version 3.4.
     - **F8-F14**: Rows D53 to D93 in version 3.0 are shifted up by 3 rows in version 3.4.
     - **F10**: Row D66 to D72 in version 3.0 are shifted up by 6 rows in version 3.4.
     - **F11**: Row D73 to D76 in version 3.0 are shifted down by 6 rows in version 3.4.

     - **F15-F19**: Rows D94 to D118 in version 3.0 are shifted up by 4 rows in version 3.4.
     - **F20-F21**: Rows D119 to D129 in version 3.0 are shifted up by 5 rows in version 3.4.
     - **F22-31**: Rows D131 to D189 in version 3.0 are shifted by 4 rows in version 3.4.
     - **F32**: Special case: D194 in version 3.4 is set to 1 if D191 in version 3.0 is 0, else 0. (New F)
     - **F33-F38**: Rows D190 to D218 in version 3.0 are shifted by 6 rows in version 3.4.
     - **F39**: Row D224 to D223 in version 3.0 is also row D224 in version 3.4.
     - **F40**: Rows D219 to D223 in version 3.0 are shifted by 5 rows in version 3.4.
     - **F41**: Row D237 to D223 in version 3.0 are shifted to row 230 in version 3.4.
     - **F42**: Row D238 to D223 in version 3.0 are shifted to row 231 in version 3.4.
     - **F43-F44**: Rows D225 to D237 in version 3.0 are shifted up by 7 rows in version 3.4.
     - **F45**: Rows D239 to D245 in version 3.0 are shifted up by 5 rows in version 3.4.
     - **F46**: Special: D250 in version 3.4 is set to 0.
     - **F47-F53**: Rows D252 to D284 in version 3.0 are shifted by 1 row in version 3.4.
     - **F53**: Special case for D285 or D286 in version 3.0 is merged into D284 in version 3.4.
     - **F54-F62**: Rows D287 to D328 in version 3.0 are shifted by 2 rows in version 3.4. (handles new option in F62)
     - **F62-F64**: Rows D329 to D338 in version 3.0 are shifted to D238 to D343 in version 3.4.

   - **S Sheet**:
     - **S**: The data starts at F3 in version 3.0 and F6 in version 3.4, ends at F85 in version 3.0 and F88 in version 3.4.


3. **Special Case Handling**:
   - For specific rows, such as D194 in the F sheet, a condition is applied: if D194 in version 3.0 is 1, then D194 in version 3.4 is set to 0, and vice versa.

## How to Use the Script

1. Place the WESP-AC 3.0 files in a folder named `3.0` in the same directory as the script.
2. Ensure the WESP-AC 3.4 template file is named `wespac_3.4.xlsx` and located in the same directory.
3. Run the script (`wespac_3.0_to_3.4.py`). It will generate a new folder named `3.0_transferred`, where the updated files will be saved with a `_3.4` suffix.
4. The script automatically adjusts the data based on the specific rules defined for each sheet.

## Dependencies

- `xlwings`: Used to read and manipulate Excel files.
- Ensure that you have Excel installed on your system for `xlwings` to function properly.

You can install the required dependencies using:

```bash
pip install xlwings
