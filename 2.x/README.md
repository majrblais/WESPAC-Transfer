# WESP-AC Data Transfer from Version 2.1 to 3.4

This script transfers data from version 2.1 WESP-AC files to the newer version 3.4. The data is mapped according to the specific shifts and modifications between these versions.

- In version 3.4, Distance to Domestic Wells is **OF10** while in 2.x version this indicator was in F, it was moved.
- in version 3.4, **F36** is extra while it is missing in previous versions.
- In version 3.4 **F64**, **F67** and **F68** are missing from earlier versions and are new.
- In version 3.4 **S6**,is new.


## 1. **File Structure**
   - Input folder: Contains WESP-AC version 2.1 files.
   - Output folder: A newly created folder named `version_transferred`, where the converted files (version 3.4) will be saved.

## 2. **Sheet Adjustments**:
   - **OF Sheet**:
     - **OF1-OF2**: Data from D4 to D15 in version 2.1 is copied directly to D4 to D16 in version 3.4.
     - **OF3-OF5**: Rows D16 to D38 in version 2.1 are shifted up by 2 rows to D18 to D40 in version 3.4.
     - **OF6**: Row D39 in version 2.1 is shifted to D42 in version 3.4.
     - **OF7**: Row D40 in version 2.1 is shifted to D44 in version 3.4.
     - **OF8**: Rows D41 to D46 in version 2.1 are shifted to D46 to D51 in version 3.4.
     - **OF9**: Rows D50 to D55 in version 2.1 are shifted to D53 to D58 in version 3.4.
     - **OF11**: Rows D56 to D62 in version 2.1 are shifted to D65 to D71 in version 3.4.
     - **OF12-OF15**: Rows D63 to D85 in version 2.1 are shifted to D72 to D96 in version 3.4.
     - **OF16-OF17**: Rows D86 to D96 in version 2.1 are shifted to D97 to D108 in version 3.4.
     - **OF18-OF22**: Rows D97 to D113 in version 2.1 are shifted to D109 to D125 in version 3.4.
     - **OF23-OF26**: Rows D114 to D132 in version 2.1 are shifted to D127 to D145 in version 3.4.
     - **OF27-OF29**: Rows D133, D135, D136, D140 to D144 in version 2.1 are shifted to D147 to D158 in version 3.4.
     - **OF30-OF38**: Rows D145 to D157 in version 2.1 are shifted to D160 to D174 in version 3.4.
	 - **OF10**: Rows D303 to D307 in F sheet in version 2.1 are shifted down by 242 into the OF sheet in version 3.4

   - **F Sheet**:
     - **F1**: Data from rows D5 to D10 in version 2.1 are shifted up by 1 row to D6 to D11 in version 3.4.
     - **F2**: Rows D13 to D16 in version 2.1 are shifted up by 1 row to D14 to D17 in version 3.4.
     - **F3**: Rows D18 to D23 in version 2.1 are shifted by 2 rows to D20 to D25 in version 3.4.
     - **F4**: Rows D26 to D27 in version 2.1 are shifted by 4 rows to D30 to D31 in version 3.4.
     - **F5-F6**: Rows D29 to D43 in version 2.1 are shifted by 5 rows to D34 to D48 in version 3.4.
     - **F7**: Rows D89 to D93 in version 2.1 are shifted by 39 rows back to D50 to D54 in version 3.4.
     - **F8**: Rows D95 to D99 in version 2.1 are shifted by 38 rows back to D57 to D61 in version 3.4.
     - **F9**: Rows D191 to D194 in version 2.1 are shifted by 127 rows back to D64 to D67 in version 3.4.
     - **F10**: Rows D106 to D107 in version 2.1 are shifted by 36 rows back to D70 to D71 in version 3.4.
     - **F11**: Rows D109 to D113 in version 2.1 are shifted by 36 rows back to D73 to D77 in version 3.4.
     - **F12**: Rows D115 to D118 in version 2.1 are shifted by 35 rows back to D80 to D83 in version 3.4.
     - **F13**: Rows D52 to D56 in version 2.1 are shifted forward by 33 rows to D85 to D89 in version 3.4.
     - **F14**: Rows D58 to D62 in version 2.1 are shifted forward by 34 rows to D92 to D96 in version 3.4.
     - **F15**: Rows D45 to D47 in version 2.1 are shifted by 54 rows forward to D99 to D101 in version 3.4.
     - **F16**: Rows D49 to D50 in version 2.1 are shifted by 55 rows forward to D104 to D105 in version 3.4.
     - **F17**: Rows D64 to D68 in version 2.1 are shifted forward by 44 rows to D108 to D112 in version 3.4.
     - **F18**: Rows D74 to D76 in version 2.1 are shifted forward by 41 rows to D115 to D117 in version 3.4.
     - **F19**: Rows D84 to D87 in version 2.1 are shifted forward by 35 rows to D119 to D122 in version 3.4.
     - **F20**: Rows D70 to D72 in version 2.1 are shifted forward by 55 rows to D125 to D127 in version 3.4.
     - **F21**: Rows D78 to D82 in version 2.1 are shifted forward by 51 rows to D129 to D133 in version 3.4.
     - **F22**: Rows D141 to D145 in version 2.1 are shifted back by 5 rows to D136 to D140 in version 3.4.
     - **F23**: Rows D129 to D133 in version 2.1 are shifted by 14 rows forward to D143 to D147 in version 3.4.
     - **F24**: Rows D122 to D128 in version 2.1 are shifted forward by 28 rows to D150 to D156 in version 3.4.
     - **F25**: Rows D135 to D139 in version 2.1 are shifted by 24 rows forward to D159 to D163 in version 3.4.
     - **F26-F27**: Rows D119 to D120 in version 2.1 are shifted forward by 46 rows to D165 to D166 in version 3.4.
     - **F28**: Rows D147 to D151 in version 2.1 are shifted forward by 22 rows to D169 to D173 in version 3.4.
     - **F29**: Rows D154 to D158 in version 2.1 are shifted forward by 22 rows to D176 to D180 in version 3.4.
     - **F30**: Rows D160 to D162 in version 2.1 are shifted forward by 23 rows to D183 to D185 in version 3.4.
     - **F31**: Rows D164 to D168 in version 2.1 are shifted forward by 24 rows to D188 to D192 in version 3.4.
     - **F33**: Rows D171 to D176 in version 2.1 are shifted forward by 25 rows to D196 to D201 in version 3.4.
     - **F34**: Rows D178 to D183 in version 2.1 are shifted forward by 25 rows to D203 to D208 in version 3.4.
     - **F35**: Rows D185 to D189 in version 2.1 are shifted forward by 26 rows to D211 to D215 in version 3.4.
     - **F36**: Rows D196 to D198 in version 2.1 are shifted forward by 22 rows to D218 to D220 in version 3.4.
     - **F37**: Row D199 in version 2.1 is mapped directly to D222 in version 3.4.
     - **F38**: Row D205 in version 2.1 is mapped directly to D223 in version 3.4.
     - **F39**: Row D204 in version 2.1 is mapped directly to D224 in version 3.4.
     - **F40**: Rows D201 to D203 in version 2.1 are shifted forward by 25 rows to D226 to D228 in version 3.4.
     - **F41-F42**: Rows D216 to D217 in version 2.1 are shifted forward by 14 rows to D230 to D231 in version 3.4.
     - **F43**: Rows D219 to D223 in version 2.1 are shifted forward by 14 rows to D233 to D237 in version 3.4.
     - **F44**: Rows D213 to D215 in version 2.1 are shifted forward by 27 rows to D240 to D242 in version 3.4.
     - **F45**: Rows D219 to D223 in version 2.1 are shifted forward by 18 rows to D245 to D249 in version 3.4.
     - **F47-F48**: Rows D234 to D240 in version 2.1 are shifted forward by 18 rows to D252 to D258 in version 3.4.
	 - **F49**: Rows D242 to D245 in version 2.1 are shifted forward by 19 rows to D261 to D264 in version 3.4.
     - **F50-F52**: Rows D248 to D260 in version 2.1 are shifted forward by 19 rows to D267 to D279 in version 3.4.
     - **F54**: Rows D263 to D268 in version 2.1 are shifted forward by 24 rows to D287 to D292 in version 3.4.
     - **F55**: Rows D270 to D273 in version 2.1 are shifted forward by 25 rows to D295 to D298 in version 3.4.
     - **F56-F58**: Rows D275 to D288 in version 2.1 are shifted forward by 26 rows to D301 to D314 in version 3.4.
     - **F59-F62**: Rows D290 to D299 in version 2.1 are shifted forward by 27 rows to D317 to D326 in version 3.4. (handles moose hunting addition)
     - **F62**: Rows D300 to D302 in version 2.1 are shifted forward by 28 rows to D328 to D330 in version 3.4.
     - **F63**: Row D307 in version 2.1 is mapped directly to D332 in version 3.4.
     - **F65-F66**: Rows D225 to D230 in version 2.1 are shifted forward by 113 rows to D338 to D343 in version 3.4.

   - **S Sheet**:
     - **S1**: Rows F18 to F22 in version 2.1 are shifted up by 2 rows to F20 to F24 in version 3.4.
     - **S2**: Rows F33 to F35 in version 2.1 are shifted up by 2 rows to F35 to F37 in version 3.4.
     - **S3**: Rows F46 to F48 in version 2.1 are shifted up by 2 rows to F48 to F50 in version 3.4.
     - **S4**: Rows F63 to F66 in version 2.1 are shifted up by 2 rows to F65 to F68 in version 3.4.
     - **S5**: Rows F81 to F84 in version 2.1 are shifted up by 2 rows to F83 to F86 in version 3.4.

## 3. **Running the Script**
To process WESP-AC version 2.1 files:
1. Place the WESP-AC version 2.1 files in the input folder.
2. Run the script, and the processed files will be saved in a folder named `2.1_transferred`.
3. The script will copy the data as per the mappings outlined above.
