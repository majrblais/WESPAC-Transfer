# WESP-AC Data Migration: Version 1.1 to 3.4

This repository contains a Python script designed to migrate data from WESP-AC version 1.1 to version 3.4. The script reads values from specified cells in the 1.1 version and writes them into the appropriate cells in version 3.4, handling both sheet-to-sheet and cross-sheet transfers. The mapping of cells has been carefully designed to accommodate changes between the versions.

## Missing/Extra Data from 1.1 to 3.4

- In version 1.1, **OF1** (Wetland Herbaceous Area) is not used in 3.4.
- In version 3.4, both **OF1** and **OF2** (Province, Ponded Area Within 1 km) are missing.
- In version 1.1, **OF16** (Flood Zone) and **OF17** (Flood Damage) are combined into **OF17** (Flood Damage from Non-tidal Waters) in 3.4, and logic was used to replace values.
- In version 1.1, **OF18** (Relative Elevation in Watershed) is a multiple-choice answer, while in 3.4 it is a float between `0` and `1.0`. We replace the options with `0`, `0.33`, and `0.66`.
- In version 1.1, **OF20** (Erosion Potential Upslope) is not used in version 3.4.
- In version 1.1, **OF28** (Growing Degree Days) is a multiple-choice option, while in version 3.4 it is an integer. We used the average of the multiple-choice answer.
- In version 1.1, **F59** (Ownership) is moved into **OF38** in version 3.4.
- In versions 1.1 and 3.4, **F1** and **F2** have distinct differences (e.g., wetland type in 1.1 vs. less specific types in 3.4). These were mapped accordingly to the best of our ability.
- In version 3.4, **F4** (Dominance of Low Shrub Genera) is missing from version 1.1.
- In version 1.1, **F18** (Sedge Cover) and **F19** (Dominance of Most Abundant Herbaceous Species) are missing from version 3.4.
- In version 3.4, **F19** (Shallow Open Ponded Water + Bare Saturated Substrate) has an additional option (>100,000 sq.m) compared to version 1.1, which we ignored.
- In version 1.1, **F55** (Cliffs, Steep Banks, or Salt Lick) is missing from version 3.4.
- In version 3.4, **F46** (Fishless) is missing from version 1.1.
- In version 3.4, **F53** (Distance to Steep Bank, Bridge, Building, or Nest Structure) is missing from 1.1. In 1.1, **F55** is similar but only refers to presence, not distance.
- In version 1.1, **F62** (Consumption Uses) is missing "Moose Hunting" while 3.4 has it.
- In version 3.4, **F64** (Keystone Species Presence), **F67** (Methane Source), and **F68** (Methane Suppression) are missing from version 1.1.
- In version 3.4, **S6** is missing from version 1.1, while 1.1 has **S1-S4** as extra fields.


## 1. **File Structure**
   - Input folder: Contains WESP-AC version 1.1 files.
   - Output folder: A newly created folder named `version_transferred`, where the converted files (version 3.4) will be saved.

## 2. **Sheet Adjustments**:
   - **OF Sheet**:
	- **OF3** D12 to D17 in version 1.1 -> D19 to D25 in version 3.4
	- **OF4** D19 to D25 in version 1.1 -> D26 to D32 in version 3.4
	- **OF5** D19 to D25 in version 1.1 -> D34 to D40 in version 3.4
	- **OF6** D34 in version 1.1 -> D42 in version 3.4
	- **OF7** D35 in version 1.1 -> D44 in version 3.4
	- **OF8** D37 to D41 in version 1.1 -> D47 to D51 in version 3.4
	- **OF9** D46 to D50 in version 1.1 -> D54 to D58 in version 3.4
	- **OF11** In the **F sheet** of 1.1, **D302 to D304** are equal to **D61 to D63** in the **OF sheet** of 3.4.
	- **OF10** D52 to D57 in version 1.1 -> D66 to D71 in version 3.4
	- **OF12** D58 in version 1.1 -> D73 in version 3.4
	- **OF13** D60 to D66 in version 1.1 -> D75 to D81 in version 3.4
	- **OF14** D68 to D73 in version 1.1 -> D83 to D88 in version 3.4
	- **OF15** D75 to D80 in version 1.1 -> D90 to D95 in version 3.4
	- **OF16** D82 to D86 in version 1.1 -> D98 to D102 in version 3.4
	- **OF17** D87 and D88 in version 1.1 -> Special handling for D104 to D107 in version 3.4
	- **OF18** D90 to D92 in version 1.1 -> D109 in version 3.4 as 0, 0.33, or 0.66
	- **OF19** D93 in version 1.1 -> D110 in version 3.4
	- **OF20** D96 to D99 in version 1.1 -> D112 to D116 in version 3.4
	- **OF21** D101 to D105 in version 1.1 -> D117 to D120 in version 3.4
	- **OF22** D106 to D109 in version 1.1 -> D122 to D125 in version 3.4
	- **OF23** D111 to D113 in version 1.1 -> D128 to D131 in version 3.4
	- **OF24** D115 to D117 in version 1.1 -> D132 to D134 in version 3.4
	- **OF25** D119 to D122 in version 1.1 -> D136 to D138 in version 3.4
	- **OF26** D123 to D128 in version 1.1 -> D140 to D145 in version 3.4
	- **OF27** D130 to D135 in version 1.1 -> Special case handling for D147 in version 3.4
	- **OF28** D137 and D140 in version 1.1 -> D149 and D150 in version 3.4
	- **OF29** D142 to D146 in version 1.1 -> D153 to D158 in version 3.4
	- **OF30-OF32** D147 to D149 in version 1.1 -> D160 to D162 in version 3.4
	- **OF33-OF37** D150 to D154 in version 1.1 -> D164 to D168 in version 3.4
	- **OF38** In the **F sheet** of 1.1, **D272 to D275** are equal to **D171 to D174** in the **OF sheet** of 3.4.

   - **F Sheet**:
	- **F1** D5 to D8 in version 1.1 -> D10, D7, D8, D11 in version 3.4
	- **F2** D10 to D13 in version 1.1 -> D16, D14, D15, D17 in version 3.4
	- **F3** D15 to D20 in version 1.1 -> D20 to D25 in version 3.4
	- **F5** D22 to D29 in version 1.1 -> D34 to D41 in version 3.4
	- **F6** D31 to D34 in version 1.1 -> D44, D45, D47, D48 in version 3.4
	- **F7** D83 to D87 in version 1.1 -> D50 to D54 in version 3.4
	- **F8** D89 to D93 in version 1.1 -> D57 to D61 in version 3.4
	- **F9** D185 to D188 in version 1.1 -> D64 to D67 in version 3.4
	- **F10** D43 to D44 in version 1.1 -> D70 to D71 in version 3.4
	- **F11** D103 to D107 in version 1.1 -> D73 to D77 in version 3.4
	- **F12** D109 to D112 in version 1.1 -> D80 to D83 in version 3.4
	- **F13** D46 to D50 in version 1.1 -> D85 to D89 in version 3.4
	- **F14** D52 to D56 in version 1.1 -> D92 to D96 in version 3.4
	- **F15** D36 to D38 in version 1.1 -> D99 to D101 in version 3.4
	- **F16** D40 to D41 in version 1.1 -> D104 to D105 in version 3.4
	- **F17** D58 to D62 in version 1.1 -> D108 to D112 in version 3.4
	- **F18** D68 to D70 in version 1.1 -> D115 to D117 in version 3.4 (EXTRA OPTION IS IGNORED)
	- **F19** D78 to D81 in version 1.1 -> D119 to D122 in version 3.4
	- **F20** D64 to D66 in version 1.1 -> D125 to D127 in version 3.4
	- **F21** D72 to D76 in version 1.1 -> D129 to D133 in version 3.4
	- **F22** D135 to D139 in version 1.1 -> D136 to D140 in version 3.4
	- **F23** D123 to D127 in version 1.1 -> D143 to D147 in version 3.4
	- **F24** D116 to D121 in version 1.1 -> D150 to D155 in version 3.4
	- **F25** D129 to D133 in version 1.1 -> D159 to D163 in version 3.4
	- **F26-F27** D113 and D114 in version 1.1 -> D165 and D166 in version 3.4
	- **F28** D141 to D145 in version 1.1 -> D169 to D173 in version 3.4
	- **F29** D148 to D152 in version 1.1 -> D176 to D180 in version 3.4
	- **F30** D154 to D156 in version 1.1 -> D183 to D185 in version 3.4
	- **F31** D158 to D162 in version 1.1 -> D188 to D192 in version 3.4
	- **F32** D163 in version 1.1 -> D194 in version 3.4 (MIssing but assumed)
	- **F33** D165 to D170 in version 1.1 -> D196 to D201 in version 3.4
	- **F34** D172 to D177 in version 1.1 -> D203 to D208 in version 3.4
	- **F35** D179 to D183 in version 1.1 -> D211 to D215 in version 3.4
	- **F36** D190 to D192 in version 1.1 -> D218 to D220 in version 3.4
	- **F37** D193 in version 1.1 -> D222 in version 3.4
	- **F38-F39** D199 and D198 in version 1.1 -> D223 and D224 in version 3.4
	- **F40** D195 to D197 in version 1.1 -> D226 to D228 in version 3.4
	- **F41-F42** D210 and D211 in version 1.1 -> D230 and D231 in version 3.4
	- **F43** D201 to D205 in version 1.1 -> D233 to D237 in version 3.4
	- **F44** D207 to D209 in version 1.1 -> D240 to D242 in version 3.4
	- **F45** D213 to D217 in version 1.1 -> D245 to D249 in version 3.4
	- **F46** Missing
	- **F47** D228 to D230 in version 1.1 -> D252 to D254 in version 3.4
	- **F48** D232 to D234 in version 1.1 -> D256 to D258 in version 3.4
	- **F49** D236 to D239 in version 1.1 -> D261 to D264 in version 3.4
	- **F50** D241 to D245 in version 1.1 -> D267 to D271 in version 3.4
	- **F51** In the **OF sheet** of 1.1, **D43 and D44** are equal to **D273 and D274** in the **F sheet** of 3.4.
	- **F52** D250 to D253 in version 1.1 -> D276 to D280 in version 3.4
	- **F53** Missing
	- **F54** D256 to D261 in version 1.1 -> D287 to D292 in version 3.4
	- **F55** D263 to D267 in version 1.1 -> D295 to D298 in version 3.4
	- **F56** D268 to D270 in version 1.1 -> D301 to D303 in version 3.4
	- **F57** D277 to D279 in version 1.1 -> D305 to D307 in version 3.4
	- **F58** D281 to D286 in version 1.1 -> D309 to D314 in version 3.4
	- **F59** D288 to D291 in version 1.1 -> D317 to D320 in version 3.4
	- **F60-F62** D292 to D297 in version 1.1 -> D321 to D326 in version 3.4
	- **F62** D298 to D300 in version 1.1 -> D328 to D330 in version 3.4
	- **F63** D305 in version 1.1 -> D332 in version 3.4
	- **F64** Missing
	- **F65** D221 in version 1.1 -> D338 in version 3.4
	- **F65** D219 and D220 in version 1.1 -> D339 and D340 in version 3.4
	- **F66** D225 and D226 in version 1.1 -> D342 and D343 in version 3.4

   - **S Sheet**:
	- **S1** F13, F14, F16, F17 in version 1.1 -> F20, F21, F23, F24 in version 3.4
	- **S2** F89 to F91 in version 1.1 -> F35 to F38 in version 3.4
	- **S3** F102 to F104 in version 1.1 -> F48 to F50 in version 3.4
	- **S4** F119 to F122 in version 1.1 -> F65 to F68 in version 3.4
	- **S5** F137 to F140 in version 1.1 -> F83 to F86 in version 3.4

## 3. **Running the Script**
To process WESP-AC version 1.1 files:
1. Place the WESP-AC version 1.1 files in the input folder.
2. Run the script, and the processed files will be saved in a folder named `1.1_transferred`.
3. The script will copy the data as per the mappings outlined above.


This document provides an exhaustive overview of the mapping between versions 1.1 and 3.4. Any additional discrepancies between the versions have been addressed with logical mappings or appropriate substitutions as mentioned.
