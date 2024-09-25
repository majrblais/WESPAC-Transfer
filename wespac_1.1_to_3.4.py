#!/usr/bin/env python
# coding: utf-8

# In[3]:


import os
import xlwings as xw

def copy_wespac_values(old_file, new_file, output_file, version='1.1'):
    # Define the sheets and their corresponding data columns
    sheets_to_copy = {
        'OF': {'column': 'D', 'last_row': 174},
        'F': {'column': 'D', 'last_row': 345},
        'S': {'column': 'F', 'last_row': 101}
    }

    # Open the old and new WESP-AC files using xlwings
    app = xw.App(visible=False)
    old_wespac = app.books.open(old_file)
    new_wespac = app.books.open(new_file)

    try:
        for sheet, info in sheets_to_copy.items():
            column = info['column']
            old_sheet = old_wespac.sheets[sheet]
            new_sheet = new_wespac.sheets[sheet]

            # Unprotect the sheet if it's protected

            password='empty'
            if password=='empty':
                try:
                    new_sheet.api.Unprotect()
                except Exception as e:
                    print(f"Failed to unprotect the sheet '{sheet}': {e}")
            else:
                new_sheet.api.Unprotect(password)
            if sheet == 'OF':
                # OF sheet mappings

                # D12 to D17 in 1.1 -> D19 to D25 in 3.4
                for row in range(12, 18):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+7}').value = old_value

                # D19 to D25 in 1.1 -> D26 to D32 in 3.4
                for row in range(19, 26):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+7}').value = old_value

                # D19 to D25 in 1.1 -> D34 to D40 in 3.4
                for row in range(19, 26):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+15}').value = old_value

                # D34 in 1.1 -> D42 in 3.4
                new_sheet.range(f'{column}42').value = old_sheet.range(f'{column}34').value

                # D35 in 1.1 -> D44 in 3.4
                new_sheet.range(f'{column}44').value = old_sheet.range(f'{column}35').value

                # D37 to D41 in 1.1 -> D47 to D51 in 3.4
                for row in range(37, 42):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+10}').value = old_value

                # D46 to D50 in 1.1 -> D54 to D58 in 3.4
                for row in range(46, 51):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+8}').value = old_value

                # D52 to D57 in 1.1 -> D66 to D71 in 3.4
                for row in range(52, 58):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+14}').value = old_value

                # D58 in 1.1 -> D73 in 3.4
                new_sheet.range(f'{column}73').value = old_sheet.range(f'{column}58').value

                # D60 to D66 in 1.1 -> D75 to D81 in 3.4
                for row in range(60, 67):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+15}').value = old_value

                # D68 to D73 in 1.1 -> D83 to D88 in 3.4
                for row in range(68, 74):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+15}').value = old_value

                # D75 to D80 in 1.1 -> D90 to D95 in 3.4
                for row in range(75, 81):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+15}').value = old_value

                # D82 to D86 in 1.1 -> D98 to D102 in 3.4
                for row in range(82, 87):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+16}').value = old_value

                # D87 and D88 special handling -> D104 to D107 in 3.4
                if old_sheet.range(f'{column}87').value == 1 and old_sheet.range(f'{column}88').value == 1:
                    new_sheet.range(f'{column}105').value = 1
                elif old_sheet.range(f'{column}87').value == 1 and old_sheet.range(f'{column}88').value == 0:
                    new_sheet.range(f'{column}104').value = 1
                    new_sheet.range(f'{column}105').value = 0
                elif old_sheet.range(f'{column}87').value == 0 and old_sheet.range(f'{column}88').value == 1:
                    new_sheet.range(f'{column}106').value = 1
                else:
                    new_sheet.range(f'{column}107').value = 1

                # D90, D91, D92 -> D109 in 3.4 as 0, 0.33, or 0.66 respectively
                if old_sheet.range(f'{column}90').value == 1:
                    new_sheet.range(f'{column}109').value = 0
                elif old_sheet.range(f'{column}91').value == 1:
                    new_sheet.range(f'{column}109').value = 0.33
                elif old_sheet.range(f'{column}92').value == 1:
                    new_sheet.range(f'{column}109').value = 0.66

                # D93 in 1.1 -> D110 in 3.4
                new_sheet.range(f'{column}110').value = old_sheet.range(f'{column}93').value

                # D96 to D99 in 1.1 -> D112 to D116 in 3.4
                for row in range(96, 100):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+16}').value = old_value

                # D101 to D105 in 1.1 -> D117 to D120 in 3.4
                for row in range(101, 106):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+16}').value = old_value

                # D106 to D109 in 1.1 -> D122 to D125 in 3.4
                for row in range(106, 110):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+16}').value = old_value

                # D111 to D113 in 1.1 -> D128 to D131 in 3.4
                for row in range(111, 114):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+17}').value = old_value

                # D115 to D117 in 1.1 -> D132 to D134 in 3.4
                for row in range(115, 118):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+17}').value = old_value

                # D119 to D122 in 1.1 -> D136 to D138 in 3.4
                for row in range(119, 123):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+17}').value = old_value

                # D123 to D128 in 1.1 -> D140 to D145 in 3.4
                for row in range(123, 129):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+17}').value = old_value

                # Special case: D130, D131, D132, D133, D134, D135 (1.1) -> D147 in 3.4 (900, 1100, etc.)
                if old_sheet.range(f'{column}130').value == 1:
                    new_sheet.range(f'{column}147').value = 900
                elif old_sheet.range(f'{column}131').value == 1:
                    new_sheet.range(f'{column}147').value = 1100
                elif old_sheet.range(f'{column}132').value == 1:
                    new_sheet.range(f'{column}147').value = 1300
                elif old_sheet.range(f'{column}133').value == 1:
                    new_sheet.range(f'{column}147').value = 1500
                elif old_sheet.range(f'{column}134').value == 1:
                    new_sheet.range(f'{column}147').value = 1700
                elif old_sheet.range(f'{column}135').value == 1:
                    new_sheet.range(f'{column}147').value = 1900

                # D137 and D140 in 1.1 -> D149 and D150 in 3.4
                new_sheet.range(f'{column}149').value = old_sheet.range(f'{column}137').value
                new_sheet.range(f'{column}150').value = old_sheet.range(f'{column}140').value

                # D142 to D146 in 1.1 -> D153 to D158 in 3.4 (except D154)
                for row in range(142, 147):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+11}').value = old_value

                # D147 to D149 in 1.1 -> D160 to D162 in 3.4
                for row in range(147, 150):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+13}').value = old_value

                # D150 to D154 in 1.1 -> D164 to D168 in 3.4
                for row in range(150, 155):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+14}').value = old_value

            elif sheet == 'F':
                # F sheet mappings

                # D5, D6, D7, D8 in 1.1 -> D10, D7, D8, D11 in 3.4
                new_sheet.range(f'{column}10').value = old_sheet.range(f'{column}5').value
                new_sheet.range(f'{column}7').value = old_sheet.range(f'{column}6').value
                new_sheet.range(f'{column}8').value = old_sheet.range(f'{column}7').value
                new_sheet.range(f'{column}11').value = old_sheet.range(f'{column}8').value

                # D10, D11, D12, D13 in 1.1 -> D16, D14, D15, D17 in 3.4
                new_sheet.range(f'{column}16').value = old_sheet.range(f'{column}10').value
                new_sheet.range(f'{column}14').value = old_sheet.range(f'{column}11').value
                new_sheet.range(f'{column}15').value = old_sheet.range(f'{column}12').value
                new_sheet.range(f'{column}17').value = old_sheet.range(f'{column}13').value

                # D15 to D20 in 1.1 -> D20 to D25 in 3.4
                for row in range(15, 21):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value

                # D22 to D29 in 1.1 -> D34 to D41 in 3.4
                for row in range(22, 30):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+12}').value = old_value

                # D31 to D34 in 1.1 -> D44, D45, D47, D48 in 3.4
                new_sheet.range(f'{column}44').value = old_sheet.range(f'{column}31').value
                new_sheet.range(f'{column}45').value = old_sheet.range(f'{column}32').value
                new_sheet.range(f'{column}47').value = old_sheet.range(f'{column}33').value
                new_sheet.range(f'{column}48').value = old_sheet.range(f'{column}34').value

                # D83 to D87 in 1.1 -> D50 to D54 in 3.4
                for row in range(83, 88):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-33}').value = old_value

                # D89 to D93 in 1.1 -> D57 to D61 in 3.4
                for row in range(89, 94):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-32}').value = old_value

                # D185 to D188 in 1.1 -> D64 to D67 in 3.4
                for row in range(185, 189):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-121}').value = old_value

                # D43 to D44 in 1.1 -> D70 to D71 in 3.4
                new_sheet.range(f'{column}70').value = old_sheet.range(f'{column}43').value
                new_sheet.range(f'{column}71').value = old_sheet.range(f'{column}44').value

                # D103 to D107 in 1.1 -> D73 to D77 in 3.4
                for row in range(103, 108):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-30}').value = old_value

                # D109 to D112 in 1.1 -> D80 to D83 in 3.4
                for row in range(109, 113):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-29}').value = old_value

                # D46 to D50 in 1.1 -> D85 to D89 in 3.4
                for row in range(46, 51):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+39}').value = old_value

                # D52 to D56 in 1.1 -> D92 to D96 in 3.4
                for row in range(52, 57):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+40}').value = old_value

                # D36 to D38 in 1.1 -> D99 to D101 in 3.4
                for row in range(36, 39):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+63}').value = old_value


                for row in range(40, 42):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+64}').value = old_value

                # D58 to D62 in 1.1 -> D108 to D112 in 3.4
                for row in range(58, 63):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+50}').value = old_value

                # D68 to D70 in 1.1 -> D115 to D117 in 3.4
                for row in range(68, 71):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+47}').value = old_value

                # D78 to D81 in 1.1 -> D119 to D122 in 3.4
                for row in range(78, 82):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+41}').value = old_value

                # D64 to D66 in 1.1 -> D125 to D127 in 3.4
                for row in range(64, 67):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+61}').value = old_value

                # D72 to D76 in 1.1 -> D129 to D133 in 3.4
                for row in range(72, 77):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+57}').value = old_value

                # D135 to D139 in 1.1 -> D136 to D140 in 3.4
                for row in range(135, 140):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value

                # D123 to D127 in 1.1 -> D143 to D147 in 3.4
                for row in range(123, 128):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+20}').value = old_value

                # D116 to D121 in 1.1 -> D150 to D155 in 3.4
                for row in range(116, 122):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+34}').value = old_value

                # D129 to D133 in 1.1 -> D159 to D163 in 3.4
                for row in range(129, 134):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+30}').value = old_value

                # D113 and D114 in 1.1 -> D165 and D166 in 3.4
                new_sheet.range(f'{column}165').value = old_sheet.range(f'{column}113').value
                new_sheet.range(f'{column}166').value = old_sheet.range(f'{column}114').value

                # D141 to D145 in 1.1 -> D169 to D173 in 3.4
                for row in range(141, 146):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D148 to D152 in 1.1 -> D176 to D180 in 3.4
                for row in range(148, 153):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D154 to D156 in 1.1 -> D183 to D185 in 3.4
                for row in range(154, 157):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+29}').value = old_value

                # D158 to D162 in 1.1 -> D188 to D192 in 3.4
                for row in range(158, 163):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+30}').value = old_value

                # D163 in 1.1 -> D194 in 3.4
                new_sheet.range(f'{column}194').value = old_sheet.range(f'{column}163').value

                # D165 to D170 in 1.1 -> D196 to D201 in 3.4
                for row in range(165, 171):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+31}').value = old_value

                # D172 to D177 in 1.1 -> D203 to D208 in 3.4
                for row in range(172, 178):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+31}').value = old_value

                # D179 to D183 in 1.1 -> D211 to D215 in 3.4
                for row in range(179, 184):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+32}').value = old_value

                # D190 to D192 in 1.1 -> D218 to D220 in 3.4
                for row in range(190, 193):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D193 in 1.1 -> D222 in 3.4
                new_sheet.range(f'{column}222').value = old_sheet.range(f'{column}193').value

                # D195 to D197 in 1.1 -> D226 to D228 in 3.4
                for row in range(195, 198):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+31}').value = old_value

                # D199 and D198 in 1.1 -> D223 and D224 in 3.4
                new_sheet.range(f'{column}223').value = old_sheet.range(f'{column}199').value
                new_sheet.range(f'{column}224').value = old_sheet.range(f'{column}198').value

                # D201 to D205 in 1.1 -> D233 to D237 in 3.4
                for row in range(201, 206):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+32}').value = old_value

                # D207 to D209 in 1.1 -> D240 to D242 in 3.4
                for row in range(207, 210):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+33}').value = old_value

                # D210 and D211 in 1.1 -> D230 and D231 in 3.4
                new_sheet.range(f'{column}230').value = old_sheet.range(f'{column}210').value
                new_sheet.range(f'{column}231').value = old_sheet.range(f'{column}211').value

                # D213 to D217 in 1.1 -> D245 to D249 in 3.4
                for row in range(213, 218):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+32}').value = old_value

                # D228 to D230 in 1.1 -> D252 to D254 in 3.4
                for row in range(228, 231):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+24}').value = old_value

                # D232 to D234 in 1.1 -> D256 to D258 in 3.4
                for row in range(232, 235):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+24}').value = old_value

                # D236 to D239 in 1.1 -> D261 to D264 in 3.4
                for row in range(236, 240):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+25}').value = old_value

                # D241 to D245 in 1.1 -> D267 to D271 in 3.4
                for row in range(241, 246):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+26}').value = old_value

                # D250 to D253 in 1.1 -> D276 to D280 in 3.4
                for row in range(250, 254):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+26}').value = old_value

                # D256 to D261 in 1.1 -> D287 to D292 in 3.4
                for row in range(256, 262):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+31}').value = old_value

                # D263 to D267 in 1.1 -> D295 to D298 in 3.4
                for row in range(263, 268):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+32}').value = old_value

                # D268 to D270 in 1.1 -> D301 to D303 in 3.4
                for row in range(268, 271):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+33}').value = old_value

                # D277 to D279 in 1.1 -> D305 to D307 in 3.4
                for row in range(277, 280):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D281 to D286 in 1.1 -> D309 to D314 in 3.4
                for row in range(281, 287):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D288 to D291 in 1.1 -> D317 to D320 in 3.4
                for row in range(288, 292):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+29}').value = old_value

                # D292 to D297 in 1.1 -> D321 to D326 in 3.4
                for row in range(292, 298):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+29}').value = old_value

                # D298 to D300 in 1.1 -> D328 to D330 in 3.4
                for row in range(298, 301):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+30}').value = old_value

                # D305 in 1.1 -> D332 in 3.4
                new_sheet.range(f'{column}332').value = old_sheet.range(f'{column}305').value

                # D221 in 1.1 -> D338 in 3.4
                new_sheet.range(f'{column}338').value = old_sheet.range(f'{column}221').value

                # D219 and D220 in 1.1 -> D339 and D340 in 3.4
                new_sheet.range(f'{column}339').value = old_sheet.range(f'{column}219').value
                new_sheet.range(f'{column}340').value = old_sheet.range(f'{column}220').value

                # D225 and D226 in 1.1 -> D342 and D343 in 3.4
                new_sheet.range(f'{column}342').value = old_sheet.range(f'{column}225').value
                new_sheet.range(f'{column}343').value = old_sheet.range(f'{column}226').value

            elif sheet == 'S':
                # S sheet mappings

                # F13, F14, F16, F17 in 1.1 -> F20, F21, F23, F24 in 3.4
                new_sheet.range(f'F20').value = old_sheet.range(f'F13').value
                new_sheet.range(f'F21').value = old_sheet.range(f'F14').value
                new_sheet.range(f'F23').value = old_sheet.range(f'F16').value
                new_sheet.range(f'F24').value = old_sheet.range(f'F17').value

                # F89 to F91 in 1.1 -> F35 to F38 in 3.4
                for row in range(89, 92):
                    old_value = old_sheet.range(f'F{row}').value
                    new_sheet.range(f'F{row-54}').value = old_value

                # F102 to F104 in 1.1 -> F48 to F50 in 3.4
                for row in range(102, 105):
                    old_value = old_sheet.range(f'F{row}').value
                    new_sheet.range(f'F{row-54}').value = old_value

                # F119 to F122 in 1.1 -> F65 to F68 in 3.4
                for row in range(119, 123):
                    old_value = old_sheet.range(f'F{row}').value
                    new_sheet.range(f'F{row-54}').value = old_value

                # F137 to F140 in 1.1 -> F83 to F86 in 3.4
                for row in range(137, 141):
                    old_value = old_sheet.range(f'F{row}').value
                    new_sheet.range(f'F{row-54}').value = old_value

            # Protect the sheet again after modifying
            new_sheet.api.Protect()

        
        # Special case 1: In the F sheet of 1.1, D302 to D304 are equal to D61 to D63 in the OF sheet of 3.4.
        old_f_sheet = old_wespac.sheets['F']
        new_of_sheet = new_wespac.sheets['OF']
        for row in range(302, 305):
            old_value = old_f_sheet.range(f'D{row}').value
            new_of_sheet.range(f'D{row-241}').value = old_value  # D61 to D63 in OF sheet

        # Special case 2: In the OF sheet of 1.1, D43 and D44 are equal to D273 and D274 in the F sheet of 3.4.
        old_of_sheet = old_wespac.sheets['OF']
        new_f_sheet = new_wespac.sheets['F']
        for row in range(43, 45):
            old_value = old_of_sheet.range(f'D{row}').value
            new_f_sheet.range(f'D{row+230}').value = old_value  # D273 and D274 in F sheet

        # Special case 3: In the F sheet of 1.1, D272 to D275 are equal to D171 to D174 in the OF sheet of 3.4.
        for row in range(272, 276):
            old_value = old_f_sheet.range(f'D{row}').value
            new_of_sheet.range(f'D{row-101}').value = old_value  # D171 to D174 in OF sheet

        # Protect the OF and F sheets again after the special cases
        new_of_sheet.api.Protect()
        new_f_sheet.api.Protect()

        # Save the new WESP-AC file with the copied values
        new_wespac.save(output_file)
        print(f"Values successfully copied from {old_file} to {output_file} (Version: {version})")


    finally:
        # Close both workbooks and quit the app
        old_wespac.close()
        new_wespac.close()
        app.quit()
def process_wespac_folder(version='1.1'):
    # Define the folder name based on version
    folder_name = f'{version}'
    transferred_folder = f'{version}_transferred'

    # Create the transferred folder if it doesn't exist
    if not os.path.exists(transferred_folder):
        os.makedirs(transferred_folder)

    # Get a list of all files in the version folder
    input_folder = os.path.join(os.getcwd(), folder_name)  # Assuming folder is in the current directory
    wespac_files = [os.path.join(input_folder, file) for file in os.listdir(input_folder) if file.endswith('.xlsx')]

    # Define the new template WESP-AC file
    new_wespac_file = 'wespac_3.4.xlsx'

    # Iterate over all WESP-AC files in the folder
    for old_wespac_file in wespac_files:
        # Extract the filename without the extension and add the _3.4 suffix
        base_name = os.path.splitext(os.path.basename(old_wespac_file))[0]
        output_wespac_file = os.path.join(transferred_folder, f'{base_name}_3.4.xlsx')

        # Copy values from the old version to the new one
        copy_wespac_values(old_wespac_file, new_wespac_file, output_wespac_file, version=version)

# Define the version you want to process
x_var = '1.1'
process_wespac_folder(version=x_var)


# In[ ]:





# In[ ]:




