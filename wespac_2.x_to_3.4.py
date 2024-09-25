#!/usr/bin/env python
# coding: utf-8

# In[3]:


import os
import xlwings as xw

def copy_wespac_values(old_file, new_file, output_file, version='2.1'):
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
                # D4 to D15 in 2.1 -> D4 to D16 in 3.4
                for row in range(4, 16):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value

                # D16 to D38 in 2.1 -> D18 to D40 in 3.4
                for row in range(16, 39):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # D39 in 2.1 -> D42 in 3.4
                old_value = old_sheet.range(f'{column}39').value
                new_sheet.range(f'{column}42').value = old_value

                # D40 in 2.1 -> D44 in 3.4
                old_value = old_sheet.range(f'{column}40').value
                new_sheet.range(f'{column}44').value = old_value

                # D41 to D46 in 2.1 -> D46 to D51 in 3.4
                for row in range(41, 47):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value

                # D50 to D55 in 2.1 -> D53 to D58 in 3.4
                for row in range(50, 56):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value

                # D56 to D62 in 2.1 -> D65 to D71 in 3.4
                for row in range(56, 63):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+9}').value = old_value

                # D63 to D85 in 2.1 -> D72 to D96 in 3.4
                for row in range(63, 86):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+10}').value = old_value

                # D86 to D96 in 2.1 -> D97 to D108 in 3.4
                for row in range(86, 97):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+11}').value = old_value

                # D97 to D113 in 2.1 -> D109 to D125 in 3.4
                for row in range(97, 114):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+12}').value = old_value

                # D114 to D132 in 2.1 -> D127 to D145 in 3.4
                for row in range(114, 133):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+13}').value = old_value

                # D133, D135, D136, D140, and D141 to D144 in 2.1 -> D147 to D158 in 3.4
                for row in [133, 135, 136, 140, 141, 142, 143, 144]:
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+14}').value = old_value

                # D145 to D146 in 2.1 -> D160 to D174 in 3.4
                for row in range(145, 148):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+15}').value = old_value
                
                # D148 to D152 in 2.1 -> D164 to D168 in 3.4
                for row in range(148, 153):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+16}').value = old_value
                    
                # D154 to D158 in 2.1 -> D171 to D174 in 3.4
                for row in range(154, 158):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+17}').value = old_value

                old_sheet = old_wespac.sheets['F']  # Explicitly using the F sheet from the old file
                new_sheet = new_wespac.sheets['OF']  # Mapping to the OF sheet in the new file
                for row in range(304, 307):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row - 243}').value = old_value


            elif sheet == 'F':
                # F sheet mappings
                # D5 to D10 in 2.1 -> D6 to D11 in 3.4
                for row in range(5, 11):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value

                # D13 to D16 in 2.1 -> D14 to D17 in 3.4
                for row in range(13, 17):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value

                # D18 to D23 in 2.1 -> D20 to D25 in 3.4
                for row in range(18, 24):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # D26 and D27 in 2.1 -> D30 and D31 in 3.4
                for row in range(26, 28):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+4}').value = old_value

                # D29 to D43 in 2.1 -> D34 to D48 in 3.4
                for row in range(29, 44):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value

                # D89 to D93 in 2.1 -> D50 to D54 in 3.4
                for row in range(89, 94):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-39}').value = old_value

                # D95 to D99 in 2.1 -> D57 to D61 in 3.4
                for row in range(95, 100):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-38}').value = old_value

                # D191 to D194 in 2.1 -> D64 to D67 in 3.4
                for row in range(191, 195):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-127}').value = old_value

                # D106 and D107 in 2.1 -> D70 and D71 in 3.4
                for row in range(106, 108):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-36}').value = old_value

                # D109 to D113 in 2.1 -> D73 to D77 in 3.4
                for row in range(109, 114):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-36}').value = old_value

                # D115 to D118 in 2.1 -> D80 to D83 in 3.4
                for row in range(115, 119):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-35}').value = old_value

                                # D52 to D56 in 2.1 -> D85 to D89 in 3.4
                for row in range(52, 57):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+33}').value = old_value

                # D58 to D62 in 2.1 -> D92 to D96 in 3.4
                for row in range(58, 63):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+34}').value = old_value

                # D45 to D47 in 2.1 -> D99 to D101 in 3.4
                for row in range(45, 48):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+54}').value = old_value

                # D49 to D50 in 2.1 -> D104 to D105 in 3.4
                for row in range(49, 51):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+55}').value = old_value

                # D64 to D68 in 2.1 -> D108 to D112 in 3.4
                for row in range(64, 69):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+44}').value = old_value

                # D74 to D76 in 2.1 -> D115 to D117 in 3.4
                for row in range(74, 77):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+41}').value = old_value

                # D84 to D87 in 2.1 -> D119 to D122 in 3.4
                for row in range(84, 88):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+35}').value = old_value

                # D70 to D72 in 2.1 -> D125 to D127 in 3.4
                for row in range(70, 73):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+55}').value = old_value

                # D78 to D82 in 2.1 -> D129 to D133 in 3.4
                for row in range(78, 83):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+51}').value = old_value

                # D141 to D145 in 2.1 -> D136 to D140 in 3.4
                for row in range(141, 146):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value

                # D129 to D133 in 2.1 -> D143 to D147 in 3.4
                for row in range(129, 134):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+14}').value = old_value

                # D122 to D128 in 2.1 -> D150 to D156 in 3.4
                for row in range(122, 129):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D135 to D139 in 2.1 -> D159 to D163 in 3.4
                for row in range(135, 140):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+24}').value = old_value

                # D119 and D120 in 2.1 -> D165 and D166 in 3.4
                for row in range(119, 121):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+46}').value = old_value

                # D147 to D151 in 2.1 -> D169 to D173 in 3.4
                for row in range(147, 152):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+22}').value = old_value

                # D154 to D158 in 2.1 -> D176 to D180 in 3.4
                for row in range(154, 159):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+22}').value = old_value

                # D160 to D162 in 2.1 -> D183 to D185 in 3.4
                for row in range(160, 163):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+23}').value = old_value

                # D164 to D168 in 2.1 -> D188 to D192 in 3.4
                for row in range(164, 169):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+24}').value = old_value

                # D171 to D176 in 2.1 -> D196 to D201 in 3.4
                for row in range(171, 177):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+25}').value = old_value

                # D178 to D183 in 2.1 -> D203 to D208 in 3.4
                for row in range(178, 184):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+25}').value = old_value

                # D185 to D189 in 2.1 -> D211 to D215 in 3.4
                for row in range(185, 190):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+26}').value = old_value

                # D196 to D198 in 2.1 -> D218 to D220 in 3.4
                for row in range(196, 199):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+22}').value = old_value

                # D199 is equal to D222
                new_sheet.range(f'{column}222').value = old_sheet.range(f'{column}199').value

                # D205 is equal to D223
                new_sheet.range(f'{column}223').value = old_sheet.range(f'{column}205').value

                # D204 is equal to D224
                new_sheet.range(f'{column}224').value = old_sheet.range(f'{column}204').value

                # D201 to D203 in 2.1 -> D226 to D228 in 3.4
                for row in range(201, 204):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+25}').value = old_value

                 # D213 to D215 in 2.1 -> D240 to D242 in 3.4
                for row in range(213, 215):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+27}').value = old_value

                 # D213 to D215 in 2.1 -> D240 to D242 in 3.4
                for row in range(219, 223):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+26}').value = old_value

                # D216 to D217 in 2.1 -> D230 to D231 in 3.4
                for row in range(216, 218):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+14}').value = old_value

                # D219 to D223 in 2.1 -> D233 to D237 in 3.4
                for row in range(219, 224):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+14}').value = old_value
                    

                
                    
                # D234 to D240 in 2.1 -> D252 to D258 in 3.4
                for row in range(234, 241):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+18}').value = old_value

                # D242 to D245 in 2.1 -> D261 to D264 in 3.4
                for row in range(242, 246):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+19}').value = old_value

                # D248 to D260 in 2.1 -> D267 to D279 in 3.4
                for row in range(248, 261):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+19}').value = old_value

                # D263 to D268 in 2.1 -> D287 to D292 in 3.4
                for row in range(263, 269):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+24}').value = old_value

                # D270 to D273 in 2.1 -> D295 to D298 in 3.4
                for row in range(270, 274):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+25}').value = old_value

                # D275 to D288 in 2.1 -> D301 to D314 in 3.4
                for row in range(275, 289):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+26}').value = old_value

                # D290 to D299 in 2.1 -> D317 to D326 in 3.4
                for row in range(290, 300):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+27}').value = old_value

                # D300 to D302 in 2.1 -> D328 to D330 in 3.4
                for row in range(300, 303):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+28}').value = old_value

                # D307 in 2.1 -> D332 in 3.4
                new_sheet.range(f'{column}332').value = old_sheet.range(f'{column}307').value

                # D225 to D230 in 2.1 -> D338 to D343 in 3.4
                for row in range(225, 231):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+113}').value = old_value

            elif sheet == 'S':
                # S sheet mappings
                # F18 to F22 in 2.1 -> F20 to F24 in 3.4
                for row in range(18, 23):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # F33 to F35 in 2.1 -> F35 to F37 in 3.4
                for row in range(33, 36):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # F46 to F48 in 2.1 -> F48 to F50 in 3.4
                for row in range(46, 49):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # F63 to F66 in 2.1 -> F65 to F68 in 3.4
                for row in range(63, 67):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # F81 to F84 in 2.1 -> F83 to F86 in 3.4
                for row in range(81, 85):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value
            # Protect the sheet again after modifying
            new_sheet.api.Protect()

        # Save the new WESP-AC file with the copied values
        new_wespac.save(output_file)
        print(f"Values successfully copied from {old_file} to {output_file} (Version: {version})")

    finally:
        # Close both workbooks and quit the app
        old_wespac.close()
        new_wespac.close()
        app.quit()

def process_wespac_folder(version='2.1'):
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
x_var = '2.x'
process_wespac_folder(version=x_var)


# In[ ]:





# In[ ]:




