#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import xlwings as xw

def copy_wespac_values(old_file, new_file, output_file, version='3.0'):
    # Define the sheets, columns, and last rows to copy from the old to the new file
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
            last_row = info['last_row']
            
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

            # Handle version 3.0 to 3.4 for each sheet
            if sheet == 'OF':
                # Data shifts and special handling for the OF sheet
                for row in range(5, 109):
                    if row <= 108:
                        # D5 in 3.1 -> D6 in 3.4 (shift by 1 up until D108)
                        old_value = old_sheet.range(f'{column}{row}').value
                        new_sheet.range(f'{column}{row+1}').value = old_value
                    elif row == 109:
                        # D109 in 3.1 is empty, skip
                        continue

                # D110 to D111 in 3.1 -> D110 to D111 in 3.4 (direct mapping)
                new_sheet.range(f'{column}110').value = old_sheet.range(f'{column}110').value

                # D112 in 3.4 corresponds to D111 in 3.1 (shift starts here)
                for row in range(111, 151):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-1}').value = old_value

                #OF22 and OF23 switch
                for row in range(129, 132):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value
                for row in range(133, 136):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value
                
                # Special case: from D151-D154 in 3.1 into D149-D150 in 3.4
                for row in range(151, 155, 2):
                    old_value1 = old_sheet.range(f'{column}{row}').value
                    old_value2 = old_sheet.range(f'{column}{row+1}').value
                    new_value = 1 if old_value1 == 1 or old_value2 == 1 else 0
                    new_sheet.range(f'{column}{row-2}').value = new_value  # D149-D150 in 3.4

                # Align from D156 onward until D169 (empty in 3.1)
                for row in range(156, 170):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-4}').value = old_value

                # D170 to D179 in 3.1 -> D165 to D174 in 3.4
                for row in range(170, 180):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value
            elif sheet == 'F':
                # Data shifts and special handling for the F sheet (version 3.0 to 3.4)
                # D3 to D52 in 3.0 -> D5 to D54 in 3.4
                for row in range(3, 53):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                # Skip D55 in 3.4
                # D53 to D93 in 3.0 -> D56 to D96 in 3.4
                for row in range(53, 67):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value
                    
                # Switch F10-F11
                for row in range(67, 71):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+6}').value = old_value

                for row in range(74, 76):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-4}').value = old_value

                
                # Skip D55 in 3.4
                # D53 to D93 in 3.0 -> D56 to D96 in 3.4
                for row in range(76, 96):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value

                
                # Skip D97 in 3.4
                # D94 to D118 in 3.0 -> D98 to D122 in 3.4
                for row in range(94, 119):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+4}').value = old_value

                # Skip D123 in 3.4
                # D119 to D129 in 3.0 -> D124 to D134 in 3.4
                for row in range(119, 130):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value

                # Extra row (130) in 3.0, skip it
                # D131 to D189 in 3.0 -> D135 to D193 in 3.4
                for row in range(131, 190):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+4}').value = old_value

                # Special case for D194 in 3.4 (set to 1 if D191 in 3.0 is 0, else 0)
                old_value = old_sheet.range(f'{column}191').value
                new_value = 1 if old_value == 0 else 0
                new_sheet.range(f'{column}194').value = new_value

                # D190 to D218 in 3.0 -> D195 to D223 in 3.4
                for row in range(190, 219):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value

                # Specific mappings
                new_sheet.range(f'{column}224').value = old_sheet.range(f'{column}224').value
                for row in range(219, 224):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+6}').value = old_value

                for row in range(225, 237):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+7}').value = old_value

                new_sheet.range(f'{column}230').value = old_sheet.range(f'{column}237').value
                new_sheet.range(f'{column}231').value = old_sheet.range(f'{column}238').value

                for row in range(239, 245):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value
                    
                for row in range(246, 252):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+92}').value = old_value
                    
                # Special: Set D250 in 3.4 to 0
                new_sheet.range(f'{column}250').value = 0

                # D252 to D284 in 3.0 -> D251 to D283 in 3.4
                for row in range(252, 285):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-1}').value = old_value

                # Special case for D285 or D286 -> D284 in 3.4
                old_value1 = old_sheet.range(f'{column}285').value
                old_value2 = old_sheet.range(f'{column}286').value
                new_value = 1 if old_value1 == 1 or old_value2 == 1 else 0
                new_sheet.range(f'{column}284').value = new_value

                # D287 to D328 in 3.0 -> D285 to D326 in 3.4
                for row in range(287, 329):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-2}').value = old_value

                # Extra row D327 in 3.4, set to 0
                new_sheet.range(f'{column}327').value = 0

                # D329 to D337 in 3.0 -> D328 to D336 in 3.4
                for row in range(329, 338):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-1}').value = old_value



                # D250 and D251 in 3.0 -> D342 and D343 in 3.4, set 344 and 345 to 0
                new_sheet.range(f'{column}342').value = old_sheet.range(f'{column}250').value
                new_sheet.range(f'{column}343').value = old_sheet.range(f'{column}251').value
                new_sheet.range(f'{column}344').value = 0
                new_sheet.range(f'{column}345').value = 0




            elif sheet == 'S':
                # The data starts at F3 in 3.0 and F6 in 3.4, ends at F85 in 3.0 and F88 in 3.4
                for row in range(3, 86):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value

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

def process_wespac_folder(version='3.0'):
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
x_var = '3.0'
process_wespac_folder(version=x_var)


# In[ ]:





# In[ ]:




