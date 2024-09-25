#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import xlwings as xw

def copy_wespac_values_3_2_to_3_4(old_file, new_file, output_file):
    """
    This function copies values from an old WESP-AC Excel file to a new one, while handling specific adjustments.
    """
    # Define the sheets and relevant data ranges to copy
    sheets_to_copy = {
        'OF': {'column': 'D', 'last_row': 174},
        'F': {'column': 'D', 'last_row': 345},
        'S': {'column': 'F', 'last_row': 101}
    }

    # Open both the old and new WESP-AC files using xlwings
    app = xw.App(visible=False)
    old_wespac = app.books.open(old_file)
    new_wespac = app.books.open(new_file)

    try:
        # Loop through each sheet and transfer data
        for sheet, info in sheets_to_copy.items():
            column = info['column']
            last_row = info['last_row']
            old_sheet = old_wespac.sheets[sheet]
            new_sheet = new_wespac.sheets[sheet]

            # Unprotect the new sheet in case it's protected
            password='empty'
            if password=='empty':
                try:
                    new_sheet.api.Unprotect()
                except Exception as e:
                    print(f"Failed to unprotect the sheet '{sheet}': {e}")
            else:
                new_sheet.api.Unprotect(password)

            # Copy data based on sheet
            if sheet == 'OF':
                # Copy values for OF sheet with specific adjustments and shifts
                for row in range(5, 109):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value

                new_sheet.range(f'{column}110').value = old_sheet.range(f'{column}110').value
                
                for row in range(111, 151):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value
                    
                #OF22 and OF23 switch
                for row in range(129, 132):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value
                for row in range(133, 136):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value
                    
                # Handle special case from D151-D154 in old sheet to D149-D150 in new
                for row in range(151, 155, 2):
                    old_value1 = old_sheet.range(f'{column}{row}').value
                    old_value2 = old_sheet.range(f'{column}{row+1}').value
                    new_value = 1 if old_value1 == 1 or old_value2 == 1 else 0
                    new_sheet.range(f'{column}{row-2}').value = new_value

                for row in range(156, 170):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-4}').value = old_value

                for row in range(170, 180):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value

            elif sheet == 'F':
                # Copy values for F sheet, with specific row shifting
                for row in range(5, 70):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row}').value = old_value
                
                # Set a specific row to 0 as part of the handling logic
                new_sheet.range(f'{column}123').value = 0

                # Switch F10-F11
                for row in range(70, 75):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value

                for row in range(77, 79):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-7}').value = old_value

                # Copy values for F sheet, with specific row shifting
                for row in range(79, 124):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row}').value = old_value
                
                # Shift rows by 1 from 124 onward
                for row in range(124, 192):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value

                # Handle further shifts by 2 from row 193 onwards
                for row in range(194, 222):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                new_sheet.range(f'{column}230').value = old_sheet.range(f'{column}240').value
                new_sheet.range(f'{column}231').value = old_sheet.range(f'{column}241').value

                # Handle specific row adjustments after row 222
                new_sheet.range(f'{column}224').value = old_sheet.range(f'{column}227').value
                
                for row in range(222, 227):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+3}').value = old_value

                for row in range(228, 240):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+4}').value = old_value

                new_sheet.range(f'{column}231').value = old_sheet.range(f'{column}240').value

                for row in range(242, 248):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+2}').value = old_value

                new_sheet.range(f'{column}250').value = old_sheet.range(f'{column}248').value

                for row in range(249, 254):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+88}').value = old_value

                new_sheet.range(f'{column}342').value = old_sheet.range(f'{column}254').value
                new_sheet.range(f'{column}343').value = old_sheet.range(f'{column}255').value

                for row in range(256, 289):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value

                # Handle special case for D289 and D290
                old_value1 = old_sheet.range(f'{column}289').value
                old_value2 = old_sheet.range(f'{column}290').value
                new_value = 1 if old_value1 == 1 or old_value2 == 1 else 0
                new_sheet.range(f'{column}284').value = new_value

                for row in range(291, 333):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-6}').value = old_value
                
                for row in range(333, 341):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-5}').value = old_value

            elif sheet == 'S':
                # Copy values for S sheet directly
                for row in range(6, 89):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row}').value = old_value

            # Protect the new sheet after changes
            new_sheet.api.Protect()

        # Save the modified WESP-AC file
        new_wespac.save(output_file)
        print(f"Values successfully copied from {old_file} to {output_file}")

    finally:
        # Close the workbooks and terminate the xlwings app
        old_wespac.close()
        new_wespac.close()
        app.quit()

def process_wespac_folder_3_1():
    """
    Processes all WESP-AC version 3.1 files in the current directory, converting them to version 3.4.
    """

    # Define the folder containing the version 3.2 files
    folder_name = '3.1'
    transferred_folder = '3.1_transferred'
    
    # Create the transferred folder if it doesn't exist
    if not os.path.exists(transferred_folder):
        os.makedirs(transferred_folder)

    # Get a list of all files in the version 3.2 folder
    input_folder = os.path.join(os.getcwd(), folder_name)
    wespac_files = [os.path.join(input_folder, file) for file in os.listdir(input_folder) if file.endswith('.xlsx')]

    # Define the new template WESP-AC version 3.4 file
    new_wespac_file = 'wespac_3.4.xlsx'

    # Iterate over all WESP-AC files in the folder
    for old_wespac_file in wespac_files:
        # Extract the filename without the extension and add the "_to_3.4" suffix
        base_name = os.path.splitext(os.path.basename(old_wespac_file))[0]
        output_wespac_file = os.path.join(transferred_folder, f'{base_name}_to_3.4.xlsx')
        
        # Copy values from the old version to the new one
        copy_wespac_values_3_2_to_3_4(old_wespac_file, new_wespac_file, output_wespac_file)
# Start the folder processing
process_wespac_folder_3_1()


# In[ ]:




