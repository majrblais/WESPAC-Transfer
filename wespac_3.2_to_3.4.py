#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os
import xlwings as xw

def copy_wespac_values_3_2_to_3_4(old_file, new_file, output_file):
    """
    Copies data from a WESP-AC version 3.2 file to a WESP-AC version 3.4 file.
    
    Handles data shifts and special cases, such as row skipping or adjustment, between the two versions.
    """

    # Define the sheets, columns, and last rows to copy from the old to the new file
    sheets_to_copy = {
        'OF': {'column': 'D', 'last_row': 174},  # OF sheet data range
        'F': {'column': 'D', 'last_row': 345},   # F sheet data range
        'S': {'column': 'F', 'last_row': 101}    # S sheet data range
    }

    # Open the old and new WESP-AC files using xlwings
    app = xw.App(visible=False)
    old_wespac = app.books.open(old_file)
    new_wespac = app.books.open(new_file)

    try:
        # Loop through each sheet to transfer data
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

            # Handle specific version changes for each sheet

            # OF sheet: Adjust data from OF38 in rows 170-173 (3.2) to 171-174 (3.4)
            if sheet == 'OF':
                for row in range(1, 170):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row}').value = old_value

                for row in range(170, 174):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+1}').value = old_value  # Shift by 1 row in 3.4

            # F sheet: Copy data, skipping the new F68 field in row 345 (introduced in 3.4)
            elif sheet == 'F':
                for row in range(1, 225):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row}').value = old_value

                #F40
                for row in range(247, 250):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row-22}').value = old_value

                 #F41-F42
                new_sheet.range(f'{column}230').value = old_sheet.range(f'{column}225').value
                new_sheet.range(f'{column}231').value = old_sheet.range(f'{column}238').value

                #F43-F44
                for row in range(227, 237):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+6}').value = old_value          

                
                #F45-F46
                for row in range(240, 250):
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row+5}').value = old_value

                for row in range(252, 345):
                    if row == 345:
                        continue  # Skip F68 in 3.4
                    old_value = old_sheet.range(f'{column}{row}').value
                    new_sheet.range(f'{column}{row}').value = old_value
                    
                    
            # S sheet: Direct data copy as no specific changes between 3.2 and 3.4
            else:
                for row in range(1, last_row + 1):
                    old_value = old_sheet.range(f'{column}{row}').value
                    if old_value is not None:
                        new_sheet.range(f'{column}{row}').value = old_value

            # Re-protect the sheet after modifications
            new_sheet.api.Protect()

        # Save the new WESP-AC file with the copied values
        new_wespac.save(output_file)
        print(f"Values successfully copied from {old_file} to {output_file} (Version 3.2 to 3.4)")

    finally:
        # Close the workbooks and quit the app
        old_wespac.close()
        new_wespac.close()
        app.quit()

def process_wespac_folder_3_2():
    """
    Processes all WESP-AC version 3.2 files in the current directory, converting them to version 3.4.
    """

    # Define the folder containing the version 3.2 files
    folder_name = '3.2'
    transferred_folder = '3.2_transferred'
    
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

# Run the processing for version 3.2 to 3.4
process_wespac_folder_3_2()


# In[ ]:




