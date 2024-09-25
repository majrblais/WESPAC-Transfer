# WESPAC-Transfer
---
This repo is used to transfer old WESP-AC files (versions 1.1, 2.x, 3.0, 3.1, 3.2) to a newer version (either 3.3 or 3.4). The code requires an empty WESP-AC 3.3/3.4 version to be filled with data from older versions. It focuses on transferring the `OF`, `F`, and `S` sheets from the older versions into the newer ones.

The passwords for the protected sheets are required, and a password prompt will appear when running the scripts. Alternatively, the password can be modified directly in the scripts for easier use.

An extensive README for each version is included in their respective folders, along with example WESP-AC files to demonstrate the transfer process.



## Instructions for Running Scripts

1. **Prerequisites**:
    - Install required dependencies by running:
    
      pip install xlwings

2. **Transferring Data**:
    - Depending on the version you wish to transfer, use the appropriate script. The scripts will take data from the old version and transfer it to a new 3.4 version.
    - Run the script as follows:
    
      python wespac_<version>_to_3.4.py
    
      For example, to transfer data from version 1.1:
    
      python wespac_1.1_to_3.4.py

3. **Password Management**:
    - The script uses a default password (12345). The password is hardcoded and can be modified within the script if needed.
    - You can manually modify the password in the script in the section:
    
      password = "12345"

4. **Special Cases**:
    - Each version may have specific transfer rules and exceptions that have been handled in the code. Refer to the comments in the respective scripts for details on how data is mapped between versions.

5. **Output**:
    - The transferred data will be saved in a new Excel file in a folder named `version_transferred` (e.g., 1.1_transferred/ for version 1.1).

6. **Logging and Troubleshooting**:
    - If the script encounters issues, it will print error messages in the console. Ensure that the files are correctly named and located in their respective folders.
	
	
## Project Structure
```plaintext
WESPAC-Transfer/
│
├── 1.1/                       # Directory containing WESP-AC files in version 1.1
│   └── ...                    # WESP-AC 1.1 files to be processed
│
├── 2.x/                       # Directory containing WESP-AC files in version 2.x
│   └── ...                    # WESP-AC 2.x files to be processed
│
├── 3.0/                       # Directory containing WESP-AC files in version 3.0
│   └── ...                    # WESP-AC 3.0 files to be processed
│
├── 3.1/                       # Directory containing WESP-AC files in version 3.1
│   └── ...                    # WESP-AC 3.1 files to be processed
│
├── 3.2/                       # Directory containing WESP-AC files in version 3.2
│   └── ...                    # WESP-AC 3.2 files to be processed
│
├── README.md                  # Main README file
│
├── WESP-AC3.3.xlsx            # Empty WESP-AC version 3.3 template
├── WESP-AC3.4.xlsx            # Empty WESP-AC version 3.4 template
│
├── wespac_1.1_to_3.4.py       # Script to transfer data from version 1.1 to 3.4
├── wespac_2.x_to_3.4.py       # Script to transfer data from version 2.x to 3.4
├── wespac_3.0_to_3.4.py       # Script to transfer data from version 3.0 to 3.4
├── wespac_3.1_to_3.4.py       # Script to transfer data from version 3.1 to 3.4
└── wespac_3.2_to_3.4.py       # Script to transfer data from version 3.2 to 3.4
```