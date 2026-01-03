# LabVIEW Bindings for libxlsxwriter

This package provides LabVIEW bindings for libxlsxwriter. It is distributed separately from the main libxlsxwriter source code and should be extracted into the libxlsxwriter source directory.

## Installation

Extract the LabVIEW bindings package into your libxlsxwriter source directory:

```
libxlsxwriter-x.x.x/
├── LabView/                           <- Package contents
│   ├── Readme.md
│   ├── libxlsxwriter_LV.h             <- LabVIEW-compatible header
│   ├── Example_template.vi            <- Template VI for examples
│   ├── parse_c_examples.py            <- C example parser script
│   └── generate_all_examples.py       <- VI generator script
├── src/
│   ├── labview_wrappers.c             <- LabVIEW wrapper functions
│   └── ...
├── sandbox/
│   ├── install_dev_environment.wsb    <- Windows Sandbox config
│   ├── README.txt                     <- Sandbox documentation
│   └── shared/
│       ├── build.bat                  <- Library build script
│       └── install_all.bat            <- Dev environment installer
├── examples/
└── ...
```

---

## Development Workflow

### 1. Set Up Windows Development Environment

**Prerequisites (not included - obtain separately):**
- Visual Studio VSLayout directory
- LabVIEW Software Bundle ISO (for LabVIEW)

**Setup:**
1. Place the prerequisites in `sandbox/shared/`:
   - `VSLayout/` directory
   - `LabVIEW Software Bundle 1.iso`

2. Double-click `sandbox/install_dev_environment.wsb`

This opens a Windows Sandbox instance and runs `install_all.bat` which installs Visual Studio and LabVIEW.

### 2. Build the Library

Inside the Windows Sandbox, double-click `shared/build.bat`

This performs:
- Git checkout of zlib
- Copies the libxlsxwriter directory from the shared directory
- Builds win32 and x64 libraries

### 3. Create LabVIEW SubVIs

1. Copy the 32-bit or 64-bit `libxlsxwriter.dll` and `libxlsxwriter_LV.h` to a development directory

2. Open LabVIEW and launch the Import Shared Library Wizard:
   - **Tools > Import > Shared Library (.dll)...**

3. Configure the wizard:
   - Select the DLL and header file
   - Leave preprocessor definitions empty
   - Check all functions to convert
   - Select the output directory
   - Select "Simple error handling"

4. Wait for the exported function prototypes to populate, then generate the SubVIs

5. The SubVI generation should complete successfully:
   - A success message will be displayed
   - "View Report" checkbox should be unchecked
   - "Open the generated library" should be checked by default

6. Click Finish to complete the wizard

7. Remove VIs from the .lvlib library:
   - The .lvlib will open automatically
   - Select all VI files in the library
   - Press the Delete key
   - Click OK on the prompt: "The selected item(s) will be permanently deleted from the project or project library"

8. **File > Save All** and exit LabVIEW

9. In the library directory path selected in the wizard, a `VIs` folder will contain the DLL function call SubVIs

10. Create the "all subvis.vi" file:
    - Open LabVIEW
    - **File > New VI**
    - Drag and drop all contents of the `VIs` directory onto the block diagram (all DLL function VIs EXCEPT the .MNU file)
    - **File > Save All** (save as "all subvis.vi")

### 4. Generate Example VI Templates

1. Copy the example generator files to your LabVIEW development directory (same directory containing the `VIs` folder):
   - `parse_c_examples.py`
   - `generate_all_examples.py`
   - `Example_template.vi`

2. Parse the C examples:
   ```
   python3 parse_c_examples.py --recipes-all ../examples
   ```
   This creates `c_example_recipes.json` containing function calls for each C example.

3. Generate the LabVIEW example VIs:
   ```
   python3 generate_all_examples.py
   ```
   This creates example VIs in the `output_vis` folder and `xlsxwriter_examples.zip`.

4. Create the "all examples.vi" file:
   - Open LabVIEW
   - **File > New VI**
   - Drag and drop all VIs from `output_vis` onto the block diagram
   - **File > Save All** (save as "all examples.vi")
   - Exit LabVIEW

   Note: Example VIs will have a broken run arrow until saving and re-opening (SubVI paths are fixed after save and re-open).

### 5. Manual Coding

The generated example VIs contain the SubVI references but require manual wiring:

- Each example needs to be manually coded from the example template
- Duplicate all functionality side by side using the C example code as reference
- Wire the SubVIs together following the logic in the corresponding C example

Where possible, reuse existing example VIs and relink to newly generated library function call SubVIs. When functionality changes significantly (DLL exported function calls, macros, or LV wrappers change), a new example will need to be wired and tested.

---

## Package Contents

### LabView Directory
| File | Description |
|------|-------------|
| `libxlsxwriter_LV.h` | LabVIEW-compatible header file for DLL import |
| `Example_template.vi` | Template VI containing base SubVIs |
| `parse_c_examples.py` | Script to parse C examples into JSON recipes |
| `generate_all_examples.py` | Script to generate example VIs from recipes |

### Source Directory
| File | Description |
|------|-------------|
| `src/labview_wrappers.c` | LabVIEW wrapper functions for libxlsxwriter |

### Sandbox Directory
| File | Description |
|------|-------------|
| `install_dev_environment.wsb` | Windows Sandbox configuration |
| `README.txt` | Sandbox documentation |
| `shared/build.bat` | Library build script |
| `shared/install_all.bat` | Development environment installation script |

## Generated Files

| File | Description |
|------|-------------|
| `c_example_recipes.json` | Generated JSON containing parsed C example function calls |
| `VIs/` | Directory containing generated DLL function call SubVIs |
| `output_vis/` | Directory containing generated example VIs |
| `xlsxwriter_examples.zip` | Zip archive of generated example VIs |

---

## Notes

- New features (sparklines, chart date axis, autofit helpers, etc.) are implemented in the main libxlsxwriter source code - see `Changes.txt`
- The LabVIEW bindings header (`libxlsxwriter_LV.h`) must be updated when new functions are added to libxlsxwriter
- The wrapper functions (`labview_wrappers.c`) provide LabVIEW-friendly interfaces for complex libxlsxwriter functions
