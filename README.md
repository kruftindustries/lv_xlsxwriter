# lv_xlsxwriter

A LabVIEW wrapper for [libxlsxwriter](https://libxlsxwriter.github.io/), a C library for creating Excel XLSX files.

## Features

- 318+ VIs wrapping libxlsxwriter functions
- Create Excel files with formatting, formulas, charts, and more
- Support for worksheets, chartsheets, images, and data validation
- Example VIs demonstrating common use cases

## Requirements

- LabVIEW (32-bit or 64-bit)
- xlsxwriter.dll (must match your LabVIEW bitness)

## Installation

1. Copy the `VIs` and `Controls` folders to your LabVIEW project or user.lib
2. Place `xlsxwriter.dll` in a location accessible to LabVIEW (e.g., alongside your VI or in the system PATH)

## Quick Start

```
1. Open Examples/Example_hello_world.vi
2. Run the VI to create a sample XLSX file
```

## Building the DLL

See `Development Resources/README.txt` for instructions on building xlsxwriter.dll using Windows Sandbox with Visual Studio.

### Prerequisites

You must provide your own LabVIEW installation ISO and specfile:

**Creating the LabVIEW Specfile:**

1. Mount your LabVIEW installer ISO
2. Open a command prompt and run:
   ```
   D:\setup.exe /generatespecfile "C:\path\to\specfile\directory\"
   ```
   (Replace `D:\` with your mounted ISO drive letter)
3. The installer GUI will open - select the required packages for your installation
4. Enter your user details and serial number when prompted
5. Complete the wizard - the installer will display the location of the created specfile
6. Copy the generated specfile folder to `Development Resources/shared/specfile_lv/`

**Adding the LabVIEW ISO:**

1. Obtain your LabVIEW installation ISO from NI (requires valid license)
2. Copy the ISO to `Development Resources/shared/LabVIEW.iso`

**Running the Build:**

1. Run `Development Resources/shared/download_vs_layout.bat` to download Visual Studio offline installer
2. Double-click `Development Resources/install_dev_environment.wsb` to launch Windows Sandbox
3. Wait for Visual Studio and LabVIEW installation to complete (20-40 minutes)
4. Run `C:\Users\WDAGUtilityAccount\Desktop\Setup\build.bat`
5. Output DLLs will be in:
   - 32-bit: `Setup\libxlsxwriter-1.2.3\build32\Release\xlsxwriter.dll`
   - 64-bit: `Setup\libxlsxwriter-1.2.3\build64\Release\xlsxwriter.dll`

## Examples

The `Examples` folder contains VIs demonstrating:
- Basic workbook creation
- Charts (bar, column, scatter, area, combined)
- Array formulas
- Autofilters
- Background images
- Chartsheets

## Project Structure

```
lv_xlsxwriter/
├── VIs/                    # Wrapper VIs for libxlsxwriter functions
├── Controls/               # Custom LabVIEW controls (enums, clusters)
├── Examples/               # Example VIs
├── Development Resources/  # DLL build environment setup
├── src/                    # Source files
└── libxlsxwriter_LV.h      # C header for LabVIEW Call Library nodes
```

## License

BSD 2-Clause License - see [LICENSE](LICENSE) for details.

## Acknowledgments

- [libxlsxwriter](https://libxlsxwriter.github.io/) by John McNamara
