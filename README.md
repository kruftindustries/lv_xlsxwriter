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
