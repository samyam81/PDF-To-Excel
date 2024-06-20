# PDF-To-Excel Converter

This repository contains a Go application (`APP.go`) that converts text data from a PDF file into an Excel spreadsheet. It utilizes the `go-fitz` library for PDF parsing and `excelize` for Excel file manipulation.

### Requirements
- Go 1.11+
- Libraries:
  - `github.com/gen2brain/go-fitz`
  - `github.com/360EntSecGroup-Skylar/excelize`

### Installation
1. Make sure you have Go installed and set up properly.
2. Clone this repository:
   ```bash
   git clone https://github.com/samyam81/PDF-To-Excel
   ```
3. Install dependencies:
   ```bash
   go mod tidy
   ```

### Usage
The application expects two command-line arguments:
- `pdf_path`: Path to the input PDF file.
- `output_path`: Path to the output Excel file.

Example usage:
```bash
go run APP.go -pdf_path input.pdf -output_path output.xlsx
```

### How It Works
1. **Argument Parsing**: Command-line arguments (`pdf_path` and `output_path`) are parsed using the `flag` package.
2. **Excel Initialization**: An Excel workbook is initialized using `excelize.NewFile()`.
3. **PDF Processing**:
   - The input PDF file is opened and read using `go-fitz`.
   - Text blocks are extracted from each page of the PDF using `page.TextBlocks()`.
   - Each text block is cleaned of newline characters and written into successive rows in the Excel sheet (`Sheet1`).
4. **Excel Writing**:
   - Text blocks are written into corresponding cells in the Excel sheet, starting from column A and incrementing the row for each text block.
5. **Saving**: The resulting Excel file is saved to the specified output path using `xlsx.SaveAs(outputPath)`.

### Notes
- Text blocks in the PDF are directly processed without additional grouping by vertical position (`groupMapsByRange` function is removed).
- Ensure the input PDF is structured such that text extraction results in meaningful rows and columns in the Excel output.

### Author
This project was developed by [Samyam](https://github.com/samyam81).

---

### Explanation of Changes:
- **How It Works**: Updated to reflect the direct extraction of text blocks from each page of the PDF using `page.TextBlocks()` method.
- **Excel Writing**: Clarified that text blocks are written into Excel starting from column A and incrementing the row for each block.
- **Notes**: Removed the section about vertical position grouping (`groupMapsByRange` function) as it was not utilized in the revised code.
  
