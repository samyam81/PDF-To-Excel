package main

import (
	"flag"
	"fmt"
	"strings"

	"github.com/gen2brain/go-fitz"
	"github.com/360EntSecGroup-Skylar/excelize"
)

const OMIT_FIRST_LINES_COORDINATE = 85

func parseArguments() (string, string) {
	var pdfPath string
	var outputPath string
	flag.StringVar(&pdfPath, "pdf_path", "", "Path to the input PDF file")
	flag.StringVar(&outputPath, "output_path", "", "Path to the output Excel file")
	flag.Parse()
	return pdfPath, outputPath
}

func initWorkbook(outputPath string) *excelize.File {
	xlsx := excelize.NewFile()
	return xlsx
}

func writeTextBoxesToExcel(pdfPath string, xlsx *excelize.File) {
	row := 1
	pdfDocument, err := fitz.Open(pdfPath)
	if err != nil {
		fmt.Println("Error opening PDF:", err)
		return
	}
	defer pdfDocument.Close()

	for i := 0; i < pdfDocument.NumPage(); i++ {
		page := pdfDocument.Page(i)
		textBlocks, err := page.TextBlocks()
		if err != nil {
			fmt.Println("Error getting text blocks:", err)
			continue
		}

		for _, block := range textBlocks {
			// Remove newlines from the text
			text := strings.Replace(block.Text, "\n", "", -1)
			// Set cell value in Excel sheet
			err := xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", 'A', row), text)
			if err != nil {
				fmt.Println("Error setting cell value:", err)
				continue
			}
			row++
		}
	}
}

func main() {
	pdfPath, outputPath := parseArguments()
	if pdfPath == "" || outputPath == "" {
		fmt.Println("Please provide both input PDF path and output Excel path")
		return
	}

	xlsx := initWorkbook(outputPath)
	writeTextBoxesToExcel(pdfPath, xlsx)

	// Save the Excel file
	if err := xlsx.SaveAs(outputPath); err != nil {
		fmt.Println("Error saving Excel file:", err)
		return
	}

	fmt.Println("Excel file successfully created:", outputPath)
}
