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
	pdfDocument, err := fitz.New(pdfPath)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer pdfDocument.Close()
	for i := 0; i < pdfDocument.NumPage(); i++ {
		page := pdfDocument.Page(i)
		boxes := page.TextBlocks()
		textBoxes := make([]map[string]interface{}, len(boxes))
		for i, box := range boxes {
			textBoxes[i] = map[string]interface{}{
				"x0":    box[0],
				"y0":    box[1],
				"x1":    box[2],
				"y1":    box[3],
				"text":  strings.Replace(box.Text, "\n", "", -1),
				"index": i,
			}
		}
		groupedTextBoxes := groupMapsByRange(textBoxes)
		for _, group := range groupedTextBoxes {
			col := 0
			for _, element := range group {
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", 'A'+col, row), element["text"])
				col++
			}
			row++
		}
	}
}

func groupMapsByRange(mapList []map[string]interface{}) [][]map[string]interface{} {
	sortedMaps := make([]map[string]interface{}, len(mapList))
	copy(sortedMaps, mapList)
	for i := 0; i < len(sortedMaps); i++ {
		for j := i + 1; j < len(sortedMaps); j++ {
			if sortedMaps[i]["y0"].(float64) > sortedMaps[j]["y0"].(float64) {
				sortedMaps[i], sortedMaps[j] = sortedMaps[j], sortedMaps[i]
			}
		}
	}
	groups := [][]map[string]interface{}{}
	for _, d := range sortedMaps {
		addedToExistingGroup := false
		for _, group := range groups {
			if d["y0"].(float64) >= group[0]["y0"].(float64) && d["y0"].(float64) <= group[len(group)-1]["y1"].(float64) {
				if d["y1"].(float64) > OMIT_FIRST_LINES_COORDINATE {
					group = append(group, d)
					addedToExistingGroup = true
					break
				}
			}
		}
		if !addedToExistingGroup {
			groups = append(groups, []map[string]interface{}{d})
		}
	}
	return groups
}

func main() {
	pdfPath, outputPath := parseArguments()
	if pdfPath == "" || outputPath == "" {
		fmt.Println("Please provide both input PDF path and output Excel path")
		return
	}
	xlsx := initWorkbook(outputPath)
	writeTextBoxesToExcel(pdfPath, xlsx)
	if err := xlsx.SaveAs(outputPath); err != nil {
		fmt.Println("Error:", err)
	}
}
