package main

import (
	"encoding/json"
	"excelreport/models"
	"fmt"
	"log"
	"net/http"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func getAll() models.ExampleResponse {
	return models.ExampleResponse{
		Prop1: "NinetyNine",
		Prop2: 99,
	}
}

func generateExcelFileExample(name string) *excelize.File {
	fs := excelize.NewFile() // new filestream

	// Create a new sheet.
	index := fs.NewSheet("Sheet2")

	// Set value of a cell.
	fs.SetCellValue("Sheet2", "A2", "Hello world.")
	fs.SetCellValue("Sheet1", "B2", 100)

	// Set active sheet of the workbook.
	fs.SetActiveSheet(index)

	// Save spreadsheet by the given path.
	if err := fs.SaveAs(name); err != nil {
		fmt.Println(err)
	}

	return fs
}

func home(resw http.ResponseWriter, req *http.Request) {
	x := getAll()
	json.NewEncoder(resw).Encode(x)
}

func generateReport(resw http.ResponseWriter, req *http.Request) {
	outputPath := "."
	fileName := "Book1.xlsx"
	filePath := fmt.Sprintf("%s/%s", outputPath, fileName)
	var fs *excelize.File = generateExcelFileExample(filePath)
	buffer, err := fs.WriteToBuffer()

	if err != nil {
		fmt.Fprint(resw, fmt.Sprintf("%s", err))
		return
	}

	resw.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	resw.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filePath))
	http.ServeContent(resw, req, filePath, time.Now(), strings.NewReader(buffer.String()))
}

func defineRoutes() {
	http.HandleFunc("/", home)
	http.HandleFunc("/generate-report", generateReport)
}

func main() {
	defineRoutes()
	log.Fatal(http.ListenAndServe(":8088", nil))
}
