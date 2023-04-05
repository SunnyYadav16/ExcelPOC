package controllers

import (
	"fmt"
	beego "github.com/beego/beego/v2/server/web"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"time"
)

type ExcelController struct {
	beego.Controller
}

// URLMapping ...
func (excelCont *ExcelController) URLMapping() {
	excelCont.Mapping("ReadExcel", excelCont.ReadExcel)
}

func (excelCont *ExcelController) ReadExcel() {
	fmt.Println(1)

	file, header, err := excelCont.GetFile("file")
	fmt.Println("-----------------> ", file, header, err)
	if err != nil {
		excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
		excelCont.ServeJSON()
		return
	}
	defer file.Close()

	filename := strconv.FormatInt(time.Now().UnixNano(), 10) + filepath.Ext(header.Filename)
	filePath := filepath.Join("uploads", filename)

	// Create a new file on the server to save the uploaded file
	out, err := os.Create(filePath)
	if err != nil {
		excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
		excelCont.ServeJSON()
		return
	}
	defer out.Close()

	// Copy the file to the server
	_, err = io.Copy(out, file)
	if err != nil {
		excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
		excelCont.ServeJSON()
		return
	}

	// Open the uploaded file for reading
	file, err = os.Open(filePath)
	if err != nil {
		excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
		excelCont.ServeJSON()
		return
	}
	defer file.Close()

	// Create a new XLSX file handler
	xlsxFile, err := excelize.OpenReader(file)
	if err != nil {
		excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
		excelCont.ServeJSON()
		return
	}

	// Iterate through each sheet in the XLSX file
	for _, sheet := range xlsxFile.GetSheetMap() {
		// Iterate through each row in the sheet
		rows, err := xlsxFile.Rows(sheet)
		if err != nil {
			excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
			excelCont.ServeJSON()
			return
		}
		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				excelCont.Data["json"] = map[string]interface{}{"error": err.Error()}
				excelCont.ServeJSON()
				return
			}
			// Iterate through each cell in the row
			for _, cellValue := range row {
				// Do something with the cell value
				fmt.Println(cellValue)
			}
		}
	}

	excelCont.Data["json"] = map[string]interface{}{"success": true}
	excelCont.ServeJSON()

}
