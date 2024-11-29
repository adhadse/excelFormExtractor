package parser

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func create_sheet() {
	fmt.Println("This is a helper function")

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// // Create a new sheet.
	// index, err := f.NewSheet("Sheet2")
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }
	// // Set value of a cell.
	// f.SetCellValue("Sheet2", "A2", "Hello world.")
	// f.SetCellValue("Sheet1", "B2", 100)
	// // Set active sheet of the workbook.
	// f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func read_sheet() {
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Get value from cell by given worksheet name and cell reference.
	cell, err := f.GetCellValue("Sheet1", "B2")
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println(cell)
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}

}

func add_checkbox() {
	// f := excelize.NewFile()
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// if err := f.AddFormControl("Sheet1", excelize.FormControl{
	// 	Cell:    "A1",
	// 	Type:    excelize.FormControlCheckBox,
	// 	Text:    "Option Button 1",
	// 	Checked: true,
	// 	Height:  20,
	// }); err != nil {
	// 	fmt.Println(err)
	// }

	if err := f.AddFormControl("Sheet1", excelize.FormControl{
		Cell:    "A1",
		Type:    excelize.FormControlOptionButton,
		Text:    "Option Button 1",
		Checked: true,
		Height:  20,
	}); err != nil {
		fmt.Println(err)
	}
	if err := f.AddFormControl("Sheet1", excelize.FormControl{
		Cell:   "A2",
		Type:   excelize.FormControlOptionButton,
		Text:   "Option Button 2",
		Height: 20,
	}); err != nil {
		fmt.Println(err)
	}

	if err := f.Save(); err != nil {
		fmt.Println(err)
	}
}

func ReadFormControls() {
	f, err := excelize.OpenFile("Example_2.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	form_controls, err := f.GetFormControls("Section A - LUK Buyer details")
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, form_control := range form_controls {
		fmt.Println(form_control.Cell, form_control.Text, form_control.Checked)
	}
}

// func main() {
// 	// create_sheet()
// 	// fmt.Println("reading the file")
// 	// read_sheet()
// 	// fmt.Println("adding form control")
// 	// add_checkbox()
// 	read_form_controls()
// }
