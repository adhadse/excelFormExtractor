package extractor

import (
	"encoding/json"
	"fmt"
	"reflect"
	"strings"

	"github.com/Kavida/excelExtractor/internal/utils"
	"github.com/xuri/excelize/v2"
)

// SearchCriteria defines what to look for and where
type SearchCriteria struct {
	SearchTerms           []string                  // Multiple possible terms to search for
	CellRanges            []CellRange               // Multiple cell ranges to search in
	DualColumnCheckBoxClf bool                      // check side by side column
	DualColumnClfCritera  DualClassificationCritera // Add this to map checkbox text to values
	BoolCheckBox          bool
	BoolClfCriteria       BoolClassificationCriteria
	Offset                int // Default offset of value for simple fields
}

type ClassificationCriteria struct {
	Label string
	bool
	SearchTerms []string // search terms for extra check if form control has that name or not
	Offset      int
}

type BoolClassificationCriteria struct {
	Offset      int
	SearchTerms []string // search terms for extra check if form control has that name or not
}

type DualClassificationCritera struct {
	TYPE_1 ClassificationCriteria
	TYPE_2 ClassificationCriteria
}

// CellRange represents an Excel cell range
type CellRange struct {
	StartCell string // e.g., "B12"
	EndCell   string // e.g., "D12"
}

type BuyerDetails struct {
	SheetName                       string `json:"sheet_name"`
	PartNumber                      string `json:"part_number"`
	PartDescription                 string `json:"part_description"`
	LeonardoClassificationOfItem    string `json:"leonardo_classification_of_item"`
	ControlListClassificationNumber string `json:"control_list_classification_number"`
	RFQ                             string `json:"rfq"`
	BuildToPrint                    bool   `json:"build_to_print"`
	ManufacturedToSpecification     bool   `json:"manufactured_to_specification"`
	OriginalEquipmentManufacturer   bool   `json:"original_equipment_manufacturer"`
	Modified                        bool   `json:"modified"`
}

type SECCFExtraction struct {
	BuyerDetails *BuyerDetails `json:"buyer_details"`
	// add more extraction if possible
}

type ExcelExtractor struct {
	file       *excelize.File
	Extraction *SECCFExtraction
}

// //////////////////////////
// // Specific extractor ////
// //////////////////////////

// Add a method to handle value extraction based on criteria type
type ValueExtractor interface {
	Extract(e *ExcelExtractor, criteria SearchCriteria, cellRange CellRange) (interface{}, error)
}

// Implement different extractors for different types of fields
type SimpleValueExtractor struct{}
type BoolCheckBoxExtractor struct{}
type DualColumnClfExtractor struct{}

func (s *SimpleValueExtractor) Extract(e *ExcelExtractor, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	adjacentRange := getAdjacentRange(cellRange, criteria.Offset)
	return e.GetCellValue(adjacentRange, e.Extraction.BuyerDetails.SheetName)
}

func (c *BoolCheckBoxExtractor) Extract(e *ExcelExtractor, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	cell := getAdjacentRange(cellRange, criteria.BoolClfCriteria.Offset).StartCell
	return e.isCheckBoxChecked(e.Extraction.BuyerDetails.SheetName, cell, criteria.BoolClfCriteria.SearchTerms)
}

func (d *DualColumnClfExtractor) Extract(e *ExcelExtractor, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	cellType1 := getAdjacentRange(cellRange, criteria.DualColumnClfCritera.TYPE_1.Offset).StartCell
	cellType2 := getAdjacentRange(cellRange, criteria.DualColumnClfCritera.TYPE_2.Offset).StartCell

	isType1, err := e.isCheckBoxChecked(e.Extraction.BuyerDetails.SheetName, cellType1, criteria.DualColumnClfCritera.TYPE_1.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.DualColumnClfCritera.TYPE_1.Label, err)
		return "", err
	}

	isType2, err := e.isCheckBoxChecked(e.Extraction.BuyerDetails.SheetName, cellType2, criteria.DualColumnClfCritera.TYPE_2.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.DualColumnClfCritera.TYPE_2.Label, err)
		return "", err
	}

	if isType1 && !isType2 {
		return criteria.DualColumnClfCritera.TYPE_1.Label, nil
	} else if !isType1 && isType2 {
		return criteria.DualColumnClfCritera.TYPE_2.Label, nil
	}
	return "", nil
}

// ///////////////////////////////
// // Specific extractor ENDS ////
// ///////////////////////////////

// Helper function to set values using reflection
func setValue(field reflect.Value, value interface{}) {
	switch field.Kind() {
	case reflect.String:
		field.SetString(value.(string))
	case reflect.Bool:
		field.SetBool(value.(bool))
		// Add more types as needed
	}
}

func MakeSECCFExtractor(filePath string) (*ExcelExtractor, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open Excel file: %w", err)
	}

	return &ExcelExtractor{
		file:       f,
		Extraction: &SECCFExtraction{},
	}, nil
}

func (e *ExcelExtractor) searchSheetName(searchWord string) (bool, error) {
	sheetList := e.file.GetSheetList()

	wordFound := false

	for index, sheetName := range sheetList {
		// Convert both strings to lowercase for case-insensitive comparison
		if strings.Contains(strings.ToLower(sheetName), strings.ToLower(searchWord)) {
			fmt.Printf("Found '%s' in sheet name: '%s' at position %d\n", searchWord, sheetName, index+1)
			e.Extraction.BuyerDetails = &BuyerDetails{
				SheetName: sheetName,
			}
			wordFound = true
		}
	}

	if !wordFound {
		return wordFound, SheetNotFoundError{searchWord: searchWord}
	}
	return wordFound, nil
}

func (e *ExcelExtractor) GetCellValue(cellRange CellRange, sheetName string) (string, error) {
	// First try getting the value from the start cell
	value, err := e.file.GetCellValue(sheetName, cellRange.StartCell)
	if err != nil {
		return "", fmt.Errorf("failed to get cell value: %w", err)
	}

	// If we got a value, return it
	if value != "" {
		return strings.TrimSpace(value), nil
	}

	// If no value in start cell, check if it's part of a merged range
	mergedCells, err := e.file.GetMergeCells(sheetName)
	if err != nil {
		return "", fmt.Errorf("failed to get merged cells: %w", err)
	}

	// Check each merged range
	for _, mergedCell := range mergedCells {
		if e.isCellInRange(cellRange.StartCell, &mergedCell) {
			// fmt.Printf("Cellvalue: %s\n", strings.TrimSpace(mergedCell.GetCellValue()))
			return strings.TrimSpace(mergedCell.GetCellValue()), nil
		}
	}

	return "", nil
}

// isCellInRange checks if a cell is within a merged cell range
func (e *ExcelExtractor) isCellInRange(cell string, mergedCell *excelize.MergeCell) bool {
	return strings.HasPrefix(mergedCell.GetStartAxis(), cell) ||
		strings.HasPrefix(cell, mergedCell.GetStartAxis())
}

// getAdjacentRange returns the adjacent cell range
func getAdjacentRange(cellRange CellRange, offset int) CellRange {
	// Parse the column and row from StartCell
	startCol := strings.Split(cellRange.StartCell, "")[0]
	startRow := strings.Split(cellRange.StartCell, "")[1:]

	// Parse the column and row from EndCell
	endCol := strings.Split(cellRange.EndCell, "")[0]
	endRow := strings.Split(cellRange.EndCell, "")[1:]

	// Calculate new columns (offset to the right)
	newStartCol := string(rune(startCol[0] + byte(offset)))
	newEndCol := string(rune(endCol[0] + byte(offset)))

	return CellRange{
		StartCell: newStartCol + strings.Join(startRow, ""),
		EndCell:   newEndCol + strings.Join(endRow, ""),
	}
}

func (e *ExcelExtractor) extractBuyerDetails(criteria map[string]SearchCriteria) {
	buyerDetailsValue := reflect.ValueOf(e.Extraction.BuyerDetails).Elem()

	for fieldName, searchCriteria := range criteria {
		for _, cellRange := range searchCriteria.CellRanges {
			// Get 'KEY' cell from the potential label cell
			value, err := e.GetCellValue(cellRange, e.Extraction.BuyerDetails.SheetName)
			if err != nil {
				fmt.Printf("Error trying to retrive cell value: %v\n", err)
				return
			}

			// Check if the value matches any of our search terms
			for _, searchTerm := range searchCriteria.SearchTerms {
				if strings.Contains(utils.RemoveExtraSpaces(strings.ToLower(value)), strings.ToLower(searchTerm)) {
					var extractor ValueExtractor

					// Select appropriate extractor based on criteria type
					if searchCriteria.DualColumnCheckBoxClf {
						extractor = &DualColumnClfExtractor{}
					} else if searchCriteria.BoolCheckBox {
						extractor = &BoolCheckBoxExtractor{}
					} else {
						extractor = &SimpleValueExtractor{}
					}

					// Extract the value
					extractedValue, err := extractor.Extract(e, searchCriteria, cellRange)
					if err != nil {
						fmt.Println("error extracting value: %w", err)
						return
					}

					// Set the field using reflection
					field := buyerDetailsValue.FieldByName(fieldName)
					if field.IsValid() && field.CanSet() {
						setValue(field, extractedValue)
					} else {
						fmt.Printf("field: %s is isValid: %v and canSet: %v\n", fieldName, field.IsValid(), field.CanSet())
					}

					break
				}
			}
		}
	}
}

func (e *ExcelExtractor) ReadFormControls() {
	formControls, err := e.file.GetFormControls(e.Extraction.BuyerDetails.SheetName) // sheet name
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, control := range formControls {
		fmt.Printf("Control Cell %s, Control checked: %v, control.Paragraph %v, control.CurrentVal %v, cellLink: %s, offsetX: %v, offsetY: %v, Control cell text: %s\n", control.Cell, control.Checked, control.Paragraph, control.CurrentVal, control.CellLink, control.Format.OffsetX, control.Format.OffsetY, control.Text)
		// if control.Type == excelize.FormControlCheckBox {
		// }
	}
}

func (e *ExcelExtractor) isCheckBoxChecked(sheetName string, cell string, classificationTexts []string) (bool, error) {
	formControls, err := e.file.GetFormControls(sheetName)
	if err != nil {
		return false, fmt.Errorf("failed to get form controls: %w", err)
	}

	for _, control := range formControls {
		if control.Cell == cell {
			if control.Type == excelize.FormControlCheckBox {
				for _, text := range classificationTexts {
					for _, paraText := range control.Paragraph {
						if strings.EqualFold(paraText.Text, text) {
							fmt.Printf("Control Cell %s, Control checked: %v Control cell text: %s, cell %s\n", control.Cell, control.Checked, control.Text, cell)
							return control.Checked, nil
						}
					}
				}
			}
		}
	}
	return false, nil
}

func (e *ExcelExtractor) Extract() {
	_, err := e.searchSheetName("buyer details")

	criteria := map[string]SearchCriteria{
		"PartNumber": {
			SearchTerms: []string{"part number", "part-nr", "part_number"},
			CellRanges: []CellRange{
				{StartCell: "B12", EndCell: "D12"},
			},
			Offset: 3,
		},
		"PartDescription": {
			SearchTerms: []string{"description", "desc", "part description"},
			CellRanges: []CellRange{
				{StartCell: "B13", EndCell: "D13"},
			},
			Offset: 3,
		},
		"ControlListClassificationNumber": {
			SearchTerms: []string{"control list classification number"},
			CellRanges: []CellRange{
				{StartCell: "B18", EndCell: "D18"},
			},
			Offset: 3,
		},
		"RFQ": {
			SearchTerms: []string{"RQF", "quote reference"},
			CellRanges: []CellRange{
				{StartCell: "B19", EndCell: "D19"},
			},
			Offset: 3,
		},
		"BuildToPrint": {
			SearchTerms: []string{"Build To Print"},
			CellRanges: []CellRange{
				{StartCell: "B21", EndCell: "F21"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"ManufacturedToSpecification": {
			SearchTerms: []string{"Manufactured to specification", "(MTS)"},
			CellRanges: []CellRange{
				{StartCell: "B22", EndCell: "F22"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"OriginalEquipmentManufacturer": {
			SearchTerms: []string{"Original Equipment Manufacturer"},
			CellRanges: []CellRange{
				{StartCell: "B23", EndCell: "F23"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"Modified": {
			SearchTerms: []string{"Modified"},
			CellRanges: []CellRange{
				{StartCell: "B25", EndCell: "F25"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"LeonardoClassificationOfItem": {
			SearchTerms: []string{"Leonardo Classification of item"},
			CellRanges: []CellRange{
				{StartCell: "B15", EndCell: "D15"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCritera: DualClassificationCritera{
				TYPE_1: ClassificationCriteria{
					Label:       "DUAL",
					SearchTerms: []string{"Dual", "DU"},
					Offset:      3,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "MILITARY",
					SearchTerms: []string{"Military", "MIL"},
					Offset:      4,
				},
			},
		},
	}

	if err != nil {
		fmt.Println("%w", err)
	} else {
		e.extractBuyerDetails(criteria)
	}

	jsonBytes, err := json.Marshal(e.Extraction)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	fmt.Printf("Extraction with nil BuyerDetails: %s \n", string(jsonBytes))
}

func (e *ExcelExtractor) Close() error {
	return e.file.Close()
}

///////////////////////
//// working on it ////
///////////////////////
