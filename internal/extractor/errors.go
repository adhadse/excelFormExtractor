package extractor

import "fmt"

type DivideByZeroError struct {
	dividend int
}

type SheetNotFoundError struct {
	searchWord string
}

func (e SheetNotFoundError) Error() string {
	return fmt.Sprintf("Sheet name not found for searchWord: %s", e.searchWord)
}
