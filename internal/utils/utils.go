package utils

import (
	"regexp"
	"strings"
)

// RemoveExtraSpaces removes multiple spaces between words and trims the string
func RemoveExtraSpaces(value string) string {
	// Convert to lowercase if needed
	value = strings.ToLower(value)

	// Use regular expression to replace multiple spaces with a single space
	spacesRegex := regexp.MustCompile(`\s+`)
	cleanedValue := spacesRegex.ReplaceAllString(value, " ")

	// Trim leading and trailing spaces
	return strings.TrimSpace(cleanedValue)
}
