package main

import (
	"encoding/json"
	"fmt"
	"os"
	"strings"

	"github.com/Kavida/excelExtractor/internal/parser"
)

// Response struct for structured output
type Response struct {
	Status  string `json:"status"`
	Message string `json:"message"`
	Data    any    `json:"data,omitempty"`
}

func printSuccessAndExit(response Response) {
	jsonResponse, _ := json.Marshal(response)
	fmt.Println(string(jsonResponse))
	os.Exit(0)
}

func printErrorAndExit(response Response, code int) {
	jsonResponse, _ := json.Marshal(response)
	fmt.Fprintln(os.Stderr, string(jsonResponse))
	os.Exit(code)
}

func main() {
	parser.ReadFormControls()

	// Check if arguments are provided
	if len(os.Args) < 2 {
		// Return error for missing arguments
		response := Response{
			Status:  "error",
			Message: "No arguments provided",
		}
		printErrorAndExit(response, 1)
	}

	// Get the command from first argument
	command := os.Args[1]

	// Example command handling
	switch command {
	case "hello":
		// Success case
		response := Response{
			Status:  "success",
			Message: "Hello command executed successfully",
			Data:    "Hello, World!",
		}
		printSuccessAndExit(response)

	case "process":
		// Check for required second argument
		if len(os.Args) < 3 {
			response := Response{
				Status:  "error",
				Message: "Process command requires a parameter",
			}
			printErrorAndExit(response, 2)
		}

		// Process the input
		input := os.Args[2]
		if strings.ToLower(input) == "fail" {
			// Simulate a process failure
			response := Response{
				Status:  "error",
				Message: "Process failed as requested",
				Data:    map[string]string{"failed_input": input},
			}
			printErrorAndExit(response, 3)
		}

		// Success case with processed data
		response := Response{
			Status:  "success",
			Message: "Processing completed",
			Data:    map[string]string{"processed": strings.ToUpper(input)},
		}
		printSuccessAndExit(response)

	default:
		// Unknown command error
		response := Response{
			Status:  "error",
			Message: fmt.Sprintf("Unknown command: %s", command),
		}
		printErrorAndExit(response, 4)
	}
}
