package extractor

// Response struct for structured output
type Response struct {
	Status  string `json:"status"`
	Message string `json:"message"`
	Data    any    `json:"data,omitempty"`
}

type CompanyNameList []string
