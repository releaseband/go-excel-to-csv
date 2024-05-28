package parser

import "fmt"

type Config struct {
	// Sheet is the name of the sheet to be parsed
	Sheet string
	// Start is the starting cell of the range to be parsed
	Start string
	// End is the ending cell of the range to be parsed
	End string
}

func (c Config) String() string {
	return fmt.Sprintf("Sheet: %s, Start: %s, End: %s", c.Sheet, c.Start, c.End)
}
