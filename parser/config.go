package parser

type Config struct {
	// Sheet is the name of the sheet to be parsed
	Sheet string
	// Start is the starting cell of the range to be parsed
	Start string
	// End is the ending cell of the range to be parsed
	End string
}
