package parser

type Field struct {
	Output
	Config
}

type Output struct {
	// Output directory
	Dir string
	// Output file name
	File string
}
