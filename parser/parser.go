package parser

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path"
)

type Parser struct {
	// Fields with configuration mapping to excel data
	// for parsing to csv
	fields []Field

	// File modes for creating directories
	fileModes os.FileMode
}

func New(
	fileModes os.FileMode,
	fields []Field,
) *Parser {
	return &Parser{
		fields:    fields,
		fileModes: fileModes,
	}
}

// Parse reads Excel file
// read/parsing config maps
// read/parsing data from Excel to csv
func (p *Parser) Parse(
	excelFilePath string,
	outDirPrefix string,
) error {
	f, err := excelize.OpenFile(excelFilePath)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}

	var outDir string
	var outPath string

	for _, field := range p.fields {
		cfg, err := parseCoordination(f, field.Config)
		if err != nil {
			return err
		}

		data, err := readExcel(f, *cfg)
		if err != nil {
			return fmt.Errorf("parse excel %w", err)
		}

		outDir, err = p.makeDir(outDirPrefix, field.Output.Dir)
		if err != nil {
			return err
		}

		outPath = path.Join(outDir, field.Output.File)
		err = writeCSV(data, outPath)
		if err != nil {
			return fmt.Errorf("write csv %s %w", outPath, err)
		}
	}

	return nil
}

func (p *Parser) makeDir(prefix, outDir string) (string, error) {
	resp := path.Join(prefix, outDir)

	err := os.MkdirAll(resp, p.fileModes)
	if err != nil {
		return "", fmt.Errorf("mkdir: %s %w", resp, err)
	}

	return resp, nil
}

// parseCoordination read Excel file
// parsing config maps
// create Config struct for parsing Excel data
func parseCoordination(f *excelize.File, cfg Config) (*Config, error) {
	data, err := readExcel(f, cfg)
	if err != nil {
		return nil, fmt.Errorf("parse excel %w", err)
	}

	cfgPath, err := parseConfigMap(data)
	if err != nil {
		return nil, fmt.Errorf("parse to config %w", err)
	}

	return &cfgPath, nil
}

// parseConfigMap parse config map
// create Config struct for parsing Excel data
// min data len is 1
// example of config map:
// []string{{ "sheet1","A1","B1"}}
func parseConfigMap(data [][]string) (Config, error) {
	const minDataLen = 1

	if len(data) != minDataLen || len(data[0]) < 2 || len(data[0]) > 3 {
		return Config{}, fmt.Errorf("invalid data len: %d", len(data))
	}

	sheet := getSheet(data)
	start := getStart(data)
	end := start

	if len(data[0]) == 3 {
		end = getEnd(data)
	}

	resp := Config{
		Sheet: sheet,
		Start: start,
		End:   end,
	}

	return resp, nil
}

func getSheet(data [][]string) string {
	return data[0][0]
}

func getStart(data [][]string) string {
	return data[0][1]
}

func getEnd(data [][]string) string {
	return data[0][2]
}
