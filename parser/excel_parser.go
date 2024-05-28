package parser

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func readExcel(file *excelize.File, cfg Config) ([][]string, error) {
	startCol, startRow, err := splitCellName(cfg.Start)
	if err != nil {
		return nil, err
	}

	colEnd, rowEnd, err := splitCellName(cfg.End)
	if err != nil {
		return nil, err
	}

	colNumStart, err := convertColNameToNumber(startCol)
	if err != nil {
		return nil, err
	}

	colNumEnd, err := convertColNameToNumber(colEnd)
	if err != nil {
		return nil, err
	}

	resp := make([][]string, 0, rowEnd-startRow+1)

	rows, err := file.GetRows(cfg.Sheet)
	if err != nil {
		return nil, fmt.Errorf("get rows from sheet '%s': %w", cfg.Sheet, err)
	}

	for rowNum := startRow - 1; rowNum < len(rows) && rowNum < rowEnd; rowNum++ {
		// in Excel, indexing starts from 1, but in Go it starts from 0
		row := rows[rowNum]

		rowData := make([]string, 0, colNumEnd-colNumStart+1)

		for colNum := colNumStart; colNum <= colNumEnd; colNum++ {
			if colNum-1 < len(row) {
				rowData = append(rowData, row[colNum-1])
			}
		}

		if len(rowData) > 0 {
			resp = append(resp, rowData)
		}
	}

	return resp, nil
}

func convertColNameToNumber(colName string) (int, error) {
	colNum, err := excelize.ColumnNameToNumber(colName)
	if err != nil {
		return 0, fmt.Errorf("convert column name - '%s' to number: %w",
			colName, err)
	}

	return colNum, nil
}

func splitCellName(cellName string) (string, int, error) {
	colName, rowNum, err := excelize.SplitCellName(cellName)
	if err != nil {
		return "", 0, fmt.Errorf("split cell name - '%s': %w", cellName, err)
	}

	return colName, rowNum, nil
}
