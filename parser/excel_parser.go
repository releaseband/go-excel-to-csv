package parser

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func readExcel(file *excelize.File, cfg Config) ([][]string, error) {
	startCol, startRow, err := excelize.SplitCellName(cfg.Start)
	if err != nil {
		return nil, err
	}

	colEnd, rowEnd, err := excelize.SplitCellName(cfg.End)
	if err != nil {
		return nil, err
	}

	colNumStart, err := excelize.ColumnNameToNumber(startCol)
	if err != nil {
		return nil, err
	}

	colNumEnd, err := excelize.ColumnNameToNumber(colEnd)
	if err != nil {
		return nil, err
	}

	resp := make([][]string, 0, rowEnd-startRow+1)

	rows, err := file.GetRows(cfg.Sheet)
	if err != nil {
		return nil, fmt.Errorf("GetRows(): %w", err)
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
