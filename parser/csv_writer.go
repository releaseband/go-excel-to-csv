package parser

import (
	"encoding/csv"
	"fmt"
	"os"
)

const fileFormat = ".csv"

func writeCSV(data [][]string, outputName string) error {
	file, err := os.Create(outputName + fileFormat)
	if err != nil {
		return err
	}

	defer func() {
		err = file.Close()
		if err != nil {
			fmt.Println("error while closing file", err)
		}
	}()

	err = csv.NewWriter(file).WriteAll(data)
	if err != nil {
		return fmt.Errorf("csv.NewWriter(file).WriteAll(): %w", err)
	}

	return nil
}
