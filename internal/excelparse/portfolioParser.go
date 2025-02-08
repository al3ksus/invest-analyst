package excelparse

import (
	"errors"
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Position struct {
	Ticker         string
	TotalPrice     float64
	Sector         string
	InstrumentType string
}

type PortfolioParser struct {
	Positions []*Position
}

func NewPortfolioParser(positions []*Position) *PortfolioParser {
	parser := &PortfolioParser{
		Positions: positions,
	}

	return parser
}

func (parser *PortfolioParser) Parse(outFolder string) error {
	fileInfo, err := os.Stat(outFolder)
	if err != nil {
		return fmt.Errorf("os.Stat: %w", err)
	}

	if !fileInfo.IsDir() {
		return errors.New("OutFolder must be a directory, not a file")
	}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			log.Panicln(err.Error())
		}
	}()

	err = parser.parseByInstrumentType(f)
	if err != nil {
		return fmt.Errorf("parseByInstrumentType: %v", err)
	}

	wr, err := os.Create(outFolder + "\\investments.xlsx")
	if err != nil {
		return fmt.Errorf("os.Create: %w", err)
	}
	defer func() {
		if err := wr.Close(); err != nil {
			log.Panicln(err.Error())
		}
	}()

	err = f.Write(wr)
	if err != nil {
		return fmt.Errorf("excelize.File.Write: %w", err)
	}

	return nil
}

func (parser *PortfolioParser) parseByInstrumentType(f *excelize.File) error {
	sheets := make(map[string]int)
	percentStyle, err := f.NewStyle(&excelize.Style{NumFmt: 10})
	if err != nil {
		return err
	}

	for _, p := range parser.Positions {
		if _, exists := sheets[p.InstrumentType]; !exists {
			_, err := f.NewSheet(p.InstrumentType)
			if err != nil {
				return err
			}
		}

		sheets[p.InstrumentType]++
		err := f.SetCellValue(p.InstrumentType, "A"+strconv.Itoa(sheets[p.InstrumentType]+1), p.Ticker)
		if err != nil {
			return err
		}

		err = f.SetCellValue(p.InstrumentType, "B"+strconv.Itoa(sheets[p.InstrumentType]+1), p.TotalPrice)
		if err != nil {
			return err
		}

		err = f.SetCellFormula(p.InstrumentType, "C"+strconv.Itoa(sheets[p.InstrumentType]+1), fmt.Sprintf("=B%d/B1", sheets[p.InstrumentType]+1))
		if err != nil {
			return err
		}
	}

	for sheetName, count := range sheets {
		err := f.SetCellValue(sheetName, "A1", "TOTAL")
		if err != nil {
			return err
		}

		err = f.SetCellFormula(sheetName, "B1", fmt.Sprintf("=SUM(B2:B%d)", count+1))
		if err != nil {
			return err
		}

		err = f.SetCellStyle(sheetName, "C2", fmt.Sprintf("C%d", count+1), percentStyle)
		if err != nil {
			return err
		}
	}

	return nil
}
