package excelparse

import (
	"errors"
	"fmt"
	"log"
	"os"

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

var INSTRUMENT_TYPE_RUSSIAN_NAMES = map[string]string{
	"share":    "Акции",
	"bond":     "Облигации",
	"gold":     "Золото",
	"etf":      "Фонды",
	"currency": "Валюта",
	"option":   "Опционы",
	"futures":  "Фьючерсы",
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

	sheets, err := parser.parseByInstrumentType(f)
	if err != nil {
		return fmt.Errorf("parseByInstrumentType: %v", err)
	}

	if len(sheets) != 0 {
		err = parser.parseTotalSheet(f, sheets)
		if err != nil {
			return fmt.Errorf("parseTotalSheet: %w", err)
		}
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

func (parser *PortfolioParser) parseByInstrumentType(f *excelize.File) ([]string, error) {
	sheetList := []string{}
	sheets := make(map[string]int)
	sectorRows := make(map[string]int)
	sectorPrices := make(map[string]float64)
	percentStyle, err := f.NewStyle(&excelize.Style{NumFmt: 10})
	if err != nil {
		return nil, err
	}

	for _, p := range parser.Positions {
		if _, exists := sheets[p.InstrumentType]; !exists {
			_, err := f.NewSheet(INSTRUMENT_TYPE_RUSSIAN_NAMES[p.InstrumentType])
			if err != nil {
				return nil, err
			}
			sheetList = append(sheetList, INSTRUMENT_TYPE_RUSSIAN_NAMES[p.InstrumentType])
		}

		isNew := false
		sectorSheet := fmt.Sprintf("%s%s", p.Sector, p.InstrumentType)
		if _, exists := sectorRows[sectorSheet]; !exists && p.Sector != "" {
			sectorRows[p.InstrumentType]++
			sectorRows[sectorSheet] = sectorRows[p.InstrumentType]
			isNew = true
		}

		sectorPrices[sectorSheet] += p.TotalPrice
		sheets[p.InstrumentType]++
		err = setPositionCell(f, p, sheets[p.InstrumentType]+1)
		if err != nil {
			return nil, err
		}

		if p.Sector != "" {
			err = setSectorCell(f, p, sectorRows[sectorSheet]+1, sectorPrices[sectorSheet], isNew)
			if err != nil {
				return nil, err
			}
		}
	}

	for sheetName, count := range sheets {
		err = setTotalCell(f, INSTRUMENT_TYPE_RUSSIAN_NAMES[sheetName], count+1)
		if err != nil {
			return nil, err
		}

		err = f.SetCellStyle(INSTRUMENT_TYPE_RUSSIAN_NAMES[sheetName], "C2", fmt.Sprintf("C%d", count+1), percentStyle)
		if err != nil {
			return nil, err
		}

		err = f.SetCellStyle(INSTRUMENT_TYPE_RUSSIAN_NAMES[sheetName], "G2", fmt.Sprintf("G%d", sectorRows[sheetName]+1), percentStyle)
		if err != nil {
			return nil, err
		}
	}

	return sheetList, nil
}

func setPositionCell(f *excelize.File, pos *Position, posRow int) error {
	err := f.SetCellValue(INSTRUMENT_TYPE_RUSSIAN_NAMES[pos.InstrumentType], fmt.Sprintf("A%d", posRow), pos.Ticker)
	if err != nil {
		return err
	}

	err = f.SetCellValue(INSTRUMENT_TYPE_RUSSIAN_NAMES[pos.InstrumentType], fmt.Sprintf("B%d", posRow), pos.TotalPrice)
	if err != nil {
		return err
	}

	err = f.SetCellFormula(INSTRUMENT_TYPE_RUSSIAN_NAMES[pos.InstrumentType], fmt.Sprintf("C%d", posRow), fmt.Sprintf("=B%d/B1", posRow))
	if err != nil {
		return err
	}

	return nil
}

func setSectorCell(f *excelize.File, pos *Position, sectorRow int, sectorPrice float64, isNew bool) error {
	if isNew {
		err := f.SetCellValue(INSTRUMENT_TYPE_RUSSIAN_NAMES[pos.InstrumentType], fmt.Sprintf("E%d", sectorRow), pos.Sector)
		if err != nil {
			return err
		}

		err = f.SetCellFormula(INSTRUMENT_TYPE_RUSSIAN_NAMES[pos.InstrumentType], fmt.Sprintf("G%d", sectorRow), fmt.Sprintf("=F%d/B1", sectorRow))
		if err != nil {
			return err
		}
	}

	err := f.SetCellValue(INSTRUMENT_TYPE_RUSSIAN_NAMES[pos.InstrumentType], fmt.Sprintf("F%d", sectorRow), sectorPrice)
	if err != nil {
		return err
	}

	return nil
}

func setTotalCell(f *excelize.File, sheetName string, countPos int) error {
	err := f.SetCellValue(sheetName, "A1", "Всего")
	if err != nil {
		return err
	}

	err = f.SetCellFormula(sheetName, "B1", fmt.Sprintf("=SUM(B2:B%d)", countPos))
	if err != nil {
		return err
	}

	return nil
}

func (parser *PortfolioParser) parseTotalSheet(f *excelize.File, sheets []string) error {
	totalSheetName := "Портфель"
	percentStyle, err := f.NewStyle(&excelize.Style{NumFmt: 10})
	if err != nil {
		return err
	}

	_, err = f.NewSheet(totalSheetName)
	if err != nil {
		return err
	}

	err = f.SetCellValue(totalSheetName, "A1", "Всего")
	if err != nil {
		return err
	}

	row := 2
	for _, s := range sheets {
		err = f.SetCellValue(totalSheetName, fmt.Sprintf("A%d", row), s)
		if err != nil {
			return err
		}

		err = f.SetCellFormula(totalSheetName, fmt.Sprintf("B%d", row), fmt.Sprintf("=%s!B1", s))
		if err != nil {
			return err
		}

		err = f.SetCellFormula(totalSheetName, fmt.Sprintf("C%d", row), fmt.Sprintf("=B%d/B1", row))
		if err != nil {
			return err
		}
		row++
	}

	err = setTotalCell(f, totalSheetName, len(sheets)+1)
	if err != nil {
		return err
	}

	err = f.SetCellStyle(totalSheetName, "C2", fmt.Sprintf("C%d", len(sheets)+1), percentStyle)
	if err != nil {
		return err
	}

	return nil
}
