package excelparse

import (
	"errors"
	"fmt"
	"log"
	"os"

	"github.com/russianinvestments/invest-api-go-sdk/investgo"
	investapi "github.com/russianinvestments/invest-api-go-sdk/proto"
	"github.com/xuri/excelize/v2"
)

type PortfolioParser struct {
	Positions          []*investapi.PortfolioPosition
	InstrumentsService *investgo.InstrumentsServiceClient
}

func NewPortfolioParser(positions []*investapi.PortfolioPosition, instrumentsService *investgo.InstrumentsServiceClient) (*PortfolioParser, error) {
	if instrumentsService == nil {
		return nil, errors.New("InstrumentsService can not be nil")
	}

	parser := &PortfolioParser{
		Positions:          positions,
		InstrumentsService: instrumentsService,
	}

	return parser, nil
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

	_, err = f.NewSheet("Акции")
	if err != nil {
		return fmt.Errorf("excelize.File.NewSheet: %w", err)
	}

	parser.parseByInstrumentType(f, "Акции", "share")

	_, err = f.NewSheet("Облигации")
	if err != nil {
		return fmt.Errorf("excelize.File.NewSheet: %w", err)
	}

	parser.parseByInstrumentType(f, "Облигации", "bond")

	_, err = f.NewSheet("Фонды")
	if err != nil {
		return fmt.Errorf("excelize.File.NewSheet: %w", err)
	}

	parser.parseByInstrumentType(f, "Фонды", "etf")

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

func (parser *PortfolioParser) parseByInstrumentType(f *excelize.File, sheetName string, instrumentType string) error {
	instruments, total := parser.getPositionsByInstrumentType(instrumentType)
	f.SetCellValue(sheetName, "A1", "TOTAL")
	f.SetCellValue(sheetName, "B1", total)

	row := "2"[0]
	for _, i := range instruments {
		shareResp, err := parser.InstrumentsService.InstrumentByFigi(i.Figi)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheetName, "A"+string(row), shareResp.Instrument.Ticker)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheetName, "B"+string(row), i.CurrentPrice.ToFloat()*float64(i.Quantity.Units))
		if err != nil {
			return err
		}

		row++
	}

	return nil
}

func (parser *PortfolioParser) getPositionsByInstrumentType(instrumentType string) ([]*investapi.PortfolioPosition, float64) {
	var shares []*investapi.PortfolioPosition
	var total float64

	for _, p := range parser.Positions {
		if p.InstrumentType == instrumentType {
			shares = append(shares, p)
			total += p.CurrentPrice.ToFloat() * float64(p.Quantity.Units)
		}
	}

	return shares, total
}
