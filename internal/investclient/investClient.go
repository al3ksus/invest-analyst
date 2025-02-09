package investclient

import (
	"context"
	"log"
	"time"

	"github.com/al3ksus/invest-analyst/internal/excelparse"
	"github.com/russianinvestments/invest-api-go-sdk/investgo"
	investapi "github.com/russianinvestments/invest-api-go-sdk/proto"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

func Run() {
	config, err := investgo.LoadConfig("config\\config.yaml")
	if err != nil {
		log.Fatalf("config loading error %v", err.Error())
	}

	zapConfig := zap.NewDevelopmentConfig()
	zapConfig.EncoderConfig.EncodeTime = zapcore.TimeEncoderOfLayout(time.DateTime)
	zapConfig.EncoderConfig.TimeKey = "time"
	l, err := zapConfig.Build()
	logger := l.Sugar()
	defer func() {
		err := logger.Sync()
		if err != nil {
			log.Println(err.Error())
		}
	}()
	if err != nil {
		log.Fatalf("logger creating error %v", err)
	}
	// создаем клиента для investAPI, он позволяет создавать нужные сервисы и уже
	// через них вызывать нужные методы
	client, err := investgo.NewClient(context.Background(), config, logger)
	if err != nil {
		logger.Fatalf("client creating error %v", err.Error())
	}
	defer func() {
		logger.Infof("closing client connection")
		err := client.Stop()
		if err != nil {
			logger.Errorf("client shutdown error %v", err.Error())
		}
	}()

	operationsService := client.NewOperationsServiceClient()
	instrumentsService := client.NewInstrumentsServiceClient()

	portfolio, err := operationsService.GetPortfolio(config.AccountId, investapi.PortfolioRequest_RUB)
	if err != nil {
		logger.Fatalf("get portfolio error. %v", err.Error())
	}

	positions, err := getPositions(portfolio, instrumentsService)
	if err != nil {
		logger.Fatalf("get positions error. %v", err.Error())
	}

	parser := excelparse.NewPortfolioParser(positions)

	err = parser.Parse("C:\\Users\\22ale\\OneDrive\\Рабочий стол")
	if err != nil {
		logger.Errorf("parse error, %v", err.Error())
	}
}

func getPositions(portfolio *investgo.PortfolioResponse, instrumentsServise *investgo.InstrumentsServiceClient) ([]*excelparse.Position, error) {
	investGoPos := portfolio.Positions

	if len(investGoPos) == 0 {
		return []*excelparse.Position{}, nil
	}

	positions := make([]*excelparse.Position, len(investGoPos))
	for i, p := range investGoPos {
		ticker, sector, err := getTickerAndSector(p.Figi, p.InstrumentType, instrumentsServise)
		if err != nil {
			return nil, err
		}

		positions[i] = &excelparse.Position{
			Ticker:         ticker,
			TotalPrice:     p.CurrentPrice.ToFloat() * p.Quantity.ToFloat(),
			Sector:         sector,
			InstrumentType: p.InstrumentType,
		}

		if p.InstrumentType == "bond" {
			positions[i].TotalPrice += p.CurrentNkd.ToFloat() * p.Quantity.ToFloat()
		}

		if ticker == "TGLD" {
			positions[i].InstrumentType = "gold"
		}
	}

	return positions, nil
}

func getTickerAndSector(figi, instrumentType string, instrumentsServise *investgo.InstrumentsServiceClient) (string, string, error) {
	var err error
	if instrumentType == "share" {
		resp, err := instrumentsServise.ShareByFigi(figi)
		if err == nil {
			return resp.Instrument.Ticker, resp.Instrument.Sector, nil
		}
	} else if instrumentType == "bond" {
		resp, err := instrumentsServise.BondByFigi(figi)
		if err == nil {
			return resp.Instrument.Ticker, resp.Instrument.Sector, nil
		}
	} else if instrumentType == "currency" {
		resp, err := instrumentsServise.CurrencyByFigi(figi)
		if err == nil {
			return resp.Instrument.Ticker, "", nil
		}
	} else if instrumentType == "futures" {
		resp, err := instrumentsServise.FutureByFigi(figi)
		if err == nil {
			return resp.Instrument.Ticker, resp.Instrument.Sector, nil
		}
	} else if instrumentType == "etf" {
		resp, err := instrumentsServise.EtfByFigi(figi)
		if err == nil {
			return resp.Instrument.Ticker, resp.Instrument.Sector, nil
		}
	} else {
		return "", "", nil
	}

	return "", "", err
}
