package main

import (
	"context"
	"log"
	"time"

	excelparse "github.com/al3ksus/invest-analyst/internal/excelParse"
	"github.com/russianinvestments/invest-api-go-sdk/investgo"
	pb "github.com/russianinvestments/invest-api-go-sdk/proto"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

func main() {

	config, err := investgo.LoadConfig("config\\config.yaml")
	if err != nil {
		log.Fatalf("config loading error %v", err.Error())
	}

	// сдк использует для внутреннего логирования investgo.Logger
	// для примера передадим uber.zap
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
	instrumentsServise := client.NewInstrumentsServiceClient()

	portfolio, err := operationsService.GetPortfolio(config.AccountId, pb.PortfolioRequest_RUB)
	if err != nil {
		logger.Fatalf("get positions error. %v", err.Error())
	}

	parser, err := excelparse.NewPortfolioParser(portfolio.Positions, instrumentsServise)
	if err != nil {
		logger.Errorf("error creating PortfolioParser. %v", err.Error())
	}

	err = parser.Parse("out")
	if err != nil {
		logger.Errorf("parse error. %v", err.Error())
	}
}
