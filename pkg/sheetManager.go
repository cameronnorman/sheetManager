package sheetManager

import (
	"io/ioutil"
	"log"
	"net/http"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"golang.org/x/oauth2/jwt"
	"gopkg.in/Iwark/spreadsheet.v2"
)

type SheetManager struct {
	config        *jwt.Config
	secrets       []byte
	client        *http.Client
	spreadsheetID string
	sheet         *spreadsheet.Sheet
}

func (s *SheetManager) LoadSheet() []map[string]string {
	keys := make([]string, 0)
	for _, col := range s.sheet.Rows[0] {
		keys = append(keys, col.Value)
	}

	rows := []map[string]string{}
	for k, row := range s.sheet.Rows {
		if k != 0 {
			rowDetail := make(map[string]string)
			for i, cell := range row {
				rowDetail[keys[i]] = cell.Value
			}
			rows = append(rows, rowDetail)
		}
	}

	return rows
}

func (s *SheetManager) UpdateValue(column int, row int, newValue string) error {
	s.sheet.Update(row, column, newValue)

	return nil
}

func (s *SheetManager) Sync() error {
	err := s.sheet.Synchronize()
	if err != nil {
		return err
	}

	return nil
}

func New(secretsFilePath string, spreadsheetID string, sheetIndex uint) *SheetManager {
	secrets, err := ioutil.ReadFile(secretsFilePath)
	if err != nil {
		log.Fatal(err)
	}

	conf, err := google.JWTConfigFromJSON(secrets, spreadsheet.Scope)
	if err != nil {
		log.Fatal(err)
	}

	client := conf.Client(oauth2.NoContext)
	service := spreadsheet.NewServiceWithClient(client)

	fetchSheetService, err := service.FetchSpreadsheet(spreadsheetID)
	if err != nil {
		log.Fatal(err)
	}

	fetchedSheet, err := fetchSheetService.SheetByIndex(sheetIndex)
	if err != nil {
		log.Fatal(err)
	}

	s := &SheetManager{
		secrets:       secrets,
		config:        conf,
		spreadsheetID: spreadsheetID,
		sheet:         fetchedSheet,
	}

	return s
}
