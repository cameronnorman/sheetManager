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
	service       *spreadsheet.Service
	googleSheet   *spreadsheet.Spreadsheet
}

func (s *SheetManager) LoadSheet(sheetIndex uint) ([]map[string]string, error) {
	rows := []map[string]string{}
	sheet, err := s.googleSheet.SheetByIndex(sheetIndex)
	if err != nil {
		return rows, err
	}

	keys := make([]string, 0)
	for _, col := range sheet.Rows[0] {
		keys = append(keys, col.Value)
	}

	for k, row := range sheet.Rows {
		if k != 0 {
			rowDetail := make(map[string]string)
			for i, cell := range row {
				rowDetail[keys[i]] = cell.Value
			}
			rows = append(rows, rowDetail)
		}
	}

	return rows, nil
}

func (s *SheetManager) SheetByName(name string) (uint, error) {
	sheet, err := s.googleSheet.SheetByTitle(name)
	if err != nil {
		return 0, err
	}

	return sheet.Properties.Index, nil
}

func (s *SheetManager) CreateSheet(name string) (uint, error) {
	err := s.service.AddSheet(s.googleSheet, spreadsheet.SheetProperties{Title: name})
	if err != nil {
		return 0, err
	}

	return s.SheetByName(name)
}

func (s *SheetManager) UpdateValue(sheetIndex uint, column int, row int, newValue string) error {
	sheet, err := s.googleSheet.SheetByIndex(sheetIndex)

	if err != nil {
		return err
	}

	sheet.Update(row, column, newValue)

	return nil
}

func (s *SheetManager) DeleteRow(sheetIndex uint, rowIndex int) error {
	sheet, err := s.googleSheet.SheetByIndex(sheetIndex)
	if err != nil {
		return err
	}

	sheet.DeleteRows(rowIndex, (rowIndex + 1))
	return nil
}

func (s *SheetManager) Sync(sheetIndex uint) error {
	sheet, err := s.googleSheet.SheetByIndex(sheetIndex)
	if err != nil {
		return err
	}

	err = sheet.Synchronize()
	if err != nil {
		return err
	}

	return nil
}

func New(secretsFilePath string, spreadsheetID string) *SheetManager {
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

	googleSheet, err := service.FetchSpreadsheet(spreadsheetID)
	if err != nil {
		log.Fatal(err)
	}

	s := &SheetManager{
		secrets:       secrets,
		config:        conf,
		spreadsheetID: spreadsheetID,
		service:       service,
		googleSheet:   &googleSheet,
	}

	return s
}
