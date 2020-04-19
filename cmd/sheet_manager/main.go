package main

import (
	"sheetManager"
)

func main() {
	s := sheetManager.New("secrets.json", "1qwrNpla2imC_5YqDtr47PmXAi9I5fF9mHIrBs54FRb0", 2)
	s.UpdateValue(2, 2, "TRUE")
	s.Sync()
}
