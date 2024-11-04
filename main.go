package main

import (
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("./data_surveys.xlsx")

	if err != nil {
		log.Fatal(err)
	}

	apiServer := NewAPIServer(":3002", f)
	apiServer.Run()
}
