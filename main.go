package main

import (
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile(os.Getenv("EXCEL_FILE_PATH"))

	if err != nil {
		log.Fatal(err)
	}

	apiServer := NewAPIServer(":3002", f)
	apiServer.Run()
}
