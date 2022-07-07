package main

import (
	"fmt"

	"github.com/redtoad/exceldb"
)

func main() {

	path := "Book1.xlsx"
	db, err := exceldb.LoadFromExcel(path, exceldb.InMemoryDb,
		exceldb.DateColum("Date", "01/02/06"))
	if err != nil {
		panic(err)
	}
	defer db.Close()

	rows, err := db.Query(`SELECT 
	strftime("%m-%Y", "Date") as 'month-year',
		Employee,
		SUM("Hours worked")
	FROM data
	WHERE "Status" != "non billable"
	GROUP BY strftime("%m-%Y", "Date"), Employee;`)
	if err != nil {
		panic(err)
	}
	defer rows.Close()

	for rows.Next() {
		var monthYear string
		var resource string
		var sum float64
		if err := rows.Scan(&monthYear, &resource, &sum); err != nil {
			panic(err)
		}
		fmt.Printf("%s -> %s -> %f\n", monthYear, resource, sum)
	}

}
