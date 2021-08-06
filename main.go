package main

import (
	"database/sql"
	"errors"
	"fmt"
	"strings"

	_ "github.com/mattn/go-sqlite3"
	"github.com/xuri/excelize/v2"
)

var ErrNoMoreRowsFound = errors.New("no rows to read from")

const (
	STYLE_TEXT   = 1
	STYLE_NUMBER = 3
	STYLE_DATE   = 3
)

// createColumnSQL will return SQLite column definition as string.
func createColumnSQL(name string, colType int) string {
	switch colType {
	case STYLE_NUMBER:
		return fmt.Sprintf(`"%s" REAL`, name)
	default:
		return fmt.Sprintf(`"%s" TEXT`, name)
	}
}

func main() {

	//path := "Projektbuchungen_637637716202213668.xlsx"
	path := "Projektbuchungen_637638315176560370.xlsx"
	fp, err := excelize.OpenFile(path)
	if err != nil {
		panic(err)
	}

	// Get all sheets in file
	sheets := fp.GetSheetList()
	fmt.Printf("sheets: %v\n", sheets)

	// Get all the rows in the Sheet1.
	sheet := sheets[0]

	// load headers
	iter, err := fp.Rows(sheet)
	if err != nil {
		panic(err)
	}

	if !iter.Next() {
		panic(ErrNoMoreRowsFound)
	}

	headers, err := iter.Columns()
	if err != nil {
		panic(err)
	}

	columns := make([]string, len(headers))

	// guess data format from second row
	// note: we will not advance row iterator for this!
	for col, header := range headers {

		// lookup data in row two
		cellAddr, err := excelize.CoordinatesToCellName(col+1, 2)
		if err != nil {
			panic(err)
		}

		// TODO return error if cell does not exist
		value, err := fp.GetCellValue(sheet, cellAddr)
		if err != nil {
			panic(err)
		}

		style, err := fp.GetCellStyle(sheet, cellAddr)
		if err != nil {
			panic(err)
		}

		columns[col] = createColumnSQL(header, style)

		fmt.Printf("header %d: %v (style: %v; example value: %v)\n",
			col, header, style, value)

	}

	tableSql := fmt.Sprintf("CREATE TABLE data (%s);", strings.Join(columns, ",\n"))

	//db, err := sql.Open("sqlite3", "file::memory:?cache=shared")
	db, err := sql.Open("sqlite3", "test.db")
	if err != nil {
		panic(err)
	}

	if _, err := db.Exec(tableSql); err != nil {
		panic(err)
	}

	for iter.Next() {

		cells, err := iter.Columns()
		if err != nil {
			panic(err)
		}

		// we need to escape headers with "xxx" in case they contain
		// spaces or reserved keywords
		columns := make([]string, len(headers))
		for i, hdr := range headers {
			columns[i] = fmt.Sprintf(`"%s"`, hdr)
		}

		placeholders := make([]string, len(headers))
		for i := range placeholders {
			placeholders[i] = "?"
		}

		stmt := fmt.Sprintf(`INSERT INTO data (%s) VALUES (%s);`,
			strings.Join(columns, ", "), strings.Join(placeholders, ", "))

		// we need to convert []string to []interface{} in order
		// to use it in db.Exec()
		values := make([]interface{}, len(headers))
		for i, v := range cells {
			values[i] = v
		}

		if _, err := db.Exec(stmt, values...); err != nil {
			panic(err)
		}

	}

}
