package exceldb

import (
	"database/sql"
	"errors"
	"fmt"
	"strings"
	"time"

	_ "github.com/mattn/go-sqlite3"
	"github.com/xuri/excelize/v2"
)

// ErrNoMoreRowsFound will be raised if no (or no more) rows can be read from an Excel file.
var ErrNoMoreRowsFound = errors.New("no rows to read from")

const (
	FormatText   = 1
	FormatDate   = 2
	FormatNumber = 3
)

const InMemoryDb = "file::memory:?cache=shared"

// Converter is a function to convert a cell value from string to a
// Go native data type.
type Converter func(val string) (interface{}, error)

// Column describes all values in a column in an Excel sheet.
type Column struct {
	// Name of the column (usually contained in the first row)
	Name string
	// Column format (e.g. FormatNumber)
	Format int
	// Converter function to convert cell content to native value
	Func Converter
}

func DateColum(name string, dateFormat string) Column {
	return Column{
		Name:   name,
		Format: FormatDate,
		Func: func(val string) (interface{}, error) {
			return time.Parse(dateFormat, val)
		},
	}
}

// GuessColumnFormats will return a list of columns based on the format of the
// cells in the second row of the Excel sheet. Note that we assume that the first
// row contains headers!
func GuessColumnFormats(fp *excelize.File, sheet string) ([]Column, error) {

	iter, err := fp.Rows(sheet)
	if err != nil {
		return nil, err
	}

	if !iter.Next() {
		return nil, ErrNoMoreRowsFound
	}

	headers, err := iter.Columns()
	if err != nil {
		return nil, err
	}

	columns := make([]Column, len(headers))

	// guess data format from second row
	for col, header := range headers {

		//// lookup data in row two
		//cellAddr, err := excelize.CoordinatesToCellName(col+1, 2)
		//if err != nil {
		//	return nil, err
		//}
		//
		//// TODO return error if cell does not exist
		//value, err := fp.GetCellValue(sheet, cellAddr)
		//if err != nil {
		//	return nil, err
		//}
		//
		//// Note to future self: GetCellStyle() refers to styles
		//// in fp.Styles.CellXfs.Xf[style] which contains ...
		//style, err := fp.GetCellStyle(sheet, cellAddr)
		//if err != nil {
		//	return nil, err
		//}
		//
		//fmt.Printf("header %d: %v (style: %v; example value: %v)\n",
		//	col, header, style, value)
		//fmt.Printf("CellXfs.Xf[%d].NumFmtId=%d\n", style, *fp.Styles.CellXfs.Xf[style].NumFmtID)
		//for i, format := range fp.Styles.NumFmts.NumFmt {
		//	fmt.Printf("format %d: id=%d code=%s\n", i, format.NumFmtID, format.FormatCode)
		//}

		columns[col] = Column{
			Name:   header,
			Format: FormatText, // FIXME use correct format
			Func:   nil,
		}

	}

	return columns, nil
}

// LoadFromExcel will load all rows from the first sheet in the Excel workbook
// into a newly created SQLite database. You may specify how data in specific
// columns is interpreted by supplying Column definitons. If you don't all will
// be text.
func LoadFromExcel(path string, dsn string, options ...Column) (*sql.DB, error) {

	fp, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}

	// Get all the rows for first sheet in workbook
	sheets := fp.GetSheetList()
	sheet := sheets[0]
	iter, err := fp.Rows(sheet)
	if err != nil {
		return nil, err
	}

	if !iter.Next() {
		return nil, ErrNoMoreRowsFound
	}

	headers, err := iter.Columns()
	if err != nil {
		return nil, err
	}

	// maps converter functions to columns, so we can convert
	// string to native types later
	converters := make(map[int]Converter)

	columns, err := GuessColumnFormats(fp, sheet)
	if err != nil {
		return nil, err
	}
	for i, col := range columns {
		for _, option := range options {
			if col.Name == option.Name {
				columns[i] = option
				if option.Func != nil {
					converters[i] = option.Func
				}
			}
		}
	}

	columnSql := make([]string, len(columns))
	for i, col := range columns {
		switch col.Format {
		case FormatNumber:
			columnSql[i] = fmt.Sprintf(`"%s" REAL`, col.Name)
		default:
			columnSql[i] = fmt.Sprintf(`"%s" TEXT`, col.Name)
		}
	}

	db, err := sql.Open("sqlite3", dsn)
	if err != nil {
		return nil, err
	}

	tableSql := fmt.Sprintf("CREATE TABLE data (%s);", strings.Join(columnSql, ","))
	if _, err := db.Exec(tableSql); err != nil {
		return nil, err
	}

	// import all rows one after the other

	for iter.Next() {

		cells, err := iter.Columns()
		if err != nil {
			return nil, err
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
			if converter, ok := converters[i]; ok {
				converted, err := converter(v)
				if err != nil {
					return nil, err
				}
				values[i] = converted
			} else {
				values[i] = v
			}
		}

		if _, err := db.Exec(stmt, values...); err != nil {
			return nil, err
		}

	}

	return db, nil
}
