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
	FormatFloat  = 4
)

const InMemoryDb = "file::memory:" //?cache=shared"

// Converter is a function to convert a cell value from string to a
// Go native data type.
type Converter func(val string) (interface{}, error)

// KeepAsIs is a Converter function which will not alter the received string value.
func KeepAsIs(val string) (interface{}, error) {
	return val, nil
}

// Column describes all values in a column in an Excel sheet.
type Column struct {
	// Name of the column (usually contained in the first row)
	Name string
	// Column format (e.g. FormatNumber)
	Format int
	// Converter function to convert cell content (string) to native value (Go type)
	Func Converter
}

// DateColum returns a Column with a Converter function parsing date strings using
// dateFormat (as used by time.Parse()).
func DateColum(name string, dateFormat string) Column {
	return Column{
		Name:   name,
		Format: FormatDate,
		Func: func(val string) (interface{}, error) {
			return time.Parse(dateFormat, val)
		},
	}
}

// LoadFromExcel will load all rows from the first sheet in the Excel workbook
// into a newly created SQLite database. You may specify how data in specific
// columns is interpreted by supplying Column definitons. For those you don't,
// they will be treated as TEXT.
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

	columns := make([]Column, len(headers))
	for i, headerLabel := range headers {
		columns[i] = Column{
			Name:   headerLabel,
			Format: FormatText,
			Func:   KeepAsIs,
		}
		for _, option := range options {
			if headerLabel == option.Name {
				columns[i] = option
			}
		}
	}

	columnSql := make([]string, len(columns))
	for i, col := range columns {
		switch col.Format {
		case FormatNumber:
		case FormatFloat:
			columnSql[i] = fmt.Sprintf(`"%s" REAL`, col.Name)
			break
		// also applies to FormatDate
		default:
			columnSql[i] = fmt.Sprintf(`"%s" TEXT`, col.Name)
		}
	}

	db, err := sql.Open("sqlite3", dsn)
	if err != nil {
		return nil, err
	}

	tableName := "data"
	dropTableSql := fmt.Sprintf("DROP TABLE IF EXISTS %s;", tableName)
	if _, err := db.Exec(dropTableSql); err != nil {
		return nil, err
	}

	createTableSql := fmt.Sprintf("CREATE TABLE %s (%s);", tableName, strings.Join(columnSql, ","))
	if _, err := db.Exec(createTableSql); err != nil {
		return nil, err
	}

	// import all rows one after the other
	fields := make([]string, len(headers))
	for i, hdr := range headers {
		// We need to escape headers with quotes ("s) in case they contain
		// spaces or reserved keywords.
		fields[i] = fmt.Sprintf(`"%s"`, hdr)
	}

	placeholders := make([]string, len(headers))
	for i := range placeholders {
		placeholders[i] = "?"
	}

	for iter.Next() {

		cells, err := iter.Columns()
		if err != nil {
			return nil, err
		}

		// Skip empty rows
		if len(cells) == 0 {
			continue
		}

		stmt := fmt.Sprintf(`INSERT INTO data (%s) VALUES (%s);`,
			strings.Join(fields, ", "), strings.Join(placeholders, ", "))

		// We need to convert []string to []interface{} in order to use it in db.Exec().
		values := make([]interface{}, len(headers))

		for i, v := range cells {
			converter := columns[i].Func
			if values[i], err = converter(v); err != nil {
				return nil, err
			}
		}

		if _, err := db.Exec(stmt, values...); err != nil {
			return nil, err
		}

	}

	return db, nil
}
