package exceldb_test

import (
	"archive/zip"
	"database/sql"
	"strconv"
	"testing"

	"github.com/redtoad/exceldb"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func Test_LoadFromExcel(t *testing.T) {

	tests := []struct {
		name  string
		input string
		err   error
	}{
		{
			"fails for invalid file",
			"tests/invalid.xlsx",
			zip.ErrFormat,
		},
		{
			"example file can be loaded",
			"example/Book1.xlsx",
			nil,
		},
		{
			"fails for empty file",
			"tests/empty.xlsx",
			exceldb.ErrNoMoreRowsFound,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			_, err := exceldb.LoadFromExcel(tc.input, exceldb.InMemoryDb)
			assert.Equal(t, tc.err, err)
		})
	}
}

func LoadTestDB() (*sql.DB, error) {

	input := "example/Book1.xlsx"
	output := exceldb.InMemoryDb

	FloatParser := func(val string) (interface{}, error) {
		num, err := strconv.ParseFloat(val, 32)
		if err != nil {
			return nil, err
		}
		return num, nil
	}

	return exceldb.LoadFromExcel(
		input, output,
		exceldb.DateColum("Date", "01/02/06"),
		exceldb.Column{"Hours worked", exceldb.FormatFloat, FloatParser},
	)

}

func Test_LoadFromExcel_CountResults(t *testing.T) {

	db, err := LoadTestDB()
	require.NoError(t, err)
	defer db.Close()

	tt := []struct {
		name  string
		query string
		count int
	}{
		{
			"All rows are imported",
			`SELECT COUNT(*) FROM data;`,
			343,
		},
		{
			"Count Kirk's entries works",
			`SELECT COUNT(*) FROM data WHERE "Employee" = "James T. Kirk";`,
			55,
		},
		{
			"Count captains works",
			`SELECT COUNT(DISTINCT "Employee") FROM data;`,
			4,
		},
	}

	for _, tc := range tt {
		t.Run(tc.name, func(t *testing.T) {
			row, err := db.Query(tc.query)
			require.NoError(t, err)
			defer row.Close()

			for row.Next() {
				var count int
				row.Scan(&count)
				assert.Equal(t, tc.count, count)
			}
		})
	}

}
