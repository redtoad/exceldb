package exceldb_test

import (
	"archive/zip"
	"testing"

	"github.com/redtoad/exceldb"
	"github.com/stretchr/testify/assert"
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
