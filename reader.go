package xlsreader

type Reader interface {
	// ReadColumns reads columns (header)
	ReadColumns(sheet string, rowIndex uint) error
	// GetColumns returns columns in the same order they are read
	GetColumns(sheet string) ([]string, error)
	// GetValue returns value of the cell
	GetValue(sheet string, columnName string, rowNumber uint) (string, error)
	// Close closes the reader
	Close() error
}
