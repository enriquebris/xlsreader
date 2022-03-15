package xlsreader

import (
	"errors"
	"fmt"
	"strings"

	"github.com/takuoki/clmconv"

	"github.com/xuri/excelize/v2"
)

type ExcelizeReader struct {
	file             *excelize.File
	filepath         string
	firstColumnIndex int
	// list of columns in the same order they appear in the file
	columnNameList map[string][]string
	// map of columns to their Excel index
	columnNameMp map[string]map[string]string
}

func NewExcelizeReader(filepath string) (Reader, error) {
	ret := &ExcelizeReader{}
	if err := ret.initialize(filepath); err != nil {
		return nil, err
	}

	return ret, nil
}

func (st *ExcelizeReader) initialize(filepath string) error {
	st.columnNameList = make(map[string][]string)
	st.columnNameMp = make(map[string]map[string]string)
	st.filepath = filepath

	if f, err := excelize.OpenFile(filepath); err != nil {
		return err
	} else {
		st.file = f
	}

	return nil
}

func cleanSheetName(sheet string) string {
	return strings.TrimSpace(sheet)
}

func (st *ExcelizeReader) ReadColumns(sheet string, rowIndex uint) error {
	sheet = cleanSheetName(sheet)
	if sheet == "" {
		return errors.New("sheet name cannot be empty")
	}

	// initialize for this sheet
	st.columnNameList[sheet] = make([]string, 0)
	st.columnNameMp[sheet] = make(map[string]string)

	// no known first column index yet
	st.firstColumnIndex = -1

	//columnIndex := int(firstColumnIndex)
	columnIndex := 0
	for {
		cell := fmt.Sprintf("%v%v", clmconv.Itoa(columnIndex), rowIndex)
		cValue, err := st.file.GetCellValue(sheet, cell)
		if err != nil {
			// remove columns for this sheet
			st.columnNameList[sheet] = nil
			st.columnNameMp[sheet] = nil

			return err
		}
		if strings.TrimSpace(cValue) == "" {
			if st.firstColumnIndex == -1 {
				columnIndex++
				continue
			}
			break
		}

		if st.firstColumnIndex == -1 {
			st.firstColumnIndex = columnIndex
		}

		st.columnNameList[sheet] = append(st.columnNameList[sheet], cValue)
		st.columnNameMp[sheet][cValue] = clmconv.Itoa(columnIndex)

		columnIndex++
	}

	return nil
}

func (st *ExcelizeReader) GetColumns(sheet string) ([]string, error) {
	sheet = cleanSheetName(sheet)

	columnList, ok := st.columnNameList[sheet]
	if !ok {
		return nil, fmt.Errorf("sheet '%v' not found", sheet)
	}

	return columnList, nil
}

func (st *ExcelizeReader) GetValue(sheet string, columnName string, rowNumber uint) (string, error) {
	sheet = cleanSheetName(sheet)
	if _, ok := st.columnNameMp[sheet]; !ok {
		return "", fmt.Errorf("sheet '%v' not found", sheet)
	}

	columnName = strings.TrimSpace(columnName)
	originalColumnName := columnName
	columnName, ok := st.columnNameMp[sheet][columnName]
	if !ok {
		return "", fmt.Errorf("column '%v' not found", originalColumnName)
	}

	return st.file.GetCellValue(sheet, fmt.Sprintf("%v%v", columnName, rowNumber))
}

func (st *ExcelizeReader) Close() error {
	return st.file.Close()
}
