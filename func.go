package excelize

import "github.com/xuri/excelize/v2"

func ColumnNumberToName(num int) string {
	name, err := ColumnNumberToNameE(num)
	if err != nil {
		return ""
	}

	return name
}

func ColumnNumberToNameE(num int) (string, error) {
	return excelize.ColumnNumberToName(num)
}
