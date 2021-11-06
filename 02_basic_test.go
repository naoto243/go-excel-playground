package go_excel_playground

import (
	"encoding/json"
	"fmt"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"os"
	"testing"
)

func Test01_Basic_ReadWrite2(t *testing.T) {

	const SheetName = `merge_test`
	const ExportPath = `.local/Book2.xlsx`

	 _ =  os.Mkdir(`.local` , 0777)

	f := excelize.NewFile()
    index := f.NewSheet(SheetName)
    f.DeleteSheet(`Sheet1`)
	f.SetActiveSheet(index)
    f.SetCellValue(SheetName, "B2", `日本語セル`)

	err := f.MergeCell(SheetName, "D3", "E9")
	f.SetCellValue(SheetName, "D3", `merged cell`)
	assert.NoError(t, err)


	s := excelize.Style{
		Fill:          excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{
				`#EEEEEE`,
			},
		},
	}

	sb  , _ := json.Marshal(s)

	style, err := f.NewStyle(string(sb))
	if err != nil {
		fmt.Println(err)
	}

	err = f.SetCellStyle(SheetName , `D3` , `E9`  , style)
	assert.NoError(t, err)

    err = f.SaveAs(ExportPath)
    assert.NoError(t, err)
    if err != nil {
        fmt.Println(err)
    }

}