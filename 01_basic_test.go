package go_excel_playground

import (
	"encoding/json"
	"fmt"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"os"
	"testing"
)

func Test01_Basic_ReadWrite(t *testing.T) {

	 _ =  os.Mkdir(`.local` , 0777)

	f := excelize.NewFile()
    index := f.NewSheet("Sheet1")
	f.SetActiveSheet(index)
    f.SetCellValue("Sheet1", "B2", `日本語セル`)

	s := excelize.Style{
		Font: &excelize.Font{
			Bold:      true,
			Family:    "MS Gothic",
		},
		Fill:          excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{
				`#EE0000`,
			},
		},
	}

	sb  , _ := json.Marshal(s)

/*
	ss := `
{
    "border" : [
        {
            "type": "left",
            "color": "000000",
            "style": 1
        },
        {
            "type": "top",
            "color": "000000",
            "style": 1
        },
        {
            "type": "bottom",
            "color": "000000",
            "style": 1
        },
        {
            "type": "right",
            "color": "000000",
            "style": 1
        }
    ],
    "font" : {
        "size": 14,
        "color": "#0000FF",
        "family": "メイリオ"
    },
    "fill" : {
        "type" : "pattern",
        "color" : ["#FFFF00"],
        "pattern" : 1
    }
}
`

 */

	style, err := f.NewStyle(string(sb))
	if err != nil {
		fmt.Println(err)
	}

	err = f.SetCellStyle(`Sheet1` , `B2` , `B2`  , style)

	//exp := "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"
	//style, err = f.NewStyle(&excelize.Style{CustomNumFmt: &exp})
	//err = f.SetCellStyle("Sheet1", "B2", "B2", style)

	assert.NoError(t, err)
    err = f.SaveAs(".local/Book1.xlsx")
    assert.NoError(t, err)
    if err != nil {
        fmt.Println(err)
    }




	f, err = excelize.OpenFile(".local/Book1.xlsx")
	assert.NoError(t, err)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get value from cell by given worksheet name and axis.
	cell, err := f.GetCellValue("Sheet1", "B2")

	assert.Equal(t, `日本語セル` , cell)

	_  , _ = f.GetCellStyle(`Sheet1` , `B2`)

}