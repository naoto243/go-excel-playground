package go_excel_playground

import (
	"fmt"
	"github.com/naoto243/go-excel-playground/ex"
	"github.com/sanity-io/litter"
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"os"
	"testing"
)

func Test03(t *testing.T) {


	s := ex.Obj{
		Fields:   []ex.Field{
			{
				Name:   "UserID",
				String:    `10`,
			},
			{
				Name:   "UserName",
				String:    `aaa`,
			},
		},
		Children: []ex.Obj{
			{
				SectionName: `Chapters`,
				Children: []ex.Obj{
					{
						SectionName: "Chapter_1",
						Fields:      []ex.Field{
							{
								Name:   "Title",
								String: "@@@@",
							},
							{
								Name:   "TitleEn",
								String: "----------",
							},
						},
						Children:    nil,
					},
					{
						SectionName: "Chapter_2",
						Fields:      []ex.Field{
							{
								Name:   "Title",
								String: "----------",
							},
							{
								Name:   "TitleEn",
								String: "----------",
							},
						},
						Children:    nil,
					},
				},
			},
		},
	}


	const SheetName = `merge_test`
	const ExportPath = `.local/Book3.xlsx`

	_ =  os.Mkdir(`.local` , 0777)

	f := excelize.NewFile()
	index := f.NewSheet(SheetName)
	f.DeleteSheet(`Sheet1`)
	f.SetActiveSheet(index)


	_ , _ , all := s.Convert(1 , 1 , ``)

	litter.Dump(all)

	for _ , a := range  all {

		litter.Dump(a)

		alpha := stringValueOf(a.Col)
		alpha2 := stringValueOf(a.Col+1)
		row := a.Row

		colName := fmt.Sprintf(`%s%d` , alpha , row)
		f.SetCellValue(SheetName  ,  colName , a.Name)

		colName2 := fmt.Sprintf(`%s%d` , alpha2 , row)
		f.SetCellValue(SheetName  ,  colName2 , a.Value)
	}


	err := f.SaveAs(ExportPath)
	assert.NoError(t, err)
	if err != nil {
		fmt.Println(err)
	}




}

func stringValueOf(i int) string {
   var foo = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
   return string(foo[i-1])
}