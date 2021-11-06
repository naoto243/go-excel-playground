package ex

type Obj struct {
	SectionName string
	Fields      []Field
	Children    []Obj
}

type Field struct {
	Name string
	String string
}

type ExportedRows []ExportedRow

func (self Obj) Convert(col , row  int , sectionName string) (i , j int , ee ExportedRows){


	l := []ExportedRow{}

	if sectionName != "" {
		l = append(l , ExportedRow{
			Col: col,
			Row: row,
			Name: sectionName,
		})
		col += 1
		row--
	}

	for _ , u := range self.Fields {
		row++
		e := ExportedRow{
			Col:   col,
			Row:   row,
			Value: u.String,
			Name: u.Name,
		}

		l = append(l , e)
	}

	for i := range self.Children {
		u := self.Children[i]
		_ , _r , ll := u.Convert(col , row + 1, u.SectionName)
		l = append(l , ll...)
		row = _r
	}

	return col , row , l
}

type ExportedRow  struct {
	Col int
	Row  int
	Name string
	Value string

}

