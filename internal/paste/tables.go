package paste

import "encoding/xml"

type Body struct {
	Table Table `xml:"w:tbl"`
}

type Table struct {
	TblPr   TblPr   `xml:"w:tblPr"`
	TblGrid TblGrid `xml:"w:tblGrid"`
	Tr      []Tr    `xml:"w:tr"`
}

type TblPr struct {
	TblStyle TblStyle `xml:"w:tblStyle"`
	TblW     TblW     `xml:"w:tblW"`
}

type TblStyle struct {
	Val string `xml:"w:val,attr"`
}

type TblW struct {
	W    string `xml:"w:w,attr"`
	Type string `xml:"w:type,attr"`
}

type TblGrid struct {
	GridCol []GridCol `xml:"w:gridCol"`
}

type GridCol struct {
	W string `xml:"w:w,attr"`
}

type Tr struct {
	Tc []Tc `xml:"w:tc"`
}

type Tc struct {
	P P `xml:"w:p"`
}

type P struct {
	R R `xml:"w:r"`
}

type R struct {
	T Text `xml:"w:t"`
}

type Text struct {
	Text string `xml:",chardata"`
}

func GetTables() {
	doc := &Body{
		Table: Table{
			TblPr: TblPr{
				TblStyle: TblStyle{Val: "TableGrid"},
				TblW:     TblW{W: "5000", Type: "pct"},
			},
			TblGrid: TblGrid{
				GridCol: []GridCol{{W: "2500"}, {W: "2500"}},
			},
			Tr: []Tr{
				{
					Tc: []Tc{
						{P: P{R: R{T: Text{Text: "Ячейка 1-1"}}}},
						{P: P{R: R{T: Text{Text: "Ячейка 1-2"}}}},
					},
				},
				{
					Tc: []Tc{
						{P: P{R: R{T: Text{Text: "Ячейка 2-1"}}}},
						{P: P{R: R{T: Text{Text: "Ячейка 2-2"}}}},
					},
				},
			},
		},
	}

	// Кодируем XML
	encoder := xml.NewEncoder(documentXml)
	encoder.Indent("", "  ")
	if err := encoder.Encode(doc); err != nil {
		panic(err)
	}
}
