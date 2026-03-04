package excelize

import (
	"bytes"
	"fmt"
	"testing"
)

func TestDebugChartTitle(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)
	f.SetCellValue(sheet, "A2", "Small")
	f.SetCellValue(sheet, "B2", 2)
	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
	}
	f.AddChart(sheet, "E1", &Chart{
		Type: Col, Series: series, Title: []RichTextRun{{Text: "TestTitle"}},
	})

	chartBytes := f.readXML("xl/charts/chart1.xml")
	var cs xlsxChartSpace
	f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(chartBytes))).Decode(&cs)

	title := cs.Chart.Title
	if title == nil {
		fmt.Println("Title is nil!")
		return
	}
	fmt.Printf("Title.Tx.Rich: %v (nil? %v)\n", title.Tx.Rich, title.Tx.Rich == nil)
	if title.Tx.Rich != nil {
		fmt.Printf("P count: %d\n", len(title.Tx.Rich.P))
		for i, p := range title.Tx.Rich.P {
			fmt.Printf("  P[%d]: PPr=%v, R=%v, EndParaRPr=%v\n", i, p.PPr != nil, p.R, p.EndParaRPr)
			if p.R != nil {
				fmt.Printf("    R.T=%q, R.RPr.Sz=%f\n", p.R.T, p.R.RPr.Sz)
			}
		}
	}
}
