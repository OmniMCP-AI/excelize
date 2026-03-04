package excelize

import (
	"fmt"
	"strings"
	"testing"
)

func TestDebugChartTitleXML(t *testing.T) {
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
	xml := string(chartBytes)
	// Find the title section
	titleStart := strings.Index(xml, "<c:title>")
	titleEnd := strings.Index(xml, "</c:title>")
	if titleStart >= 0 && titleEnd >= 0 {
		fmt.Println("Title XML:")
		fmt.Println(xml[titleStart : titleEnd+len("</c:title>")])
	} else {
		fmt.Println("No <c:title> found")
		// try without prefix
		titleStart = strings.Index(xml, "<title>")
		titleEnd = strings.Index(xml, "</title>")
		if titleStart >= 0 && titleEnd >= 0 {
			fmt.Println("Title XML (no prefix):")
			fmt.Println(xml[titleStart : titleEnd+len("</title>")])
		}
	}
}
