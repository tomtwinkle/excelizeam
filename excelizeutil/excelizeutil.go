package excelizeutil

import (
	"bytes"
	"strings"

	"github.com/xuri/excelize/v2"
)

func GetStyle(excelBuffer bytes.Buffer, sheetName string, corIdx, rowIdx int) excelize.Style {
	f, err := excelize.OpenReader(&excelBuffer)
	if err != nil {
		panic(err)
	}
	cell, err := excelize.CoordinatesToCellName(corIdx, rowIdx)
	if err != nil {
		panic(err)
	}
	styleID, err := f.GetCellStyle(sheetName, cell)
	if err != nil {
		panic(err)
	}
	return excelize.Style{
		Border:        getCellBorder(f, styleID),
		Fill:          getCellFill(f, styleID),
		Font:          nil,
		Alignment:     getCellAlignment(f, styleID),
		Protection:    nil,
		NumFmt:        0,
		DecimalPlaces: 0,
		CustomNumFmt:  nil,
		Lang:          "",
		NegRed:        false,
	}
}

type xlsxLine struct {
	Style string     `xml:"style,attr,omitempty"`
	Color *xlsxColor `xml:"color,omitempty"`
}

func getCellBorder(f *excelize.File, styleID int) []excelize.Border {
	borderID := f.Styles.CellXfs.Xf[styleID].BorderID
	if borderID == nil {
		return nil
	}

	getStyleID := func(style string) int {
		var styles = []string{
			"none",
			"thin",
			"medium",
			"dashed",
			"dotted",
			"thick",
			"double",
			"hair",
			"mediumDashed",
			"dashDot",
			"mediumDashDot",
			"dashDotDot",
			"mediumDashDotDot",
			"slantDashDot",
		}
		for i, v := range styles {
			if v == style {
				return i
			}
		}
		return 0
	}
	getBorder := func(style string, color *xlsxColor) excelize.Border {
		b := excelize.Border{
			Type:  "top",
			Style: getStyleID(style),
		}
		if color != nil {
			b.Color = getCellFillColor(f, color)[0]
		}
		return b
	}

	borders := make([]excelize.Border, 0, 4)

	if topStyle := f.Styles.Borders.Border[*borderID].Top; topStyle.Style != "" {
		borders = append(borders, getBorder(topStyle.Style, (*xlsxColor)(topStyle.Color)))
	}

	return borders
}

func getCellFill(f *excelize.File, styleID int) excelize.Fill {
	fillID := f.Styles.CellXfs.Xf[styleID].FillID
	if fillID == nil {
		return excelize.Fill{}
	}
	patternFill := *f.Styles.Fills.Fill[*fillID].PatternFill

	fgColor := f.Styles.Fills.Fill[*fillID].PatternFill.FgColor

	return excelize.Fill{
		Type:    patternFill.PatternType,
		Pattern: 0,
		Color:   getCellFillColor(f, (*xlsxColor)(fgColor)),
	}
}

func getCellAlignment(f *excelize.File, styleID int) *excelize.Alignment {
	alignment := f.Styles.CellXfs.Xf[styleID].Alignment
	if alignment == nil {
		return nil
	}
	return &excelize.Alignment{
		Horizontal:      "",
		Indent:          0,
		JustifyLastLine: false,
		ReadingOrder:    0,
		RelativeIndent:  0,
		ShrinkToFit:     false,
		TextRotation:    0,
		Vertical:        "",
		WrapText:        alignment.WrapText,
	}
}

type xlsxColor struct {
	Auto    bool    `xml:"auto,attr,omitempty"`
	RGB     string  `xml:"rgb,attr,omitempty"`
	Indexed int     `xml:"indexed,attr,omitempty"`
	Theme   *int    `xml:"theme,attr"`
	Tint    float64 `xml:"tint,attr,omitempty"`
}

func getCellFillColor(f *excelize.File, color *xlsxColor) []string {
	if color != nil {
		if color.Theme != nil {
			children := f.Theme.ThemeElements.ClrScheme.Children
			if *color.Theme < 4 {
				dklt := map[int]string{
					0: children[1].SysClr.LastClr,
					1: children[0].SysClr.LastClr,
					2: *children[3].SrgbClr.Val,
					3: *children[2].SrgbClr.Val,
				}
				return []string{
					strings.TrimPrefix(
						excelize.ThemeColor(dklt[*color.Theme], color.Tint), "FF"),
				}
			}
			srgbClr := *children[*color.Theme].SrgbClr.Val
			return []string{strings.TrimPrefix(excelize.ThemeColor(srgbClr, color.Tint), "FF")}
		}
		return []string{strings.TrimPrefix(color.RGB, "FF")}
	}
	return nil
}
