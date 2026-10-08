package excelizeutil

import (
	"bytes"
	"encoding/xml"
	"strings"

	"github.com/xuri/excelize/v2"
)

var fillPatterns = []string{
	"none",
	"solid",
	"mediumGray",
	"darkGray",
	"lightGray",
	"darkHorizontal",
	"darkVertical",
	"darkDown",
	"darkUp",
	"darkGrid",
	"darkTrellis",
	"lightHorizontal",
	"lightVertical",
	"lightDown",
	"lightUp",
	"lightGrid",
	"lightTrellis",
	"gray125",
	"gray0625",
}

var borderStyles = []string{
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

func GetStyle(excelBuffer bytes.Buffer, sheetName string, corIdx, rowIdx int) excelize.Style {
	f, err := excelize.OpenReader(&excelBuffer)
	if err != nil {
		panic(err)
	}
	defer func() { _ = f.Close() }()

	cell, err := excelize.CoordinatesToCellName(corIdx, rowIdx)
	if err != nil {
		panic(err)
	}
	styleID, err := f.GetCellStyle(sheetName, cell)
	if err != nil {
		panic(err)
	}

	numFmt, customNumFmt := getCellNumberFormat(f, styleID)
	return excelize.Style{
		Border:       getCellBorder(f, styleID),
		Fill:         getCellFill(f, styleID),
		Font:         getCellFont(f, styleID),
		Alignment:    getCellAlignment(f, styleID),
		Protection:   getCellProtection(f, styleID),
		NumFmt:       numFmt,
		CustomNumFmt: customNumFmt,
	}
}

func getCellBorder(f *excelize.File, styleID int) []excelize.Border {
	borderID := f.Styles.CellXfs.Xf[styleID].BorderID
	if borderID == nil || *borderID < 0 || *borderID >= len(f.Styles.Borders.Border) {
		return nil
	}

	definition := f.Styles.Borders.Border[*borderID]
	borders := make([]excelize.Border, 0, 6)
	appendBorder := func(side, style string, color *xlsxColor) {
		if style == "" {
			return
		}
		border := excelize.Border{
			Type:  side,
			Style: borderStyleID(style),
		}
		if colors := getCellFillColor(f, color); len(colors) > 0 {
			border.Color = colors[0]
		}
		borders = append(borders, border)
	}

	appendBorder("top", definition.Top.Style, (*xlsxColor)(definition.Top.Color))
	appendBorder("bottom", definition.Bottom.Style, (*xlsxColor)(definition.Bottom.Color))
	appendBorder("left", definition.Left.Style, (*xlsxColor)(definition.Left.Color))
	appendBorder("right", definition.Right.Style, (*xlsxColor)(definition.Right.Color))
	if definition.DiagonalUp {
		appendBorder("diagonalUp", definition.Diagonal.Style, (*xlsxColor)(definition.Diagonal.Color))
	}
	if definition.DiagonalDown {
		appendBorder("diagonalDown", definition.Diagonal.Style, (*xlsxColor)(definition.Diagonal.Color))
	}
	return borders
}

func borderStyleID(style string) int {
	for i, candidate := range borderStyles {
		if candidate == style {
			return i
		}
	}
	return 0
}

func getCellFill(f *excelize.File, styleID int) excelize.Fill {
	fillID := f.Styles.CellXfs.Xf[styleID].FillID
	if fillID == nil || *fillID < 0 || *fillID >= len(f.Styles.Fills.Fill) {
		return excelize.Fill{}
	}
	patternFill := *f.Styles.Fills.Fill[*fillID].PatternFill

	pattern := 0
	for i, name := range fillPatterns {
		if name == patternFill.PatternType {
			pattern = i
			break
		}
	}

	return excelize.Fill{
		Type:    "pattern",
		Pattern: pattern,
		Color:   getCellFillColor(f, (*xlsxColor)(patternFill.FgColor)),
	}
}

func getCellFont(f *excelize.File, styleID int) *excelize.Font {
	fontID := f.Styles.CellXfs.Xf[styleID].FontID
	if fontID == nil || *fontID <= 0 || *fontID >= len(f.Styles.Fonts.Font) {
		return nil
	}
	font := f.Styles.Fonts.Font[*fontID]
	result := &excelize.Font{}
	if font.B != nil && font.B.Val != nil {
		result.Bold = *font.B.Val
	}
	if font.I != nil && font.I.Val != nil {
		result.Italic = *font.I.Val
	}
	if font.U != nil && font.U.Val != nil {
		result.Underline = *font.U.Val
	}
	if font.Name != nil && font.Name.Val != nil {
		result.Family = *font.Name.Val
	}
	if font.Sz != nil && font.Sz.Val != nil {
		result.Size = *font.Sz.Val
	}
	if font.Strike != nil && font.Strike.Val != nil {
		result.Strike = *font.Strike.Val
	}
	if colors := getCellFillColor(f, (*xlsxColor)(font.Color)); len(colors) > 0 {
		result.Color = colors[0]
	}
	return result
}

func getCellAlignment(f *excelize.File, styleID int) *excelize.Alignment {
	alignment := f.Styles.CellXfs.Xf[styleID].Alignment
	if alignment == nil {
		return nil
	}
	return &excelize.Alignment{
		Horizontal:      alignment.Horizontal,
		Indent:          alignment.Indent,
		JustifyLastLine: alignment.JustifyLastLine,
		ReadingOrder:    alignment.ReadingOrder,
		RelativeIndent:  alignment.RelativeIndent,
		ShrinkToFit:     alignment.ShrinkToFit,
		TextRotation:    alignment.TextRotation,
		Vertical:        alignment.Vertical,
		WrapText:        alignment.WrapText,
	}
}

func getCellProtection(f *excelize.File, styleID int) *excelize.Protection {
	protection := f.Styles.CellXfs.Xf[styleID].Protection
	if protection == nil {
		return nil
	}
	result := &excelize.Protection{Locked: true}
	if protection.Locked != nil {
		result.Locked = *protection.Locked
	}
	if protection.Hidden != nil {
		result.Hidden = *protection.Hidden
	}
	return result
}

func getCellNumberFormat(f *excelize.File, styleID int) (int, *string) {
	numFmtID := f.Styles.CellXfs.Xf[styleID].NumFmtID
	if numFmtID == nil {
		return 0, nil
	}
	if *numFmtID < 164 {
		return *numFmtID, nil
	}
	if f.Styles.NumFmts == nil {
		return 0, nil
	}
	for _, numFmt := range f.Styles.NumFmts.NumFmt {
		if numFmt.NumFmtID == *numFmtID {
			format := numFmt.FormatCode
			return 0, &format
		}
	}
	return 0, nil
}

type xlsxColor struct {
	Auto    bool    `xml:"auto,attr,omitempty"`
	RGB     string  `xml:"rgb,attr,omitempty"`
	Indexed int     `xml:"indexed,attr,omitempty"`
	Theme   *int    `xml:"theme,attr"`
	Tint    float64 `xml:"tint,attr,omitempty"`
}

type xlsxThemeColor struct {
	SysClr *struct {
		LastClr string `xml:"lastClr,attr"`
	} `xml:"sysClr"`
	SrgbClr *struct {
		Val *string `xml:"val,attr"`
	} `xml:"srgbClr"`
}

type xlsxThemeColorScheme struct {
	Dk1      xlsxThemeColor `xml:"dk1"`
	Lt1      xlsxThemeColor `xml:"lt1"`
	Dk2      xlsxThemeColor `xml:"dk2"`
	Lt2      xlsxThemeColor `xml:"lt2"`
	Accent1  xlsxThemeColor `xml:"accent1"`
	Accent2  xlsxThemeColor `xml:"accent2"`
	Accent3  xlsxThemeColor `xml:"accent3"`
	Accent4  xlsxThemeColor `xml:"accent4"`
	Accent5  xlsxThemeColor `xml:"accent5"`
	Accent6  xlsxThemeColor `xml:"accent6"`
	Hlink    xlsxThemeColor `xml:"hlink"`
	FolHlink xlsxThemeColor `xml:"folHlink"`
}

func getCellFillColor(f *excelize.File, color *xlsxColor) []string {
	if color == nil {
		return nil
	}

	rgb := color.RGB
	if color.Theme != nil {
		if f.Theme == nil {
			return nil
		}
		schemeXML, err := xml.Marshal(f.Theme.ThemeElements.ClrScheme)
		if err != nil {
			return nil
		}
		var scheme xlsxThemeColorScheme
		if err := xml.Unmarshal(schemeXML, &scheme); err != nil {
			return nil
		}
		children := [...]xlsxThemeColor{
			scheme.Lt1, scheme.Dk1, scheme.Lt2, scheme.Dk2,
			scheme.Accent1, scheme.Accent2, scheme.Accent3, scheme.Accent4,
			scheme.Accent5, scheme.Accent6, scheme.Hlink, scheme.FolHlink,
		}
		index := *color.Theme
		if index < 0 || index >= len(children) {
			return nil
		}
		child := children[index]
		switch {
		case child.SysClr != nil:
			rgb = child.SysClr.LastClr
		case child.SrgbClr != nil && child.SrgbClr.Val != nil:
			rgb = *child.SrgbClr.Val
		default:
			return nil
		}
	}

	rgb = strings.TrimPrefix(strings.ToUpper(rgb), "#")
	if len(rgb) == 8 {
		rgb = rgb[2:]
	}
	if len(rgb) != 6 {
		return nil
	}
	if color.Theme != nil && color.Tint != 0 {
		rgb = strings.TrimPrefix(excelize.ThemeColor(rgb, color.Tint), "FF")
	}
	return []string{"#" + rgb}
}
