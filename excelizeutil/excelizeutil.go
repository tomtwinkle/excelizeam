package excelizeutil

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
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
	themeColors, numberFormats, indexedColors := readStyleMetadata(excelBuffer)
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

	format := resolveCellFormat(f, styleID)
	numFmt, customNumFmt := getCellNumberFormat(format.NumFmtID, numberFormats)
	return excelize.Style{
		Border:       getCellBorder(f, format.BorderID, themeColors, indexedColors),
		Fill:         getCellFill(f, format.FillID, themeColors, indexedColors),
		Font:         getCellFont(f, format.FontID, themeColors, indexedColors),
		Alignment:    format.Alignment,
		Protection:   format.Protection,
		NumFmt:       numFmt,
		CustomNumFmt: customNumFmt,
	}
}

type resolvedCellFormat struct {
	FontID     int
	FillID     int
	BorderID   int
	NumFmtID   int
	Alignment  *excelize.Alignment
	Protection *excelize.Protection
}

func resolveCellFormat(f *excelize.File, styleID int) resolvedCellFormat {
	cellXF := f.Styles.CellXfs.Xf[styleID]
	resolved := resolvedCellFormat{
		FontID:   integerValue(cellXF.FontID),
		FillID:   integerValue(cellXF.FillID),
		BorderID: integerValue(cellXF.BorderID),
		NumFmtID: integerValue(cellXF.NumFmtID),
	}
	alignment := cellXF.Alignment
	protection := cellXF.Protection
	if cellXF.XfID != nil && f.Styles.CellStyleXfs != nil && *cellXF.XfID >= 0 && *cellXF.XfID < len(f.Styles.CellStyleXfs.Xf) {
		baseXF := f.Styles.CellStyleXfs.Xf[*cellXF.XfID]
		resolved.FontID = inheritedStyleID(cellXF.FontID, baseXF.FontID, cellXF.ApplyFont)
		resolved.FillID = inheritedStyleID(cellXF.FillID, baseXF.FillID, cellXF.ApplyFill)
		resolved.BorderID = inheritedStyleID(cellXF.BorderID, baseXF.BorderID, cellXF.ApplyBorder)
		resolved.NumFmtID = inheritedStyleID(cellXF.NumFmtID, baseXF.NumFmtID, cellXF.ApplyNumberFormat)

		if cellXF.ApplyAlignment == nil || !*cellXF.ApplyAlignment {
			if baseXF.Alignment != nil {
				alignment = baseXF.Alignment
			} else if cellXF.ApplyAlignment != nil {
				alignment = nil
			}
		}
		if cellXF.ApplyProtection == nil || !*cellXF.ApplyProtection {
			if baseXF.Protection != nil {
				protection = baseXF.Protection
			} else if cellXF.ApplyProtection != nil {
				protection = nil
			}
		}
	}
	if alignment != nil {
		resolved.Alignment = &excelize.Alignment{
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
	if protection != nil {
		resolved.Protection = &excelize.Protection{Locked: true}
		if protection.Locked != nil {
			resolved.Protection.Locked = *protection.Locked
		}
		if protection.Hidden != nil {
			resolved.Protection.Hidden = *protection.Hidden
		}
	}
	return resolved
}

func integerValue(value *int) int {
	if value == nil {
		return 0
	}
	return *value
}

func inheritedStyleID(child, parent *int, apply *bool) int {
	if apply != nil && *apply {
		return integerValue(child)
	}
	if parent != nil {
		return *parent
	}
	if apply != nil && !*apply {
		return 0
	}
	return integerValue(child)
}

func getCellBorder(f *excelize.File, borderID int, themeColors []string, indexedColors map[int]string) []excelize.Border {
	if f.Styles.Borders == nil || borderID < 0 || borderID >= len(f.Styles.Borders.Border) {
		return nil
	}

	definition := f.Styles.Borders.Border[borderID]
	borders := make([]excelize.Border, 0, 6)
	appendBorder := func(side, style string, color *xlsxColor) {
		if style == "" {
			return
		}
		border := excelize.Border{
			Type:  side,
			Style: borderStyleID(style),
		}
		if colors := getCellFillColor(color, themeColors, indexedColors); len(colors) > 0 {
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

func getCellFill(f *excelize.File, fillID int, themeColors []string, indexedColors map[int]string) excelize.Fill {
	if f.Styles.Fills == nil || fillID < 0 || fillID >= len(f.Styles.Fills.Fill) {
		return excelize.Fill{}
	}
	patternFill := *f.Styles.Fills.Fill[fillID].PatternFill

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
		Color:   getCellFillColor((*xlsxColor)(patternFill.FgColor), themeColors, indexedColors),
	}
}

func getCellFont(f *excelize.File, fontID int, themeColors []string, indexedColors map[int]string) *excelize.Font {
	if f.Styles.Fonts == nil || fontID <= 0 || fontID >= len(f.Styles.Fonts.Font) {
		return nil
	}
	font := f.Styles.Fonts.Font[fontID]
	result := &excelize.Font{}
	if font.B != nil && font.B.Val != nil {
		result.Bold = *font.B.Val
	}
	if font.I != nil && font.I.Val != nil {
		result.Italic = *font.I.Val
	}
	if font.U != nil {
		result.Underline = "single"
		if font.U.Val != nil && *font.U.Val != "" {
			result.Underline = *font.U.Val
		}
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
	if colors := getCellFillColor((*xlsxColor)(font.Color), themeColors, indexedColors); len(colors) > 0 {
		result.Color = colors[0]
	}
	return result
}

func getCellNumberFormat(numFmtID int, numberFormats map[int]string) (int, *string) {
	if format, ok := numberFormats[numFmtID]; ok {
		return 0, &format
	}
	if numFmtID >= 0 && numFmtID < 164 {
		return numFmtID, nil
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

func getCellFillColor(color *xlsxColor, themeColors []string, indexedColors map[int]string) []string {
	if color == nil || color.Auto {
		return nil
	}

	rgb := color.RGB
	if color.Theme != nil {
		index := *color.Theme
		if index < 0 || index >= len(themeColors) {
			return nil
		}
		rgb = themeColors[index]
	} else if rgb == "" {
		rgb = indexedColor(indexedColors, color.Indexed)
	}

	rgb = normalizeRGB(rgb)
	if rgb == "" {
		return nil
	}
	if color.Tint != 0 {
		rgb = normalizeRGB(excelize.ThemeColor(strings.TrimPrefix(rgb, "#"), color.Tint))
	}
	return []string{rgb}
}

func normalizeRGB(rgb string) string {
	rgb = strings.TrimPrefix(strings.ToUpper(rgb), "#")
	if len(rgb) == 8 {
		rgb = rgb[2:]
	}
	if len(rgb) != 6 {
		return ""
	}
	return "#" + rgb
}

func indexedColor(indexedColors map[int]string, index int) string {
	if index < 0 {
		return ""
	}
	if color, ok := indexedColors[index]; ok {
		return color
	}
	if index < len(indexedColorMapping) {
		return indexedColorMapping[index]
	}
	return ""
}

func readStyleMetadata(buffer bytes.Buffer) ([]string, map[int]string, map[int]string) {
	reader, err := zip.NewReader(bytes.NewReader(buffer.Bytes()), int64(buffer.Len()))
	if err != nil {
		return nil, nil, nil
	}

	var themeColors []string
	var numberFormats map[int]string
	var indexedColors map[int]string
	for _, entry := range reader.File {
		if entry.Name != "xl/theme/theme1.xml" && entry.Name != "xl/styles.xml" {
			continue
		}
		file, err := entry.Open()
		if err != nil {
			continue
		}
		data, readErr := io.ReadAll(file)
		_ = file.Close()
		if readErr != nil {
			continue
		}
		switch entry.Name {
		case "xl/theme/theme1.xml":
			themeColors = parseThemeColors(data)
		case "xl/styles.xml":
			numberFormats = parseNumberFormats(data)
			indexedColors = parseIndexedColors(data)
		}
	}
	return themeColors, numberFormats, indexedColors
}

func parseThemeColors(data []byte) []string {
	type colorChoice struct {
		SysColor *struct {
			LastColor string `xml:"lastClr,attr"`
		} `xml:"sysClr"`
		SRGBColor *struct {
			Value string `xml:"val,attr"`
		} `xml:"srgbClr"`
	}
	var theme struct {
		Elements struct {
			Scheme struct {
				Dk1      colorChoice `xml:"dk1"`
				Lt1      colorChoice `xml:"lt1"`
				Dk2      colorChoice `xml:"dk2"`
				Lt2      colorChoice `xml:"lt2"`
				Accent1  colorChoice `xml:"accent1"`
				Accent2  colorChoice `xml:"accent2"`
				Accent3  colorChoice `xml:"accent3"`
				Accent4  colorChoice `xml:"accent4"`
				Accent5  colorChoice `xml:"accent5"`
				Accent6  colorChoice `xml:"accent6"`
				Hlink    colorChoice `xml:"hlink"`
				FolHlink colorChoice `xml:"folHlink"`
			} `xml:"clrScheme"`
		} `xml:"themeElements"`
	}
	if xml.Unmarshal(data, &theme) != nil {
		return nil
	}
	choices := []colorChoice{
		theme.Elements.Scheme.Lt1, theme.Elements.Scheme.Dk1,
		theme.Elements.Scheme.Lt2, theme.Elements.Scheme.Dk2,
		theme.Elements.Scheme.Accent1, theme.Elements.Scheme.Accent2,
		theme.Elements.Scheme.Accent3, theme.Elements.Scheme.Accent4,
		theme.Elements.Scheme.Accent5, theme.Elements.Scheme.Accent6,
		theme.Elements.Scheme.Hlink, theme.Elements.Scheme.FolHlink,
	}
	colors := make([]string, len(choices))
	for index, choice := range choices {
		switch {
		case choice.SysColor != nil:
			colors[index] = choice.SysColor.LastColor
		case choice.SRGBColor != nil:
			colors[index] = choice.SRGBColor.Value
		}
	}
	return colors
}

func parseNumberFormats(data []byte) map[int]string {
	type numberFormat struct {
		ID           int    `xml:"numFmtId,attr"`
		FormatCode   string `xml:"formatCode,attr"`
		FormatCode16 string `xml:"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main formatCode16,attr"`
	}
	var formats struct {
		NumFmts struct {
			NumFmt []numberFormat `xml:"numFmt"`
		} `xml:"numFmts"`
	}
	if xml.Unmarshal(data, &formats) != nil {
		return nil
	}
	result := make(map[int]string, len(formats.NumFmts.NumFmt))
	for _, format := range formats.NumFmts.NumFmt {
		code := format.FormatCode16
		if code == "" {
			code = format.FormatCode
		}
		if code != "" {
			result[format.ID] = code
		}
	}
	return result
}

func parseIndexedColors(data []byte) map[int]string {
	var styles struct {
		Colors struct {
			IndexedColors struct {
				RGBColors []struct {
					RGB string `xml:"rgb,attr"`
				} `xml:"rgbColor"`
			} `xml:"indexedColors"`
		} `xml:"colors"`
	}
	if xml.Unmarshal(data, &styles) != nil {
		return nil
	}
	colors := make(map[int]string, len(styles.Colors.IndexedColors.RGBColors))
	for index, color := range styles.Colors.IndexedColors.RGBColors {
		colors[index] = color.RGB
	}
	return colors
}

var indexedColorMapping = []string{
	"000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
	"000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
	"800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
	"9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
	"000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
	"00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
	"3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
	"003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
	"000000", "FFFFFF",
}
