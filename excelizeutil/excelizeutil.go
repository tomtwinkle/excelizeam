package excelizeutil

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"reflect"
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

// GetStyle returns the effective style for a cell in an XLSX workbook.
// It returns an error when the workbook, coordinates, sheet, or cell style is invalid.
func GetStyle(excelBuffer bytes.Buffer, sheetName string, corIdx, rowIdx int) (excelize.Style, error) {
	themeColors, numberFormats, indexedColors := readStyleMetadata(excelBuffer)
	f, err := excelize.OpenReader(&excelBuffer)
	if err != nil {
		return excelize.Style{}, fmt.Errorf("open workbook: %w", err)
	}
	defer func() { _ = f.Close() }()

	cell, err := excelize.CoordinatesToCellName(corIdx, rowIdx)
	if err != nil {
		return excelize.Style{}, fmt.Errorf("convert cell coordinates: %w", err)
	}
	styleID, err := f.GetCellStyle(sheetName, cell)
	if err != nil {
		return excelize.Style{}, fmt.Errorf("get cell style: %w", err)
	}

	format, err := resolveCellFormat(f, styleID)
	if err != nil {
		return excelize.Style{}, fmt.Errorf("resolve cell style: %w", err)
	}
	numFmt, customNumFmt := getCellNumberFormat(format.NumFmtID, numberFormats)
	style := excelize.Style{
		Border:       getCellBorder(f, format.BorderID, themeColors, indexedColors),
		Fill:         getCellFill(f, format.FillID, themeColors, indexedColors),
		Font:         getCellFont(f, format.FontID, format.HasFont, themeColors, indexedColors),
		Alignment:    format.Alignment,
		Protection:   format.Protection,
		NumFmt:       numFmt,
		CustomNumFmt: customNumFmt,
	}
	applyNativeNumberFormatMetadata(&style, f, styleID, format, customNumFmt)
	return style, nil
}

type resolvedCellFormat struct {
	FontID     int
	HasFont    bool
	FillID     int
	BorderID   int
	NumFmtID   int
	Alignment  *excelize.Alignment
	Protection *excelize.Protection
}

func resolveCellFormat(f *excelize.File, styleID int) (resolvedCellFormat, error) {
	if f.Styles == nil || f.Styles.CellXfs == nil || styleID < 0 || styleID >= len(f.Styles.CellXfs.Xf) {
		return resolvedCellFormat{}, fmt.Errorf("invalid cell style ID %d", styleID)
	}
	cellXF := f.Styles.CellXfs.Xf[styleID]
	resolved := resolvedCellFormat{
		FontID:   integerValue(cellXF.FontID),
		HasFont:  inheritedStylePresent(cellXF.FontID, nil, cellXF.ApplyFont),
		FillID:   integerValue(cellXF.FillID),
		BorderID: integerValue(cellXF.BorderID),
		NumFmtID: integerValue(cellXF.NumFmtID),
	}
	alignment := cellXF.Alignment
	protection := cellXF.Protection
	if cellXF.XfID != nil && f.Styles.CellStyleXfs != nil && *cellXF.XfID >= 0 && *cellXF.XfID < len(f.Styles.CellStyleXfs.Xf) {
		baseXF := f.Styles.CellStyleXfs.Xf[*cellXF.XfID]
		resolved.FontID = inheritedStyleID(cellXF.FontID, baseXF.FontID, cellXF.ApplyFont)
		resolved.HasFont = inheritedStylePresent(cellXF.FontID, baseXF.FontID, cellXF.ApplyFont)
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
	return resolved, nil
}

func inheritedStylePresent(child, parent *int, apply *bool) bool {
	if apply != nil && *apply {
		// fontId defaults to zero in CT_Xf, so apply=true with no explicit
		// fontId still applies the default (possibly customized) font record.
		return true
	}
	if parent != nil {
		return true
	}
	if apply != nil { // apply=false with no base XF means no child style is active.
		return false
	}
	return child != nil && *child != 0
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
	if f.Styles.Borders == nil || borderID < 0 || borderID >= len(f.Styles.Borders.Border) || f.Styles.Borders.Border[borderID] == nil {
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

	topStyle, topColor := borderLine(definition.Top)
	bottomStyle, bottomColor := borderLine(definition.Bottom)
	leftStyle, leftColor := borderLine(definition.Left)
	rightStyle, rightColor := borderLine(definition.Right)
	appendBorder("top", topStyle, topColor)
	appendBorder("bottom", bottomStyle, bottomColor)
	appendBorder("left", leftStyle, leftColor)
	appendBorder("right", rightStyle, rightColor)
	if definition.DiagonalUp {
		diagonalStyle, diagonalColor := borderLine(definition.Diagonal)
		appendBorder("diagonalUp", diagonalStyle, diagonalColor)
	}
	if definition.DiagonalDown {
		diagonalStyle, diagonalColor := borderLine(definition.Diagonal)
		appendBorder("diagonalDown", diagonalStyle, diagonalColor)
	}
	return borders
}

func borderLine(line any) (string, *xlsxColor) {
	value := reflect.ValueOf(line)
	if !value.IsValid() {
		return "", nil
	}
	if value.Kind() == reflect.Ptr {
		if value.IsNil() {
			return "", nil
		}
		value = value.Elem()
	}
	if value.Kind() != reflect.Struct {
		return "", nil
	}
	style := value.FieldByName("Style")
	if !style.IsValid() || style.Kind() != reflect.String {
		return "", nil
	}
	return style.String(), colorValue(value.FieldByName("Color"))
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
	if f.Styles.Fills == nil || fillID < 0 || fillID >= len(f.Styles.Fills.Fill) || f.Styles.Fills.Fill[fillID] == nil {
		return excelize.Fill{}
	}
	definition := f.Styles.Fills.Fill[fillID]
	if definition.GradientFill != nil {
		return getCellGradientFill(definition.GradientFill, themeColors, indexedColors)
	}
	if definition.PatternFill == nil {
		return excelize.Fill{}
	}
	patternFill := definition.PatternFill

	pattern := 0
	for i, name := range fillPatterns {
		if name == patternFill.PatternType {
			pattern = i
			break
		}
	}

	color := patternFill.FgColor
	if color == nil {
		color = patternFill.BgColor
	}
	return excelize.Fill{
		Type:    "pattern",
		Pattern: pattern,
		Color:   getCellFillColor((*xlsxColor)(color), themeColors, indexedColors),
	}
}

type gradientGeometry struct {
	typeName  string
	degree    float64
	left      float64
	right     float64
	top       float64
	bottom    float64
	positions []float64
}

func getCellGradientFill(value any, themeColors []string, indexedColors map[int]string) excelize.Fill {
	geometry, colors, ok := readGradient(value)
	if !ok {
		return excelize.Fill{Type: "gradient"}
	}
	for shading, preset := range gradientGeometries() {
		if !sameGradientGeometry(geometry, preset) || len(colors) != len(preset.positions) {
			continue
		}
		matchesPositions := true
		for index, position := range preset.positions {
			if colors[index].position != position {
				matchesPositions = false
				break
			}
		}
		if !matchesPositions {
			continue
		}
		resolved := make([][]string, len(colors))
		for index, stop := range colors {
			resolved[index] = getCellFillColor(stop.color, themeColors, indexedColors)
			if len(resolved[index]) == 0 {
				return excelize.Fill{Type: "gradient"}
			}
		}
		if len(colors) == 3 && !reflect.DeepEqual(resolved[0], resolved[2]) {
			// Excelize's public API models these presets as color0, color1,
			// color0. A different final stop cannot be represented.
			return excelize.Fill{Type: "gradient"}
		}
		return excelize.Fill{Type: "gradient", Shading: shading, Color: []string{resolved[0][0], resolved[1][0]}}
	}
	return excelize.Fill{Type: "gradient"}
}

func readGradient(value any) (gradientGeometry, []struct {
	position float64
	color    *xlsxColor
}, bool) {
	ref := reflect.ValueOf(value)
	if !ref.IsValid() {
		return gradientGeometry{}, nil, false
	}
	if ref.Kind() == reflect.Ptr {
		if ref.IsNil() {
			return gradientGeometry{}, nil, false
		}
		ref = ref.Elem()
	}
	if ref.Kind() != reflect.Struct {
		return gradientGeometry{}, nil, false
	}
	getFloat := func(name string) float64 {
		field := ref.FieldByName(name)
		if field.IsValid() && field.CanFloat() {
			return field.Float()
		}
		return 0
	}
	getString := func(name string) string {
		field := ref.FieldByName(name)
		if field.IsValid() && field.Kind() == reflect.String {
			return field.String()
		}
		return ""
	}
	geometry := gradientGeometry{
		typeName: getString("Type"),
		degree:   getFloat("Degree"),
		left:     getFloat("Left"),
		right:    getFloat("Right"),
		top:      getFloat("Top"),
		bottom:   getFloat("Bottom"),
	}
	stops := ref.FieldByName("Stop")
	if !stops.IsValid() || stops.Kind() != reflect.Slice {
		return gradientGeometry{}, nil, false
	}
	colors := make([]struct {
		position float64
		color    *xlsxColor
	}, 0, stops.Len())
	for index := 0; index < stops.Len(); index++ {
		stop := stops.Index(index)
		if stop.Kind() == reflect.Ptr {
			if stop.IsNil() {
				return gradientGeometry{}, nil, false
			}
			stop = stop.Elem()
		}
		position := stop.FieldByName("Position")
		color := stop.FieldByName("Color")
		if !position.IsValid() || !position.CanFloat() || !color.IsValid() {
			return gradientGeometry{}, nil, false
		}
		colors = append(colors, struct {
			position float64
			color    *xlsxColor
		}{position: position.Float(), color: colorValue(color)})
	}
	geometry.positions = make([]float64, len(colors))
	for index, stop := range colors {
		geometry.positions[index] = stop.position
	}
	return geometry, colors, true
}

func gradientGeometries() []gradientGeometry {
	if !hasExcelizeField(reflect.TypeOf(excelize.Fill{}), "Transparency") {
		return []gradientGeometry{
			{degree: 90, positions: []float64{0, 1}},
			{positions: []float64{0, 1}},
			{degree: 45, positions: []float64{0, 1}},
			{degree: 135, positions: []float64{0, 1}},
			{typeName: "path", positions: []float64{0, 1}},
			{typeName: "path", left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, positions: []float64{0, 1}},
		}
	}
	return []gradientGeometry{
		{degree: 90, positions: []float64{0, 1}},
		{degree: 270, positions: []float64{0, 1}},
		{degree: 90, positions: []float64{0, 0.5, 1}},
		{positions: []float64{0, 1}},
		{degree: 180, positions: []float64{0, 1}},
		{positions: []float64{0, 0.5, 1}},
		{degree: 45, positions: []float64{0, 1}},
		{degree: 255, positions: []float64{0, 1}},
		{degree: 45, positions: []float64{0, 0.5, 1}},
		{degree: 135, positions: []float64{0, 1}},
		{degree: 315, positions: []float64{0, 1}},
		{degree: 135, positions: []float64{0, 0.5, 1}},
		{typeName: "path", positions: []float64{0, 1}},
		{typeName: "path", left: 1, right: 1, positions: []float64{0, 1}},
		{typeName: "path", bottom: 1, top: 1, positions: []float64{0, 1}},
		{typeName: "path", left: 1, right: 1, top: 1, bottom: 1, positions: []float64{0, 1}},
		{typeName: "path", left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, positions: []float64{0, 1}},
	}
}

func hasExcelizeField(typ reflect.Type, fieldName string) bool {
	_, ok := typ.FieldByName(fieldName)
	return ok
}

func sameGradientGeometry(a, b gradientGeometry) bool {
	return a.typeName == b.typeName && a.degree == b.degree && a.left == b.left && a.right == b.right && a.top == b.top && a.bottom == b.bottom
}

func getCellFont(f *excelize.File, fontID int, hasFont bool, themeColors []string, indexedColors map[int]string) *excelize.Font {
	if f.Styles.Fonts == nil || fontID < 0 || fontID >= len(f.Styles.Fonts.Font) || (!hasFont && fontID == 0) || f.Styles.Fonts.Font[fontID] == nil {
		return nil
	}
	font := f.Styles.Fonts.Font[fontID]
	result := &excelize.Font{}
	// Excelize v2.10.1 treats ColorIndexed's zero value as an explicit
	// indexed color. Use -1 unless the OOXML color actually selects an index.
	setExcelizeField(result, "ColorIndexed", -1)
	result.Bold = xmlBooleanValue(font.B)
	result.Italic = xmlBooleanValue(font.I)
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
	result.Strike = xmlBooleanValue(font.Strike)
	color := colorValue(reflect.ValueOf(font.Color))
	if color != nil {
		if hasExcelizeField(reflect.TypeOf(excelize.Font{}), "ColorTint") {
			// v2.10.1 exposes tint and the theme/indexed selectors separately.
			// Keep the untinted RGB fallback; returning a resolved tint here too
			// would cause NewStyle to serialize and apply the tint twice.
			result.Color = normalizeRGB(color.RGB)
		} else if colors := getCellFillColor(color, themeColors, indexedColors); len(colors) > 0 {
			// Older Style APIs lack selector and tint fields, so keep appearance.
			result.Color = colors[0]
		}
	}
	if color != nil {
		if color.RGB == "" && color.Theme == nil && !color.Auto {
			setExcelizeField(result, "ColorIndexed", color.Indexed)
		}
		if color.Theme != nil {
			setExcelizeField(result, "ColorTheme", color.Theme)
		}
		setExcelizeField(result, "ColorTint", color.Tint)
	}
	if charset, ok := xmlAttributeValue(font, "Charset", "Val"); ok {
		setExcelizeField(result, "Charset", charset)
	}
	if vertical, ok := xmlAttributeValue(font, "VertAlign", "Val"); ok {
		setExcelizeField(result, "VertAlign", vertical)
	}
	return result
}

func xmlBooleanValue(value any) bool {
	ref := reflect.ValueOf(value)
	if !ref.IsValid() || ref.Kind() != reflect.Ptr || ref.IsNil() {
		return false
	}
	attribute := ref.Elem().FieldByName("Val")
	if !attribute.IsValid() || (attribute.Kind() == reflect.Ptr && attribute.IsNil()) {
		return true
	}
	if attribute.Kind() == reflect.Ptr {
		attribute = attribute.Elem()
	}
	return attribute.Bool()
}

func xmlAttributeValue(value any, elementName, attributeName string) (any, bool) {
	ref := reflect.ValueOf(value)
	if !ref.IsValid() {
		return nil, false
	}
	if ref.Kind() == reflect.Ptr {
		if ref.IsNil() {
			return nil, false
		}
		ref = ref.Elem()
	}
	element := ref.FieldByName(elementName)
	if !element.IsValid() || element.Kind() != reflect.Ptr || element.IsNil() {
		return nil, false
	}
	attribute := element.Elem().FieldByName(attributeName)
	if !attribute.IsValid() {
		return nil, false
	}
	if attribute.Kind() == reflect.Ptr {
		if attribute.IsNil() {
			return nil, false
		}
		attribute = attribute.Elem()
	}
	return attribute.Interface(), true
}

func colorValue(value reflect.Value) *xlsxColor {
	if !value.IsValid() {
		return nil
	}
	if value.Kind() == reflect.Ptr {
		if value.IsNil() {
			return nil
		}
		value = value.Elem()
	}
	if value.Kind() != reflect.Struct {
		return nil
	}
	result := &xlsxColor{}
	if field := value.FieldByName("Auto"); field.IsValid() && field.Kind() == reflect.Bool {
		result.Auto = field.Bool()
	}
	if field := value.FieldByName("RGB"); field.IsValid() && field.Kind() == reflect.String {
		result.RGB = field.String()
	}
	if field := value.FieldByName("Indexed"); field.IsValid() && field.Kind() == reflect.Int {
		result.Indexed = int(field.Int())
	}
	if field := value.FieldByName("Theme"); field.IsValid() && field.Kind() == reflect.Ptr && !field.IsNil() {
		theme := int(field.Elem().Int())
		result.Theme = &theme
	}
	if field := value.FieldByName("Tint"); field.IsValid() && field.Kind() == reflect.Float64 {
		result.Tint = field.Float()
	}
	return result
}

func setExcelizeField(target any, name string, value any) bool {
	ref := reflect.ValueOf(target)
	if !ref.IsValid() || ref.Kind() != reflect.Ptr || ref.IsNil() {
		return false
	}
	field := ref.Elem().FieldByName(name)
	if !field.IsValid() || !field.CanSet() {
		return false
	}
	input := reflect.ValueOf(value)
	if field.Kind() == reflect.Ptr && input.Type().AssignableTo(field.Type().Elem()) {
		pointer := reflect.New(field.Type().Elem())
		pointer.Elem().Set(input)
		field.Set(pointer)
		return true
	}
	if input.Type().AssignableTo(field.Type()) {
		field.Set(input)
		return true
	}
	if input.Type().ConvertibleTo(field.Type()) {
		field.Set(input.Convert(field.Type()))
		return true
	}
	return false
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

type excelizeStyleGetter interface {
	GetStyle(int) (*excelize.Style, error)
}

func applyNativeNumberFormatMetadata(style *excelize.Style, f *excelize.File, styleID int, resolved resolvedCellFormat, customNumFmt *string) {
	if customNumFmt != nil || f.Styles.CellXfs == nil || styleID < 0 || styleID >= len(f.Styles.CellXfs.Xf) {
		return
	}
	cellXF := f.Styles.CellXfs.Xf[styleID]
	if cellXF.NumFmtID == nil || *cellXF.NumFmtID != resolved.NumFmtID {
		// A named-style XF can supply the effective ID. Its custom code is
		// retained above; do not infer currency metadata from format text.
		return
	}
	getter, ok := any(f).(excelizeStyleGetter)
	if !ok {
		return
	}
	nativeStyle, err := getter.GetStyle(styleID)
	if err != nil || nativeStyle == nil {
		return
	}
	if decimalPlaces, ok := excelizePublicFieldValue(nativeStyle, "DecimalPlaces"); ok {
		setExcelizeField(style, "DecimalPlaces", decimalPlaces)
	}
	if negativeRed, ok := excelizePublicFieldValue(nativeStyle, "NegRed"); ok {
		setExcelizeField(style, "NegRed", negativeRed)
	}
}

func excelizePublicFieldValue(target any, name string) (any, bool) {
	ref := reflect.ValueOf(target)
	if !ref.IsValid() {
		return nil, false
	}
	if ref.Kind() == reflect.Ptr {
		if ref.IsNil() {
			return nil, false
		}
		ref = ref.Elem()
	}
	field := ref.FieldByName(name)
	if !field.IsValid() {
		return nil, false
	}
	if field.Kind() == reflect.Ptr {
		if field.IsNil() {
			return nil, false
		}
		field = field.Elem()
	}
	return field.Interface(), true
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
