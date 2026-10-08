package excelizeutil

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"reflect"
	"regexp"
	"sort"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestExcelizeV2101StyleFieldInventory(t *testing.T) {
	styleFields := publicFieldNames(reflect.TypeOf(excelize.Style{}))
	fontFields := publicFieldNames(reflect.TypeOf(excelize.Font{}))
	fillFields := publicFieldNames(reflect.TypeOf(excelize.Fill{}))
	if !containsString(styleFields, "DecimalPlaces") || !containsString(fontFields, "ColorTheme") {
		t.Skip("the linked Excelize version predates the v2.10.1 Style API")
	}

	assert.Equal(t, []string{"Alignment", "Border", "CustomNumFmt", "DecimalPlaces", "Fill", "Font", "NegRed", "NumFmt", "Protection"}, styleFields)
	assert.Equal(t, []string{"Bold", "Charset", "Color", "ColorIndexed", "ColorTheme", "ColorTint", "Family", "Italic", "Size", "Strike", "Underline", "VertAlign"}, fontFields)
	assert.Equal(t, []string{"Color", "Pattern", "Shading", "Transparency", "Type"}, fillFields)
	assert.Equal(t, []string{"Color", "Style", "Type"}, publicFieldNames(reflect.TypeOf(excelize.Border{})))
	assert.Equal(t, []string{"Hidden", "Locked"}, publicFieldNames(reflect.TypeOf(excelize.Protection{})))
	assert.Equal(t, []string{"Horizontal", "Indent", "JustifyLastLine", "ReadingOrder", "RelativeIndent", "ShrinkToFit", "TextRotation", "Vertical", "WrapText"}, publicFieldNames(reflect.TypeOf(excelize.Alignment{})))
}

func TestGetStyleRoundTripsAllGradientPresets(t *testing.T) {
	maxShading := 16
	if !hasPublicField(reflect.TypeOf(excelize.Fill{}), "Transparency") {
		maxShading = 5
	}
	for shading := 0; shading <= maxShading; shading++ {
		t.Run(fmt.Sprintf("shading-%d", shading), func(t *testing.T) {
			style := &excelize.Style{Fill: excelize.Fill{Type: "gradient", Shading: shading, Color: []string{"#123456", "#ABCDEF"}}}
			original := workbookWithStyle(t, style)
			got := GetStyle(original, "Sheet1", 1, 1)
			require.Equal(t, "gradient", got.Fill.Type)
			require.Equal(t, shading, got.Fill.Shading)
			require.Equal(t, []string{"#123456", "#ABCDEF"}, got.Fill.Color)

			reapplied := workbookWithStyle(t, &got)
			reread := GetStyle(reapplied, "Sheet1", 1, 1)
			assert.Equal(t, got.Fill, reread.Fill)
			assert.Equal(t, styleGradientXML(t, original), styleGradientXML(t, reapplied), "the saved OOXML gradient geometry and stops must survive reapplication")
		})
	}
}

func TestGetStyleRoundTripsEveryPublicFontField(t *testing.T) {
	if !hasPublicField(reflect.TypeOf(excelize.Font{}), "ColorTheme") {
		t.Skip("the linked Excelize version predates the v2.10.1 Font API")
	}
	tests := []struct {
		name      string
		field     string
		value     any
		want      any
		xmlMarker string
	}{
		{name: "Bold", field: "Bold", value: true, want: true, xmlMarker: `<b val="1"`},
		{name: "Italic", field: "Italic", value: true, want: true, xmlMarker: `<i val="1"`},
		{name: "Underline", field: "Underline", value: "double", want: "double", xmlMarker: `val="double"`},
		{name: "Family", field: "Family", value: "Aptos", want: "Aptos", xmlMarker: `name val="Aptos"`},
		{name: "Size", field: "Size", value: float64(13), want: float64(13), xmlMarker: `sz val="13"`},
		{name: "Strike", field: "Strike", value: true, want: true, xmlMarker: `<strike val="1"`},
		{name: "Color", field: "Color", value: "#123456", want: "#123456", xmlMarker: `rgb="FF123456"`},
		{name: "ColorIndexed", field: "ColorIndexed", value: 7, want: 7, xmlMarker: `indexed="7"`},
		{name: "ColorTheme", field: "ColorTheme", value: 4, want: 4, xmlMarker: `theme="4"`},
		{name: "ColorTint", field: "ColorTint", value: float64(0.25), want: float64(0.25), xmlMarker: `tint="0.25"`},
		{name: "VertAlign", field: "VertAlign", value: "superscript", want: "superscript", xmlMarker: `vertAlign val="superscript"`},
		{name: "Charset", field: "Charset", value: 204, want: 204, xmlMarker: `charset val="204"`},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			font := &excelize.Font{}
			require.True(t, setPublicField(font, test.field, test.value), "field %s must exist", test.field)
			style := &excelize.Style{Font: font}
			original := workbookWithStyle(t, style)
			got := GetStyle(original, "Sheet1", 1, 1)
			require.NotNil(t, got.Font)
			gotValue := publicFieldValue(got.Font, test.field)
			assert.Equal(t, test.want, gotValue)

			reapplied := workbookWithStyle(t, &got)
			reread := GetStyle(reapplied, "Sheet1", 1, 1)
			require.NotNil(t, reread.Font)
			rereadValue := publicFieldValue(reread.Font, test.field)
			assert.Equal(t, test.want, rereadValue)
			assert.Contains(t, xlsxEntry(t, reapplied, "xl/styles.xml"), test.xmlMarker)
		})
	}
}

func TestGetStyleRoundTripsEveryAlignmentFieldInXML(t *testing.T) {
	alignment := &excelize.Alignment{
		Horizontal:      "center",
		Indent:          2,
		JustifyLastLine: true,
		ReadingOrder:    1,
		RelativeIndent:  3,
		ShrinkToFit:     true,
		TextRotation:    45,
		Vertical:        "center",
		WrapText:        true,
	}
	original := workbookWithStyle(t, &excelize.Style{Alignment: alignment})
	got := GetStyle(original, "Sheet1", 1, 1)
	assert.Equal(t, alignment, got.Alignment)
	assert.Equal(t, alignmentStyleXML{
		Horizontal: "center", Indent: "2", JustifyLastLine: "true", ReadingOrder: "1",
		RelativeIndent: "3", ShrinkToFit: "true", TextRotation: "45", Vertical: "center", WrapText: "true",
	}, styleAlignmentXML(t, original))

	reapplied := workbookWithStyle(t, &got)
	reread := GetStyle(reapplied, "Sheet1", 1, 1)
	assert.Equal(t, alignment, reread.Alignment)
	assert.Equal(t, styleAlignmentXML(t, original), styleAlignmentXML(t, reapplied), "all nine alignment attributes must survive in saved OOXML")
}

func TestGetStylePreservesProtectionDefaultAndExplicitRelease(t *testing.T) {
	t.Run("omitted attributes use OOXML defaults", func(t *testing.T) {
		buffer := workbookWithStyle(t, &excelize.Style{Protection: &excelize.Protection{Hidden: true, Locked: false}})
		buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
			return stripProtectionAttributes(t, data)
		})
		assertProtectionAttrsOmitted(t, styleProtectionXML(t, buffer))

		got := GetStyle(buffer, "Sheet1", 1, 1)
		assert.Equal(t, &excelize.Protection{Locked: true, Hidden: false}, got.Protection)
		reapplied := workbookWithStyle(t, &got)
		reread := GetStyle(reapplied, "Sheet1", 1, 1)
		assert.Equal(t, got.Protection, reread.Protection)
		attrs := styleProtectionXML(t, reapplied)
		assertXMLBool(t, attrs.Locked, true)
		assertXMLBool(t, attrs.Hidden, false)
	})

	t.Run("explicit unlock survives", func(t *testing.T) {
		protection := &excelize.Protection{Hidden: true, Locked: false}
		original := workbookWithStyle(t, &excelize.Style{Protection: protection})
		got := GetStyle(original, "Sheet1", 1, 1)
		assert.Equal(t, protection, got.Protection)
		assertXMLBool(t, styleProtectionXML(t, original).Locked, false)
		assertXMLBool(t, styleProtectionXML(t, original).Hidden, true)

		reapplied := workbookWithStyle(t, &got)
		reread := GetStyle(reapplied, "Sheet1", 1, 1)
		assert.Equal(t, protection, reread.Protection)
		assertXMLBool(t, styleProtectionXML(t, reapplied).Locked, false)
		assertXMLBool(t, styleProtectionXML(t, reapplied).Hidden, true)
	})
}

func TestGetStylePreservesFontColorIdentityAndResolvedRGBSeparately(t *testing.T) {
	if !hasPublicField(reflect.TypeOf(excelize.Font{}), "ColorTheme") {
		t.Skip("the linked Excelize version predates the v2.10.1 Font API")
	}

	themeIndex := 4
	themeFont := &excelize.Font{}
	require.True(t, setPublicField(themeFont, "ColorTheme", themeIndex))
	require.True(t, setPublicField(themeFont, "ColorTint", float64(0.25)))
	source := workbookWithStyle(t, &excelize.Style{Font: themeFont})
	got := GetStyle(source, "Sheet1", 1, 1)
	require.NotNil(t, got.Font)
	assert.Equal(t, themeIndex, publicFieldValue(got.Font, "ColorTheme"))
	assert.Equal(t, float64(0.25), publicFieldValue(got.Font, "ColorTint"))
	assert.Equal(t, -1, publicFieldValue(got.Font, "ColorIndexed"))
	wantThemeRGB := normalizeRGB(excelize.ThemeColor("5B9BD5", 0.25))
	assert.Empty(t, got.Font.Color, "theme-only color keeps its selector instead of resolving to a second RGB color")

	// The target workbook has the same default theme, so retaining the theme
	// selector and tint preserves both the reference and rendered color.
	sameTheme := workbookWithStyle(t, &got)
	assert.Equal(t, xlsxEntry(t, source, "xl/theme/theme1.xml"), xlsxEntry(t, sameTheme, "xl/theme/theme1.xml"))
	themeColor := styleFontColors(t, sameTheme)[1]
	assert.Equal(t, "", themeColor.RGB)
	assert.Nil(t, themeColor.Indexed)
	require.NotNil(t, themeColor.Theme)
	assert.Equal(t, themeIndex, *themeColor.Theme)
	assert.Equal(t, 0.25, themeColor.Tint)
	assert.Equal(t, got.Font.Color, GetStyle(sameTheme, "Sheet1", 1, 1).Font.Color)
	assert.NotEmpty(t, wantThemeRGB, "resolved appearance color is computed separately from selector identity")

	// A caller can choose appearance-only preservation by applying the resolved
	// RGB with theme metadata cleared.
	rgbFont := excelize.Font{Color: wantThemeRGB}
	require.True(t, setPublicField(&rgbFont, "ColorIndexed", -1))
	rgbOnly := &excelize.Style{Font: &rgbFont}
	rgbWorkbook := workbookWithStyle(t, rgbOnly)
	rgbReread := GetStyle(rgbWorkbook, "Sheet1", 1, 1)
	assert.Equal(t, wantThemeRGB, rgbReread.Font.Color)
	rgbColor := styleFontColors(t, rgbWorkbook)[1]
	assert.Equal(t, "FF"+strings.TrimPrefix(wantThemeRGB, "#"), rgbColor.RGB)
	assert.Nil(t, rgbColor.Theme)
	assert.Nil(t, rgbColor.Indexed)

	indexedFont := &excelize.Font{}
	require.True(t, setPublicField(indexedFont, "ColorIndexed", 7))
	require.True(t, setPublicField(indexedFont, "ColorTint", float64(0.25)))
	indexedSource := workbookWithStyle(t, &excelize.Style{Font: indexedFont})
	indexed := GetStyle(indexedSource, "Sheet1", 1, 1)
	require.NotNil(t, indexed.Font)
	assert.Equal(t, 7, publicFieldValue(indexed.Font, "ColorIndexed"))
	assert.Equal(t, float64(0.25), publicFieldValue(indexed.Font, "ColorTint"))
	assert.Empty(t, indexed.Font.Color)
	indexedTarget := workbookWithStyle(t, &indexed)
	indexedColor := styleFontColors(t, indexedTarget)[1]
	assert.Equal(t, "", indexedColor.RGB)
	require.NotNil(t, indexedColor.Indexed)
	assert.Equal(t, 7, *indexedColor.Indexed)
	assert.Nil(t, indexedColor.Theme)
	assert.Equal(t, 0.25, indexedColor.Tint)
	assert.Equal(t, indexed.Font.Color, GetStyle(indexedTarget, "Sheet1", 1, 1).Font.Color)
	wantIndexedRGB := normalizeRGB(excelize.ThemeColor("00FFFF", 0.25))
	assert.Equal(t, []string{wantIndexedRGB}, resolvedFontColor(t, indexedSource, 1))
	assert.Equal(t, []string{wantIndexedRGB}, resolvedFontColor(t, indexedTarget, 1), "the indexed palette appearance remains the same after selector round-trip")
}

func TestGetStylePreservesRGBBaseAndFontTintOnReapply(t *testing.T) {
	if !hasPublicField(reflect.TypeOf(excelize.Font{}), "ColorTint") {
		t.Skip("the linked Excelize version predates the v2.10.1 Font API")
	}
	font := &excelize.Font{Color: "#123456"}
	require.True(t, setPublicField(font, "ColorIndexed", -1))
	require.True(t, setPublicField(font, "ColorTint", float64(0.25)))
	original := workbookWithStyle(t, &excelize.Style{Font: font})
	got := GetStyle(original, "Sheet1", 1, 1)
	require.NotNil(t, got.Font)
	assert.Equal(t, "#123456", got.Font.Color)
	assert.Equal(t, -1, publicFieldValue(got.Font, "ColorIndexed"))
	assert.Equal(t, float64(0.25), publicFieldValue(got.Font, "ColorTint"))

	reapplied := workbookWithStyle(t, &got)
	color := styleFontColors(t, reapplied)[1]
	assert.Equal(t, "FF123456", color.RGB)
	assert.Nil(t, color.Indexed)
	assert.Nil(t, color.Theme)
	assert.Equal(t, 0.25, color.Tint)
}

func TestGetStyleDoesNotTurnAutomaticFontColorIntoIndexedBlack(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{Font: &excelize.Font{Color: "#123456"}})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		old := []byte(`rgb="FF123456"`)
		require.Equal(t, 1, bytes.Count(data, old))
		return bytes.Replace(data, old, []byte(`auto="1"`), 1)
	})

	got := GetStyle(buffer, "Sheet1", 1, 1)
	require.NotNil(t, got.Font)
	assert.Empty(t, got.Font.Color)
	if hasPublicField(reflect.TypeOf(excelize.Font{}), "ColorIndexed") {
		assert.Equal(t, -1, publicFieldValue(got.Font, "ColorIndexed"))
	}

	reapplied := workbookWithStyle(t, &got)
	assert.NotContains(t, xlsxEntry(t, reapplied, "xl/styles.xml"), `indexed="0"`, "the API cannot represent automatic color, but must not silently turn it black")
	if hasPublicField(reflect.TypeOf(excelize.Font{}), "ColorIndexed") {
		assert.Equal(t, -1, publicFieldValue(GetStyle(reapplied, "Sheet1", 1, 1).Font, "ColorIndexed"))
	}
}

func TestGetStylePreservesExplicitDefaultFontAndEmptyBooleanValues(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{Font: &excelize.Font{Color: "#123456"}})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		fontStart := bytes.Index(data, []byte("<font>"))
		require.NotEqual(t, -1, fontStart)
		fontEnd := bytes.Index(data[fontStart:], []byte("</font>"))
		require.NotEqual(t, -1, fontEnd)
		fontEnd += fontStart + len("</font>")
		font := append([]byte(nil), data[fontStart:fontEnd]...)
		font = bytes.Replace(font, []byte("<font>"), []byte("<font><b/><i/><strike/>"), 1)
		data = append(append(append([]byte(nil), data[:fontStart]...), font...), data[fontEnd:]...)
		oldXF := []byte(`fontId="1"`)
		require.Equal(t, 1, bytes.Count(data, oldXF))
		return bytes.Replace(data, oldXF, []byte(`fontId="0" applyFont="1"`), 1)
	})

	got := GetStyle(buffer, "Sheet1", 1, 1)
	require.NotNil(t, got.Font, "an explicitly applied fontId=0 is a real style, not an absent font")
	assert.True(t, got.Font.Bold, "an empty <b/> means true")
	assert.True(t, got.Font.Italic, "an empty <i/> means true")
	assert.True(t, got.Font.Strike, "an empty <strike/> means true")
	assert.Contains(t, xlsxEntry(t, buffer, "xl/styles.xml"), `applyFont="1"`)
}

func TestGetStyleHandlesSparseBorderXML(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{Border: []excelize.Border{{Type: "top", Style: 1, Color: "#112233"}}})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		data = bytes.Replace(data, []byte(`<right></right><top></top><bottom></bottom><diagonal></diagonal>`), nil, 1)
		data = bytes.Replace(data, []byte(`<right/><top/><bottom/><diagonal/>`), nil, 1)
		return data
	})
	var got excelize.Style
	require.NotPanics(t, func() { got = GetStyle(buffer, "Sheet1", 1, 1) })
	assert.Equal(t, []excelize.Border{{Type: "top", Style: 1, Color: "#112233"}}, got.Border)
	assert.Equal(t, got.Border, GetStyle(workbookWithStyle(t, &got), "Sheet1", 1, 1).Border)
}

func TestGetStyleUsesForegroundColorWhenPatternHasBothColors(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: 10, Color: []string{"#123456"}}})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		patternPosition := bytes.Index(data, []byte(`patternType="darkTrellis"`))
		require.NotEqual(t, -1, patternPosition)
		fillClose := bytes.Index(data[patternPosition:], []byte(`</patternFill>`))
		require.NotEqual(t, -1, fillClose)
		fillClose += patternPosition
		data = append(append(append([]byte(nil), data[:fillClose]...), []byte(`<bgColor rgb="FFABCDEF"></bgColor>`)...), data[fillClose:]...)
		return data
	})

	got := GetStyle(buffer, "Sheet1", 1, 1)
	assert.Equal(t, []string{"#123456"}, got.Fill.Color)
	reapplied := workbookWithStyle(t, &got)
	assert.Equal(t, []string{"#123456"}, GetStyle(reapplied, "Sheet1", 1, 1).Fill.Color)
	assert.NotContains(t, xlsxEntry(t, reapplied, "xl/styles.xml"), `FFABCDEF`, "Style.Fill has only one color and cannot encode both pattern colors")
}

func TestGetStyleRoundTripsEveryPatternIDAndBorderStyle(t *testing.T) {
	for pattern := 0; pattern <= 18; pattern++ {
		t.Run(fmt.Sprintf("pattern-%d", pattern), func(t *testing.T) {
			style := &excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: pattern, Color: []string{"#315D3C"}}}
			got := GetStyle(workbookWithStyle(t, style), "Sheet1", 1, 1)
			require.Equal(t, "pattern", got.Fill.Type)
			assert.Equal(t, pattern, got.Fill.Pattern)
			assert.Equal(t, []string{"#315D3C"}, got.Fill.Color)
			reapplied := workbookWithStyle(t, &got)
			assert.Equal(t, got.Fill, GetStyle(reapplied, "Sheet1", 1, 1).Fill)
		})
	}
	for borderStyle := 0; borderStyle <= 13; borderStyle++ {
		t.Run(fmt.Sprintf("border-style-%d", borderStyle), func(t *testing.T) {
			style := &excelize.Style{Border: []excelize.Border{{Type: "top", Style: borderStyle, Color: "#112233"}}}
			got := GetStyle(workbookWithStyle(t, style), "Sheet1", 1, 1)
			require.Equal(t, style.Border, got.Border)
			reapplied := workbookWithStyle(t, &got)
			assert.Equal(t, got.Border, GetStyle(reapplied, "Sheet1", 1, 1).Border)
		})
	}
}

func TestGetStylePreservesNumberFormatIntentWithoutGuessingFormatTokens(t *testing.T) {
	if !hasPublicField(reflect.TypeOf(excelize.Style{}), "DecimalPlaces") {
		t.Skip("the linked Excelize version predates the v2.10.1 Style API")
	}

	// A custom code contains the decimal count and negative-red section already.
	// Keep its exact spelling; do not derive fields by scanning punctuation that
	// could be quoted or escaped.
	code := `"$"#,##0.000;[Red]("$"#,##0.000)`
	style := &excelize.Style{NumFmt: 164, CustomNumFmt: &code, NegRed: true}
	require.True(t, setPublicField(style, "DecimalPlaces", 3))
	original := workbookWithStyle(t, style)
	got := GetStyle(original, "Sheet1", 1, 1)
	require.NotNil(t, got.CustomNumFmt)
	assert.Equal(t, code, *got.CustomNumFmt)
	reapplied := workbookWithStyle(t, &got)
	reread := GetStyle(reapplied, "Sheet1", 1, 1)
	require.NotNil(t, reread.CustomNumFmt)
	assert.Equal(t, code, *reread.CustomNumFmt)
	assert.Contains(t, customNumberFormats(t, reapplied), code)
}

func TestGetStyleMapsBuiltInDecimalPlacesField(t *testing.T) {
	field, ok := reflect.TypeOf(excelize.Style{}).FieldByName("DecimalPlaces")
	if !ok || field.Type.Kind() != reflect.Ptr {
		t.Skip("the linked Excelize version predates the v2.10.1 DecimalPlaces API")
	}
	style := &excelize.Style{NumFmt: 44}
	original := workbookWithStyle(t, style)
	got := GetStyle(original, "Sheet1", 1, 1)
	assert.Equal(t, 44, got.NumFmt)
	assert.Equal(t, 2, publicFieldValue(&got, "DecimalPlaces"))
	assert.Nil(t, got.CustomNumFmt)

	reapplied := workbookWithStyle(t, &got)
	reread := GetStyle(reapplied, "Sheet1", 1, 1)
	assert.Equal(t, 44, reread.NumFmt)
	assert.Equal(t, 2, publicFieldValue(&reread, "DecimalPlaces"))
}

type gradientFillXML struct {
	Type   string            `xml:"type,attr"`
	Degree float64           `xml:"degree,attr"`
	Left   float64           `xml:"left,attr"`
	Right  float64           `xml:"right,attr"`
	Top    float64           `xml:"top,attr"`
	Bottom float64           `xml:"bottom,attr"`
	Stops  []gradientStopXML `xml:"stop"`
}

type fontColorXML struct {
	RGB     string  `xml:"rgb,attr"`
	Auto    bool    `xml:"auto,attr"`
	Indexed *int    `xml:"indexed,attr"`
	Theme   *int    `xml:"theme,attr"`
	Tint    float64 `xml:"tint,attr"`
}

type alignmentStyleXML struct {
	Horizontal      string `xml:"horizontal,attr"`
	Indent          string `xml:"indent,attr"`
	JustifyLastLine string `xml:"justifyLastLine,attr"`
	ReadingOrder    string `xml:"readingOrder,attr"`
	RelativeIndent  string `xml:"relativeIndent,attr"`
	ShrinkToFit     string `xml:"shrinkToFit,attr"`
	TextRotation    string `xml:"textRotation,attr"`
	Vertical        string `xml:"vertical,attr"`
	WrapText        string `xml:"wrapText,attr"`
}

type protectionStyleXML struct {
	Locked *string `xml:"locked,attr"`
	Hidden *string `xml:"hidden,attr"`
}

type cellXFStyleXML struct {
	Alignment  *alignmentStyleXML  `xml:"alignment"`
	Protection *protectionStyleXML `xml:"protection"`
}

func cellXFStylesXML(t *testing.T, buffer bytes.Buffer) []cellXFStyleXML {
	t.Helper()
	var styles struct {
		CellXFs struct {
			XFs []cellXFStyleXML `xml:"xf"`
		} `xml:"cellXfs"`
	}
	require.NoError(t, xml.Unmarshal([]byte(xlsxEntry(t, buffer, "xl/styles.xml")), &styles))
	return styles.CellXFs.XFs
}

func styleAlignmentXML(t *testing.T, buffer bytes.Buffer) alignmentStyleXML {
	t.Helper()
	var alignments []alignmentStyleXML
	for _, xf := range cellXFStylesXML(t, buffer) {
		if xf.Alignment != nil {
			alignments = append(alignments, *xf.Alignment)
		}
	}
	require.Len(t, alignments, 1)
	return alignments[0]
}

func styleProtectionXML(t *testing.T, buffer bytes.Buffer) protectionStyleXML {
	t.Helper()
	var protections []protectionStyleXML
	for _, xf := range cellXFStylesXML(t, buffer) {
		if xf.Protection != nil {
			protections = append(protections, *xf.Protection)
		}
	}
	require.Len(t, protections, 1)
	return protections[0]
}

func assertXMLBool(t *testing.T, value *string, want bool) {
	t.Helper()
	require.NotNil(t, value)
	var got bool
	switch strings.ToLower(*value) {
	case "1", "true":
		got = true
	case "0", "false":
		got = false
	default:
		t.Fatalf("unsupported XML boolean value %q", *value)
	}
	assert.Equal(t, want, got)
}

func stripProtectionAttributes(t *testing.T, data []byte) []byte {
	t.Helper()
	openingTag := regexp.MustCompile(`<protection\b[^>]*>`).Find(data)
	require.NotEmpty(t, openingTag, "fixture must contain a protection element")
	require.NotContains(t, string(openingTag), `/>`, "fixture must retain an explicit empty protection element")
	stripped := regexp.MustCompile(`\s+(?:locked|hidden)="[^"]*"`).ReplaceAll(openingTag, nil)
	require.NotEqual(t, string(openingTag), string(stripped), "fixture must contain explicit protection attributes")
	return bytes.Replace(data, openingTag, []byte(`<protection>`), 1)
}

func assertProtectionAttrsOmitted(t *testing.T, protection protectionStyleXML) {
	t.Helper()
	assert.Nil(t, protection.Locked)
	assert.Nil(t, protection.Hidden)
}

func styleFontColors(t *testing.T, buffer bytes.Buffer) []fontColorXML {
	t.Helper()
	var styles struct {
		Fonts struct {
			Font []struct {
				Color *fontColorXML `xml:"color"`
			} `xml:"font"`
		} `xml:"fonts"`
	}
	require.NoError(t, xml.Unmarshal([]byte(xlsxEntry(t, buffer, "xl/styles.xml")), &styles))
	colors := make([]fontColorXML, 0, len(styles.Fonts.Font))
	for _, font := range styles.Fonts.Font {
		if font.Color != nil {
			colors = append(colors, *font.Color)
		}
	}
	return colors
}

func resolvedFontColor(t *testing.T, buffer bytes.Buffer, fontIndex int) []string {
	t.Helper()
	colors := styleFontColors(t, buffer)
	require.Greater(t, fontIndex, -1)
	require.Less(t, fontIndex, len(colors))
	color := colors[fontIndex]
	themeColors, _, indexedColors := readStyleMetadata(buffer)
	selector := &xlsxColor{Auto: color.Auto, RGB: color.RGB, Tint: color.Tint}
	if color.Indexed != nil {
		selector.Indexed = *color.Indexed
	}
	if color.Theme != nil {
		selector.Theme = color.Theme
	}
	return getCellFillColor(selector, themeColors, indexedColors)
}

type gradientStopXML struct {
	Position float64 `xml:"position,attr"`
	Color    struct {
		RGB     string  `xml:"rgb,attr"`
		Indexed int     `xml:"indexed,attr"`
		Theme   *int    `xml:"theme,attr"`
		Tint    float64 `xml:"tint,attr"`
	} `xml:"color"`
}

func styleGradientXML(t *testing.T, buffer bytes.Buffer) []gradientFillXML {
	t.Helper()
	data := []byte(xlsxEntry(t, buffer, "xl/styles.xml"))
	var styles struct {
		Fills struct {
			Fill []struct {
				Gradient *gradientFillXML `xml:"gradientFill"`
			} `xml:"fill"`
		} `xml:"fills"`
	}
	require.NoError(t, xml.Unmarshal(data, &styles))
	var result []gradientFillXML
	for _, fill := range styles.Fills.Fill {
		if fill.Gradient != nil {
			result = append(result, *fill.Gradient)
		}
	}
	require.Len(t, result, 1)
	return result
}

func xlsxEntry(t *testing.T, buffer bytes.Buffer, name string) string {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(buffer.Bytes()), int64(buffer.Len()))
	require.NoError(t, err)
	for _, file := range reader.File {
		if file.Name != name {
			continue
		}
		entry, err := file.Open()
		require.NoError(t, err)
		data, readErr := io.ReadAll(entry)
		require.NoError(t, readErr)
		require.NoError(t, entry.Close())
		return string(data)
	}
	t.Fatalf("XLSX archive did not contain %q", name)
	return ""
}

func customNumberFormats(t *testing.T, buffer bytes.Buffer) []string {
	t.Helper()
	var styles struct {
		NumFmts struct {
			Formats []struct {
				Code string `xml:"formatCode,attr"`
			} `xml:"numFmt"`
		} `xml:"numFmts"`
	}
	require.NoError(t, xml.Unmarshal([]byte(xlsxEntry(t, buffer, "xl/styles.xml")), &styles))
	result := make([]string, 0, len(styles.NumFmts.Formats))
	for _, format := range styles.NumFmts.Formats {
		result = append(result, format.Code)
	}
	return result
}

func publicFieldNames(typ reflect.Type) []string {
	result := make([]string, 0, typ.NumField())
	for i := 0; i < typ.NumField(); i++ {
		if typ.Field(i).IsExported() {
			result = append(result, typ.Field(i).Name)
		}
	}
	sort.Strings(result)
	return result
}

func hasPublicField(typ reflect.Type, name string) bool {
	_, ok := typ.FieldByName(name)
	return ok
}

func containsString(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}

func setPublicField(target any, name string, value any) bool {
	field := reflect.ValueOf(target).Elem().FieldByName(name)
	if !field.IsValid() || !field.CanSet() {
		return false
	}
	valueRef := reflect.ValueOf(value)
	if field.Kind() == reflect.Pointer && valueRef.Type().AssignableTo(field.Type().Elem()) {
		pointer := reflect.New(field.Type().Elem())
		pointer.Elem().Set(valueRef)
		field.Set(pointer)
		return true
	}
	if valueRef.Type().AssignableTo(field.Type()) {
		field.Set(valueRef)
		return true
	}
	if valueRef.Type().ConvertibleTo(field.Type()) {
		field.Set(valueRef.Convert(field.Type()))
		return true
	}
	return false
}

func publicFieldValue(target any, name string) any {
	field := reflect.ValueOf(target).Elem().FieldByName(name)
	if !field.IsValid() {
		return nil
	}
	if field.Kind() == reflect.Pointer {
		if field.IsNil() {
			return nil
		}
		field = field.Elem()
	}
	return field.Interface()
}
