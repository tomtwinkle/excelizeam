package excelizeutil

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func workbookWithStyle(t *testing.T, style *excelize.Style) bytes.Buffer {
	t.Helper()

	f := excelize.NewFile()
	t.Cleanup(func() { require.NoError(t, f.Close()) })

	styleID, err := f.NewStyle(style)
	require.NoError(t, err)
	require.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", styleID))

	buffer, err := f.WriteToBuffer()
	require.NoError(t, err)
	return *buffer
}

func mustGetStyle(t *testing.T, buffer bytes.Buffer, sheetName string, corIdx, rowIdx int) excelize.Style {
	t.Helper()
	style, err := GetStyle(buffer, sheetName, corIdx, rowIdx)
	require.NoError(t, err)
	return style
}

func TestGetStyleReturnsOpenReaderError(t *testing.T) {
	_, err := GetStyle(bytes.Buffer{}, "Sheet1", 1, 1)
	require.Error(t, err)
	assert.Contains(t, err.Error(), "open workbook")
}

func TestGetStyleReturnsCoordinateError(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{})
	for _, coordinates := range [][2]int{{0, 1}, {1, 0}} {
		_, err := GetStyle(buffer, "Sheet1", coordinates[0], coordinates[1])
		require.Error(t, err)
		assert.Contains(t, err.Error(), "convert cell coordinates")
	}
}

func TestGetStyleReturnsGetCellStyleError(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{})
	_, err := GetStyle(buffer, "Missing", 1, 1)
	require.Error(t, err)
	assert.Contains(t, err.Error(), "get cell style")
}

func TestGetStyleReturnsInvalidCellStyleIDError(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{Font: &excelize.Font{Bold: true}})
	buffer = rewriteXLSXEntry(t, buffer, "xl/worksheets/sheet1.xml", func(data []byte) []byte {
		old := []byte(`s="1"`)
		require.Contains(t, string(data), string(old))
		return bytes.Replace(data, old, []byte(`s="999"`), 1)
	})

	_, err := GetStyle(buffer, "Sheet1", 1, 1)
	require.Error(t, err)
}

func TestGetStylePreservesPatternFill(t *testing.T) {
	tests := []struct {
		name    string
		pattern int
	}{
		{name: "solid", pattern: 1},
		{name: "dark horizontal", pattern: 5},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			buffer := workbookWithStyle(t, &excelize.Style{
				Fill: excelize.Fill{
					Type:    "pattern",
					Pattern: test.pattern,
					Color:   []string{"#315D3C"},
				},
			})

			got := mustGetStyle(t, buffer, "Sheet1", 1, 1).Fill
			assert.Equal(t, "pattern", got.Type)
			assert.Equal(t, test.pattern, got.Pattern)
			assert.Equal(t, []string{"#315D3C"}, got.Color)
		})
	}
}

func TestGetStylePreservesFullStyle(t *testing.T) {
	tests := []struct {
		name  string
		style *excelize.Style
		check func(*testing.T, excelize.Style)
	}{
		{
			name: "built in number format and style fields",
			style: &excelize.Style{
				Font: &excelize.Font{
					Bold:      true,
					Italic:    true,
					Underline: "single",
					Family:    "Aptos",
					Size:      11,
					Strike:    true,
					Color:     "#123456",
				},
				Alignment: &excelize.Alignment{
					Horizontal:      "center",
					Indent:          2,
					JustifyLastLine: true,
					ReadingOrder:    1,
					RelativeIndent:  3,
					ShrinkToFit:     true,
					TextRotation:    45,
					Vertical:        "center",
					WrapText:        true,
				},
				Protection: &excelize.Protection{Hidden: true, Locked: false},
				NumFmt:     14,
			},
			check: func(t *testing.T, got excelize.Style) {
				require.NotNil(t, got.Font)
				wantFont := &excelize.Font{
					Bold: true, Italic: true, Underline: "single", Family: "Aptos",
					Size: 11, Strike: true, Color: "#123456",
				}
				setPublicField(wantFont, "ColorIndexed", -1)
				assert.Equal(t, wantFont, got.Font)
				assert.Equal(t, &excelize.Alignment{
					Horizontal: "center", Indent: 2, JustifyLastLine: true,
					ReadingOrder: 1, RelativeIndent: 3, ShrinkToFit: true,
					TextRotation: 45, Vertical: "center", WrapText: true,
				}, got.Alignment)
				assert.Equal(t, &excelize.Protection{Hidden: true, Locked: false}, got.Protection)
				assert.Equal(t, 14, got.NumFmt)
			},
		},
		{
			name: "custom number format",
			style: func() *excelize.Style {
				format := `yyyy-mm-dd "date"`
				return &excelize.Style{CustomNumFmt: &format}
			}(),
			check: func(t *testing.T, got excelize.Style) {
				require.NotNil(t, got.CustomNumFmt)
				assert.Equal(t, `yyyy-mm-dd "date"`, *got.CustomNumFmt)
			},
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			buffer := workbookWithStyle(t, test.style)
			got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
			test.check(t, got)
			reapplied := workbookWithStyle(t, &got)
			test.check(t, mustGetStyle(t, reapplied, "Sheet1", 1, 1))
		})
	}
}

func TestGetStyleInheritsNamedStyleXF(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{NumFmt: 14})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func([]byte) []byte {
		data := []byte(namedStyleStylesXML)
		oldChild := []byte(`<xf xfId="1"/>`)
		require.Equal(t, 1, bytes.Count(data, oldChild))
		return bytes.Replace(data, oldChild, []byte(`<xf xfId="1" applyAlignment="0" applyProtection="0"/>`), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
	require.Equal(t, 14, got.NumFmt)
	wantFont := &excelize.Font{
		Bold: true, Italic: true, Underline: "single", Family: "Aptos",
		Size: 14, Color: "#123456",
	}
	setPublicField(wantFont, "ColorIndexed", -1)
	require.NotNil(t, got.Font)
	assert.Equal(t, wantFont, got.Font)
	assert.Equal(t, excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"#315D3C"}}, got.Fill)
	assert.ElementsMatch(t, []excelize.Border{
		{Type: "left", Style: 1, Color: "#111111"},
		{Type: "right", Style: 2, Color: "#222222"},
	}, got.Border)
	assert.Equal(t, &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true}, got.Alignment)
	assert.Equal(t, &excelize.Protection{Locked: false, Hidden: true}, got.Protection)

	reapplied := workbookWithStyle(t, &got)
	reread := mustGetStyle(t, reapplied, "Sheet1", 1, 1)
	assert.Equal(t, got.Alignment, reread.Alignment)
	assert.Equal(t, got.Protection, reread.Protection)
	assert.Equal(t, styleAlignmentXML(t, reapplied), alignmentStyleXML{
		Horizontal: "center", Vertical: "center", WrapText: "true",
	})
	protection := styleProtectionXML(t, reapplied)
	assertXMLBool(t, protection.Locked, false)
	assertXMLBool(t, protection.Hidden, true)
}

func TestGetStyleInheritsCustomizedFontZeroWhenChildDoesNotApplyFont(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{NumFmt: 14})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func([]byte) []byte {
		data := []byte(namedStyleStylesXML)
		oldFont := []byte(`<font><sz val="11"/><name val="Calibri"/></font>`)
		newFont := []byte(`<font><b/><sz val="15"/><name val="InheritedZero"/></font>`)
		require.Equal(t, 1, bytes.Count(data, oldFont))
		data = bytes.Replace(data, oldFont, newFont, 1)
		oldParent := []byte(`numFmtId="14" fontId="1"`)
		require.Equal(t, 1, bytes.Count(data, oldParent))
		data = bytes.Replace(data, oldParent, []byte(`numFmtId="14" fontId="0"`), 1)
		oldChild := []byte(`<xf xfId="1"/>`)
		require.Equal(t, 1, bytes.Count(data, oldChild))
		return bytes.Replace(data, oldChild, []byte(`<xf xfId="1" applyFont="0"/>`), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
	require.NotNil(t, got.Font, "fontId=0 on the parent XF remains effective when the child explicitly inherits it")
	assert.True(t, got.Font.Bold)
	assert.Equal(t, float64(15), got.Font.Size)
	assert.Equal(t, "InheritedZero", got.Font.Family)

	reapplied := workbookWithStyle(t, &got)
	reread := mustGetStyle(t, reapplied, "Sheet1", 1, 1)
	require.NotNil(t, reread.Font)
	assert.True(t, reread.Font.Bold)
	assert.Equal(t, float64(15), reread.Font.Size)
	assert.Equal(t, "InheritedZero", reread.Font.Family)
	assert.Contains(t, xlsxEntry(t, reapplied, "xl/styles.xml"), `name val="InheritedZero"`)
}

func TestGetStyleAppliesImplicitFontZeroWhenApplyFontIsTrue(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{NumFmt: 14})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func([]byte) []byte {
		data := []byte(namedStyleStylesXML)
		oldFont := []byte(`<font><sz val="11"/><name val="Calibri"/></font>`)
		newFont := []byte(`<font><b/><sz val="15"/><name val="ImplicitZero"/></font>`)
		require.Equal(t, 1, bytes.Count(data, oldFont))
		data = bytes.Replace(data, oldFont, newFont, 1)
		oldChild := []byte(`<xf xfId="1"/>`)
		require.Equal(t, 1, bytes.Count(data, oldChild))
		return bytes.Replace(data, oldChild, []byte(`<xf xfId="1" applyFont="1"/>`), 1)
	})
	assert.Contains(t, xlsxEntry(t, buffer, "xl/styles.xml"), `<xf xfId="1" applyFont="1"/>`)

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
	require.NotNil(t, got.Font, "applyFont=true with omitted fontId applies schema-default fontId=0")
	assert.True(t, got.Font.Bold)
	assert.Equal(t, float64(15), got.Font.Size)
	assert.Equal(t, "ImplicitZero", got.Font.Family)

	reapplied := workbookWithStyle(t, &got)
	reread := mustGetStyle(t, reapplied, "Sheet1", 1, 1)
	require.NotNil(t, reread.Font)
	assert.True(t, reread.Font.Bold)
	assert.Equal(t, float64(15), reread.Font.Size)
	assert.Equal(t, "ImplicitZero", reread.Font.Family)
	assert.Contains(t, xlsxEntry(t, reapplied, "xl/styles.xml"), `name val="ImplicitZero"`)
}

func TestGetStylePreservesLocalizedBuiltInNumberFormatID(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{NumFmt: 14})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		old := []byte(`<xf numFmtId="14"`)
		require.Equal(t, 1, bytes.Count(data, old))
		return bytes.Replace(data, old, []byte(`<xf numFmtId="27"`), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
	assert.Equal(t, 27, got.NumFmt)
	assert.Nil(t, got.CustomNumFmt)
}

func TestGetStylePreservesLocalizedBuiltInNumberFormatCode(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{NumFmt: 14})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		code := `[$-ja-JP]yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;`
		return bytes.Replace(data, []byte("<fonts"), []byte(`<numFmts count="1"><numFmt numFmtId="14" formatCode="`+code+`"/></numFmts><fonts`), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
	require.NotNil(t, got.CustomNumFmt)
	assert.Equal(t, `[$-ja-JP]yyyy"年"m"月"d"日"`, *got.CustomNumFmt)
	assert.Zero(t, got.NumFmt)
}

func TestGetStylePrefersLocalizedNumberFormatCode16(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{NumFmt: 14})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		data = bytes.Replace(data, []byte("<styleSheet "), []byte(`<styleSheet xmlns:x16test="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" `), 1)
		local := `[$-ja-JP-x-gannen,80]ggge&quot;年&quot;m&quot;月&quot;d&quot;日&quot;`
		return bytes.Replace(data, []byte("<fonts"), []byte(`<numFmts count="1"><numFmt numFmtId="14" formatCode="yyyy-mm-dd" x16test:formatCode16="`+local+`"/></numFmts><fonts`), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
	require.NotNil(t, got.CustomNumFmt)
	assert.Equal(t, `[$-ja-JP-x-gannen,80]ggge"年"m"月"d"日"`, *got.CustomNumFmt)
	assert.Zero(t, got.NumFmt)
}

func TestGetStyleResolvesIndexedOOXMLColor(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"#123456"}},
	})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		old := []byte(`rgb="FF123456"`)
		require.Equal(t, 1, bytes.Count(data, old))
		data = bytes.Replace(data, old, []byte(`indexed="3"`), 1)
		colors := []byte(`<colors><indexedColors><rgbColor rgb="FF010101"/><rgbColor rgb="FF020202"/><rgbColor rgb="FF030303"/><rgbColor rgb="FFBEAD10"/></indexedColors></colors>`)
		return bytes.Replace(data, []byte("</styleSheet>"), append(colors, []byte("</styleSheet>")...), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1).Fill.Color
	assert.Equal(t, []string{"#BEAD10"}, got)
}

func TestGetStyleResolvesDefaultIndexedOOXMLColor(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"#123456"}},
	})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		old := []byte(`rgb="FF123456"`)
		require.Equal(t, 1, bytes.Count(data, old))
		return bytes.Replace(data, old, []byte(`indexed="5"`), 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1).Fill.Color
	assert.Equal(t, []string{"#FFFF00"}, got)
}

func TestGetStyleDefaultsOmittedUnderlineValueToSingle(t *testing.T) {
	buffer := workbookWithStyle(t, &excelize.Style{Font: &excelize.Font{Underline: "single"}})
	buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
		old := []byte(`val="single"`)
		require.Equal(t, 1, bytes.Count(data, old))
		return bytes.Replace(data, old, nil, 1)
	})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1).Font
	require.NotNil(t, got)
	assert.Equal(t, "single", got.Underline)
}

func TestGetStylePreservesAllBorderSides(t *testing.T) {
	want := []excelize.Border{
		{Type: "top", Style: 1, Color: "#111111"},
		{Type: "bottom", Style: 2, Color: "#222222"},
		{Type: "left", Style: 3, Color: "#333333"},
		{Type: "right", Style: 4, Color: "#444444"},
		{Type: "diagonalUp", Style: 6, Color: "#666666"},
		{Type: "diagonalDown", Style: 6, Color: "#666666"},
	}
	buffer := workbookWithStyle(t, &excelize.Style{Border: want})

	got := mustGetStyle(t, buffer, "Sheet1", 1, 1).Border
	assert.ElementsMatch(t, want, got)
	reapplied := workbookWithStyle(t, &excelize.Style{Border: got})
	assert.ElementsMatch(t, want, mustGetStyle(t, reapplied, "Sheet1", 1, 1).Border)
}

func TestGetStyleResolvesThemeColorChoiceKinds(t *testing.T) {
	tests := []struct {
		themeIndex int
		want       string
	}{
		{themeIndex: 0, want: "#AABBCC"},
		{themeIndex: 1, want: "#112233"},
		{themeIndex: 2, want: "#DDEEFF"},
		{themeIndex: 3, want: "#445566"},
		{themeIndex: 4, want: "#5B9BD5"},
		{themeIndex: 5, want: "#ED7D31"},
		{themeIndex: 6, want: "#A5A5A5"},
		{themeIndex: 7, want: "#FFC000"},
		{themeIndex: 8, want: "#4472C4"},
		{themeIndex: 9, want: "#70AD47"},
		{themeIndex: 10, want: "#0563C1"},
		{themeIndex: 11, want: "#954F72"},
	}

	for _, test := range tests {
		t.Run(fmt.Sprintf("theme-%d", test.themeIndex), func(t *testing.T) {
			buffer := workbookWithStyle(t, &excelize.Style{
				Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"#123456"}},
			})
			buffer = rewriteXLSXEntry(t, buffer, "xl/styles.xml", func(data []byte) []byte {
				old := []byte(`rgb="FF123456"`)
				newColor := []byte(fmt.Sprintf(`theme="%d"`, test.themeIndex))
				require.Equal(t, 1, bytes.Count(data, old))
				return bytes.Replace(data, old, newColor, 1)
			})
			buffer = rewriteXLSXEntry(t, buffer, "xl/theme/theme1.xml", func(data []byte) []byte {
				replacements := [][2]string{
					{"dk1", `<a:srgbClr val="112233"/>`},
					{"lt1", `<a:srgbClr val="AABBCC"/>`},
					{"dk2", `<a:sysClr val="windowText" lastClr="445566"/>`},
					{"lt2", `<a:sysClr val="window" lastClr="DDEEFF"/>`},
				}
				for _, replacement := range replacements {
					pattern := regexp.MustCompile(`(?s)<a:` + replacement[0] + `>.*?</a:` + replacement[0] + `>`)
					require.Equal(t, 1, len(pattern.FindAll(data, -1)))
					data = pattern.ReplaceAll(data, []byte(`<a:`+replacement[0]+`>`+replacement[1]+`</a:`+replacement[0]+`>`))
				}
				return data
			})

			got := mustGetStyle(t, buffer, "Sheet1", 1, 1)
			require.Len(t, got.Fill.Color, 1)
			assert.Equal(t, test.want, got.Fill.Color[0])
		})
	}
}

func TestGetStyleClosesLargeWorkbookTemporaryFiles(t *testing.T) {
	tempDir, err := os.MkdirTemp("", "excelizeutil-temp-")
	require.NoError(t, err)
	t.Cleanup(func() { require.NoError(t, os.RemoveAll(tempDir)) })
	t.Setenv("TMPDIR", tempDir)

	f := excelize.NewFile()
	sheet := f.GetSheetName(0)
	longText := strings.Repeat("x", 30_000)
	for row := 1; row <= 570; row++ {
		value := fmt.Sprintf("%04d%s", row, longText)
		axis := fmt.Sprintf("A%d", row)
		require.NoError(t, f.SetCellValue(sheet, axis, value))
	}
	buffer, err := f.WriteToBuffer()
	require.NoError(t, err)
	require.NoError(t, f.Close())

	_ = mustGetStyle(t, *buffer, sheet, 1, 1)

	entries, err := os.ReadDir(tempDir)
	require.NoError(t, err)
	assert.Empty(t, entries, "GetStyle should remove Excelize temporary files")
}

func rewriteXLSXEntry(t *testing.T, buffer bytes.Buffer, name string, rewrite func([]byte) []byte) bytes.Buffer {
	t.Helper()

	reader, err := zip.NewReader(bytes.NewReader(buffer.Bytes()), int64(buffer.Len()))
	require.NoError(t, err)

	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	found := false
	for _, file := range reader.File {
		contents, err := file.Open()
		require.NoError(t, err)
		data, err := io.ReadAll(contents)
		require.NoError(t, err)
		require.NoError(t, contents.Close())
		if file.Name == name {
			data = rewrite(data)
			found = true
		}

		entry, err := writer.Create(file.Name)
		require.NoError(t, err)
		_, err = entry.Write(data)
		require.NoError(t, err)
	}
	require.True(t, found, "XLSX archive did not contain %q", name)
	require.NoError(t, writer.Close())
	return output
}

const namedStyleStylesXML = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b val="1"/><i val="1"/><u val="single"/><sz val="14"/><color rgb="FF123456"/><name val="Aptos"/></font></fonts>
	<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF315D3C"/></patternFill></fill></fills>
	<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF111111"/></left><right style="medium"><color rgb="FF222222"/></right><top/><bottom/><diagonal/></border></borders>
	<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="1" fillId="2" borderId="1"><alignment horizontal="center" vertical="center" wrapText="1"/><protection locked="0" hidden="1"/></xf></cellStyleXfs>
	<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf xfId="1"/></cellXfs>
	<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Inherited" xfId="1" builtinId="1"/></cellStyles>
</styleSheet>`
