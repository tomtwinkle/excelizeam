package excelizeutil

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
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

			got := GetStyle(buffer, "Sheet1", 1, 1).Fill
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
				assert.Equal(t, &excelize.Font{
					Bold: true, Italic: true, Underline: "single", Family: "Aptos",
					Size: 11, Strike: true, Color: "#123456",
				}, got.Font)
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
			test.check(t, GetStyle(buffer, "Sheet1", 1, 1))
		})
	}
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

	got := GetStyle(buffer, "Sheet1", 1, 1).Border
	assert.ElementsMatch(t, want, got)
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
					{`<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>`, `<a:dk1><a:srgbClr val="112233"/></a:dk1>`},
					{`<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>`, `<a:lt1><a:srgbClr val="AABBCC"/></a:lt1>`},
					{`<a:dk2><a:srgbClr val="44546A"/></a:dk2>`, `<a:dk2><a:sysClr val="windowText" lastClr="445566"/></a:dk2>`},
					{`<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>`, `<a:lt2><a:sysClr val="window" lastClr="DDEEFF"/></a:lt2>`},
				}
				for _, replacement := range replacements {
					old := []byte(replacement[0])
					if bytes.Count(data, old) == 0 {
						for _, colorElement := range []string{"sysClr", "srgbClr"} {
							expanded := []byte(strings.Replace(replacement[0], "/>", "></a:"+colorElement+">", 1))
							if bytes.Count(data, expanded) == 1 {
								old = expanded
								break
							}
						}
					}
					require.Equal(t, 1, bytes.Count(data, old))
					data = bytes.Replace(data, old, []byte(replacement[1]), 1)
				}
				return data
			})

			var got excelize.Style
			if !assert.NotPanics(t, func() { got = GetStyle(buffer, "Sheet1", 1, 1) }) {
				return
			}
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

	_ = GetStyle(*buffer, sheet, 1, 1)

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
