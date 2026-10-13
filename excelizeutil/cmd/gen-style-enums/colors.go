package main

import (
	"bytes"
	"fmt"
	"go/ast"
)

// These identifiers are defined by OOXML, not by a closed enum in Excelize.
// Arbitrary underlying values remain expressible for future/custom definitions.
func colorEnums(out *bytes.Buffer, templates *ast.File) error {
	paletteSize := 64
	if expr := initializer(templates, "IndexedColorMapping"); expr != nil {
		colors, err := stringValues(expr, "IndexedColorMapping")
		if err != nil {
			return err
		}
		paletteSize = len(colors)
	}
	fmt.Fprintln(out, "type IndexedColor int\nconst (\nIndexedColorUnset IndexedColor = -1")
	for i := 0; i < paletteSize; i++ {
		fmt.Fprintf(out, "IndexedColorIndex%d IndexedColor = %d\n", i, i)
	}
	fmt.Fprintf(out, "IndexedColorSystemForeground IndexedColor = %d\nIndexedColorSystemBackground IndexedColor = %d\n)\n", 64, 65)
	fmt.Fprintln(out, "type ThemeColor int\nconst (")
	for i, name := range []string{"Light1", "Dark1", "Light2", "Dark2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"} {
		fmt.Fprintf(out, "ThemeColor%s ThemeColor = %d\n", name, i)
	}
	fmt.Fprintln(out, ")\ntype FontCharset int\nconst (")
	for _, v := range []struct {
		name string
		id   int
	}{
		{"ANSI", 0}, {"Default", 1}, {"Symbol", 2}, {"Mac", 77}, {"ShiftJIS", 128}, {"Hangul", 129}, {"Johab", 130}, {"GB2312", 134}, {"Big5", 136}, {"Greek", 161}, {"Turkish", 162}, {"Vietnamese", 163}, {"Hebrew", 177}, {"Arabic", 178}, {"Baltic", 186}, {"Russian", 204}, {"Thai", 222}, {"EastEurope", 238}, {"OEM", 255},
	} {
		fmt.Fprintf(out, "FontCharset%s FontCharset = %d\n", v.name, v.id)
	}
	fmt.Fprintln(out, ")")
	return nil
}
