package main

import (
	"bytes"
	"go/parser"
	"go/token"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const fixtureStyles = `package excelize
var styleFillPatterns = []string{"none", "newPattern"}
var styleBorders = []string{"none", "newBorder"}
var styleBorderTypes = []string{"left", "diagonalUp"}
var builtInNumFmt = map[int]string{0:"general", 14:"mm-dd-yy"}
var currencyNumFmt = map[int]string{164:"currency-code"}
var langNumFmt = map[string]map[int]string{"ja-jp":{27:"localized-code"}}
var styleFillVariants = func() []xlsxGradientFill { return []xlsxGradientFill{
 {Degree:111, Stop:[]*xlsxGradientFillStop{{}, {Position:0.5}, {Position:1}}},
 {Type:"path", Left:0.5, Stop:[]*xlsxGradientFillStop{{}, {Position:1}}},
} }
// The following the type of horizontal alignment in cells:
//
// Style
// -----
// left
// newHorizontal
//
// The following the type of vertical alignment in cells:
//
// Style
// -----
// top
// newVertical
//
func NewStyle() {}
func newFills(style Style) {switch style.Fill.Type {case "gradient":case "pattern":}}
`
const fixtureTypes = `package excelize
type Border struct {Type string; Color string; Style int}
type Fill struct {Type string; Pattern int; Color []string; Shading int}
type Font struct {Bold bool; Underline string; NewFlag bool}
type Alignment struct {Horizontal string; Vertical string; ReadingOrder uint64}
type Protection struct {Locked bool; Hidden bool}
type Style struct {Border []Border; Fill Fill; Font *Font; Alignment *Alignment; Protection *Protection; NumFmt int; NewValue string}
`

func TestGenerationTracksNewFieldsValuesAndGradientPresets(t *testing.T) {
	dir := t.TempDir()
	for name, content := range map[string]string{"xmlStyles.go": fixtureTypes, "templates.go": `package excelize; var supportedUnderlineTypes=[]string{"none","single","double","newUnderline"}`} {
		if err := os.WriteFile(filepath.Join(dir, name), []byte(content), 0600); err != nil {
			t.Fatal(err)
		}
	}
	file, err := parser.ParseFile(token.NewFileSet(), "styles.go", fixtureStyles, parser.ParseComments)
	if err != nil {
		t.Fatal(err)
	}
	specs := []enumSpec{{typeName: "fillPattern", constantPrefix: "fillPattern", packageVar: "styleFillPatterns"}, {typeName: "borderStyle", constantPrefix: "borderStyle", packageVar: "styleBorders"}}
	public, geometry, err := additionalEnums(dir, file, specs)
	if err != nil {
		t.Fatal(err)
	}
	for _, fragment := range []string{"FillPatternNewPattern", "BorderStyleNewBorder", "AlignmentHorizontalNewHorizontal", "FontUnderlineNewUnderline", "FontFieldNewFlag", "StyleFieldNewValue", "out.NewValue = in.NewValue", "NumberFormatID164", "NumberFormatID27", "GradientShadingPreset1"} {
		if !strings.Contains(string(public), fragment) {
			t.Errorf("generated output does not track upstream addition %q", fragment)
		}
	}
	if !bytes.Contains(geometry, []byte("degree:111")) || !bytes.Contains(geometry, []byte("0,0.5,1")) {
		t.Fatalf("gradient geometry was not derived from source: %s", geometry)
	}
	if _, err := parser.ParseFile(token.NewFileSet(), "generated.go", public, 0); err != nil {
		t.Fatal(err)
	}
}

func TestUnknownGradientGeometryFailsGeneration(t *testing.T) {
	source := strings.Replace(fixtureStyles, "Degree:111", "NewGeometry:111", 1)
	file, err := parser.ParseFile(token.NewFileSet(), "styles.go", source, 0)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := gradientPresets(file); err == nil {
		t.Fatal("unknown gradient field was silently ignored")
	}
}

func TestCheckReportsStaleOutputWithoutWriting(t *testing.T) {
	old := *checkOutput
	*checkOutput = true
	defer func() { *checkOutput = old }()
	path := filepath.Join(t.TempDir(), "generated.go")
	original := []byte("package stale\n")
	if err := os.WriteFile(path, original, 0600); err != nil {
		t.Fatal(err)
	}
	if err := saveGenerated(path, []byte("package updated\n")); err == nil {
		t.Fatal("stale generated file accepted")
	}
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, original) {
		t.Fatal("check mode modified a file")
	}
}

func TestGeneratedFilesMatchSelectedDependency(t *testing.T) {
	dir, err := excelizeModuleDir()
	if err != nil {
		t.Fatal(err)
	}
	public, private, err := generatedSources(dir)
	if err != nil {
		t.Fatal(err)
	}
	for path, expected := range map[string][]byte{
		"../../../excelizestyle/style_enums_gen.go": public,
		"../../style_enums_gen.go":                  private,
	} {
		actual, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(actual, expected) {
			t.Errorf("%s is stale for the selected Excelize dependency; run go generate ./excelizeutil", path)
		}
	}
}
