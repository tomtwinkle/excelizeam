package main

import (
	"fmt"
	"go/ast"
	"go/parser"
	"go/token"
	"sort"
	"strconv"
	"strings"
)

const officialStyleDocs = "https://xuri.me/excelize/en/style.html"

type enumDocumentation struct{ description, section string }

var enumDocs = map[string]enumDocumentation{
	"FillPattern":           {"selects the pattern index used by excelize.Fill.Pattern", "pattern"},
	"BorderStyle":           {"selects the line style index used by excelize.Border.Style", "border"},
	"BorderPosition":        {"selects the edge or diagonal used by excelize.Border.Type", "border"},
	"FillType":              {"selects pattern or gradient filling for excelize.Fill.Type; the empty value leaves the fill unset", "fill"},
	"AlignmentHorizontal":   {"selects horizontal cell alignment, including general alignment", "align"},
	"AlignmentVertical":     {"selects vertical cell alignment, including bottom alignment", "align"},
	"AlignmentReadingOrder": {"selects context-dependent, left-to-right, or right-to-left cell text direction", "align"},
	"FontUnderline":         {"selects the underline mode accepted by the selected Excelize dependency; the empty value leaves it unset", "underline"},
	"FontVerticalAlign":     {"selects baseline, superscript, or subscript placement of font text", ""},
	"NumberFormat":          {"selects an Excelize number-format ID; Code and CodeForLanguage resolve the dependency's format tables, not custom workbook definitions", "number_format"},
	"StyleLanguage":         {"identifies a language tag in Excelize's localized number-format tables; newer Style APIs use the resolved format code", "number_format"},
	"GradientShading":       {"selects a gradient preset in the selected Excelize dependency; preset numbering may change between versions", "shading"},
	"IndexedColor":          {"selects an indexed palette entry or system color for excelize.Font.ColorIndexed; workbook palettes can override the displayed RGB", ""},
	"ThemeColor":            {"selects a workbook theme slot for excelize.Font.ColorTheme; its displayed color depends on the workbook theme", ""},
	"FontCharset":           {"identifies an OOXML font character set for excelize.Font.Charset; these numeric selectors are not RGB colors or text encodings to apply to cell values", ""},
}

func enumDoc(name string) enumDocumentation {
	if doc, ok := enumDocs[name]; ok {
		return doc
	}
	if strings.HasSuffix(name, "Field") {
		owner := strings.TrimSuffix(name, "Field")
		section := ""
		switch owner {
		case "Border":
			section = "border"
		case "Fill":
			section = "fill"
		case "Alignment":
			section = "align"
		}
		return enumDocumentation{"identifies a public field of excelize." + owner + "; these values are field names rather than formatting values", section}
	}
	return enumDocumentation{"represents an Excelize style selector", ""}
}

func docLink(_ enumDocumentation) string {
	return officialStyleDocs
}

// documentEnums adds GoDoc to generated enum types and every named constant.
// Descriptions are produced at generation time, so regeneration preserves them.
func documentEnums(source []byte) ([]byte, error) {
	set := token.NewFileSet()
	file, err := parser.ParseFile(set, "generated.go", source, parser.ParseComments)
	if err != nil {
		return nil, err
	}
	formatCodes := map[string]string{}
	for _, declaration := range file.Decls {
		function, ok := declaration.(*ast.FuncDecl)
		if !ok || function.Name.Name != "Code" {
			continue
		}
		ast.Inspect(function.Body, func(node ast.Node) bool {
			clause, ok := node.(*ast.CaseClause)
			if !ok || len(clause.List) != 1 || len(clause.Body) != 1 {
				return true
			}
			name, ok := clause.List[0].(*ast.Ident)
			if !ok {
				return true
			}
			statement, ok := clause.Body[0].(*ast.ReturnStmt)
			if !ok || len(statement.Results) < 1 {
				return true
			}
			value, err := stringLiteral(statement.Results[0])
			if err == nil {
				formatCodes[name.Name] = value
			}
			return true
		})
	}
	typeNames := map[string]bool{}
	type insertion struct {
		offset int
		text   string
	}
	var edits []insertion
	add := func(pos token.Pos, name, description string, doc enumDocumentation) {
		edits = append(edits, insertion{set.Position(pos).Offset, fmt.Sprintf("// %s %s.\n//\n// See %s\n", name, description, docLink(doc))})
	}
	for _, decl := range file.Decls {
		group, ok := decl.(*ast.GenDecl)
		if !ok || group.Tok != token.TYPE {
			continue
		}
		for _, raw := range group.Specs {
			spec := raw.(*ast.TypeSpec)
			underlying, ok := spec.Type.(*ast.Ident)
			if !ok {
				continue
			}
			if underlying.Name != "int" && underlying.Name != "uint64" && underlying.Name != "string" {
				continue
			}
			name := spec.Name.Name
			typeNames[name] = true
			doc := enumDoc(name)
			if spec.Doc == nil && group.Doc == nil {
				add(group.Pos(), name, doc.description, doc)
			}
		}
	}
	for _, decl := range file.Decls {
		group, ok := decl.(*ast.GenDecl)
		if !ok || group.Tok != token.CONST {
			continue
		}
		currentType := ""
		for _, raw := range group.Specs {
			spec := raw.(*ast.ValueSpec)
			if id, ok := spec.Type.(*ast.Ident); ok {
				currentType = id.Name
			}
			if !typeNames[currentType] {
				continue
			}
			for _, id := range spec.Names {
				if spec.Doc != nil {
					continue
				}
				name := id.Name
				doc := enumDoc(currentType)
				description := enumValueDescription(currentType, strings.TrimPrefix(name, currentType))
				switch {
				case strings.HasSuffix(name, "Count"):
					description = "is the number of generated choices; it is not a valid formatting value"
				case strings.HasSuffix(currentType, "Field"):
					description = "identifies excelize." + strings.TrimSuffix(currentType, "Field") + "." + strings.TrimPrefix(name, currentType)
				case strings.HasPrefix(name, "NumberFormatID"):
					description = "selects Excelize number-format ID " + strings.TrimPrefix(name, "NumberFormatID")
					if code, ok := formatCodes[name]; ok {
						description += " with format code " + strconv.Quote(code)
					} else {
						description += "; its format depends on the language selected through CodeForLanguage"
					}
				case strings.HasPrefix(name, "GradientShadingPreset"):
					description = "selects gradient preset " + strings.TrimPrefix(name, "GradientShadingPreset") + " in the selected Excelize dependency"
				case name == "IndexedColorUnset":
					description = "leaves the indexed font color selector unspecified (ID -1)"
				case strings.HasPrefix(name, "IndexedColorIndex"):
					description = "selects palette index " + strings.TrimPrefix(name, "IndexedColorIndex") + "; the workbook palette determines its RGB color"
				case name == "IndexedColorSystemForeground":
					description = "selects the system foreground color (index 64)"
				case name == "IndexedColorSystemBackground":
					description = "selects the system background color (index 65)"
				case name == "AlignmentReadingOrderContext":
					description = "derives text direction from the first strong directional character"
				case name == "AlignmentReadingOrderLeftToRight":
					description = "reads cell text from left to right"
				case name == "AlignmentReadingOrderRightToLeft":
					description = "reads cell text from right to left"
				case len(spec.Values) == 1:
					if literal, ok := spec.Values[0].(*ast.BasicLit); ok {
						if literal.Kind == token.STRING {
							value, e := strconv.Unquote(literal.Value)
							if e != nil {
								return nil, e
							}
							if value == "" {
								description = "leaves the formatting selector unset"
							} else {
								description = stringValueDescription(currentType, value)
							}
						} else {
							description += " (ID " + literal.Value + ")"
						}
					}
				}
				add(spec.Pos(), name, description, doc)
			}
		}
	}
	sort.SliceStable(edits, func(i, j int) bool { return edits[i].offset > edits[j].offset })
	result := string(source)
	for _, edit := range edits {
		result = result[:edit.offset] + edit.text + result[edit.offset:]
	}
	return []byte(result), nil
}

func enumValueDescription(kind, value string) string {
	if kind == "BorderStyle" {
		names := map[string]string{"None": "no border", "Thin": "a thin solid border", "Medium": "a medium solid border", "Dashed": "a dashed border", "Dotted": "a dotted border", "Thick": "a thick solid border", "Double": "a double-line border", "Hair": "a hairline border", "MediumDashed": "a medium dashed border", "DashDot": "a dash-dot border", "MediumDashDot": "a medium dash-dot border", "DashDotDot": "a dash-dot-dot border", "MediumDashDotDot": "a medium dash-dot-dot border", "SlantDashDot": "a slanted dash-dot border"}
		if description, ok := names[value]; ok {
			return "selects " + description
		}
	}
	if kind == "FillPattern" {
		return "selects the " + strconv.Quote(strings.ToLower(value[:1])+value[1:]) + " cell-fill pattern"
	}
	if kind == "ThemeColor" {
		return "selects the workbook theme's " + value + " color slot"
	}
	if kind == "FontCharset" {
		return "selects the " + value + " font character set"
	}
	return "selects " + value
}

func stringValueDescription(kind, value string) string {
	meanings := map[string]map[string]string{
		"BorderPosition":      {"left": "selects the left cell edge", "right": "selects the right cell edge", "top": "selects the top cell edge", "bottom": "selects the bottom cell edge", "diagonalUp": "selects the diagonal from bottom-left to top-right", "diagonalDown": "selects the diagonal from top-left to bottom-right"},
		"FillType":            {"pattern": "uses the selected fill pattern and its foreground color", "gradient": "uses the selected gradient preset with two endpoint colors"},
		"AlignmentHorizontal": {"general": "aligns according to the cell value's data type", "left": "aligns text to the left edge", "center": "centers text horizontally", "right": "aligns text to the right edge", "fill": "repeats the cell text across the available width", "justify": "justifies text to both horizontal edges", "centerContinuous": "centers text across adjacent cells", "distributed": "distributes text evenly across the available width"},
		"AlignmentVertical":   {"top": "aligns text to the top edge", "center": "centers text vertically", "bottom": "aligns text to the bottom edge", "justify": "justifies text vertically", "distributed": "distributes text evenly across the available height"},
		"FontUnderline":       {"none": "explicitly disables underlining", "single": "uses a single underline", "double": "uses a double underline"},
		"FontVerticalAlign":   {"baseline": "uses normal baseline text", "superscript": "raises text above the baseline", "subscript": "lowers text below the baseline"},
	}
	if text, ok := meanings[kind][value]; ok {
		return text
	}
	if kind == "StyleLanguage" {
		return "identifies language tag " + strconv.Quote(value) + " for localized number-format lookup"
	}
	return "selects " + strconv.Quote(value) + " for " + kind
}
