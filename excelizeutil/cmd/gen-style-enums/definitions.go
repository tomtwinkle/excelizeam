package main

import (
	"bytes"
	"fmt"
	"go/ast"
	"go/parser"
	"go/token"
	"path/filepath"
	"strings"
)

var definitionTypes = map[string]bool{"Style": true, "Border": true, "Fill": true, "Font": true, "Alignment": true, "Protection": true}
var enumFields = map[string]string{
	"Border.Type": "BorderPosition", "Border.Style": "BorderStyle",
	"Fill.Type": "FillType", "Fill.Pattern": "FillPattern", "Fill.Shading": "GradientShading",
	"Font.Underline": "FontUnderline", "Font.VertAlign": "FontVerticalAlign",
	"Font.ColorIndexed": "IndexedColor", "Font.ColorTheme": "*ThemeColor", "Font.Charset": "*FontCharset",
	"Alignment.Horizontal": "AlignmentHorizontal", "Alignment.Vertical": "AlignmentVertical", "Alignment.ReadingOrder": "AlignmentReadingOrder",
	"Style.NumFmt": "NumberFormat", "Style.Lang": "StyleLanguage",
}

func typedDefinitions(out *bytes.Buffer, dir string) error {
	file, err := parser.ParseFile(token.NewFileSet(), filepath.Join(dir, "xmlStyles.go"), nil, 0)
	if err != nil {
		return err
	}
	for _, decl := range file.Decls {
		group, ok := decl.(*ast.GenDecl)
		if !ok {
			continue
		}
		for _, raw := range group.Specs {
			spec, ok := raw.(*ast.TypeSpec)
			if !ok || !definitionTypes[spec.Name.Name] {
				continue
			}
			structure, ok := spec.Type.(*ast.StructType)
			if !ok {
				return fmt.Errorf("%s is no longer a struct", spec.Name.Name)
			}
			name := spec.Name.Name
			fmt.Fprintf(out, "// %sDefinition preserves all public excelize.%s fields with typed categorical values.\ntype %sDefinition struct {\n", name, name, name)
			for _, field := range structure.Fields.List {
				for _, id := range field.Names {
					if !id.IsExported() {
						continue
					}
					kind := definitionFieldType(field.Type)
					if enum, ok := enumFields[name+"."+id.Name]; ok {
						kind = enum
					}
					fmt.Fprintf(out, "%s %s\n", id.Name, kind)
				}
			}
			fmt.Fprintln(out, "}")
			for _, reverse := range []bool{false, true} {
				if reverse {
					fmt.Fprintf(out, "func %sDefinitionFromExcelize(in excelize.%s) %sDefinition {\n out := %sDefinition{}\n", name, name, name, name)
				} else {
					fmt.Fprintf(out, "func (in %sDefinition) ToExcelize() excelize.%s {\n out := excelize.%s{}\n", name, name, name)
				}
				for _, field := range structure.Fields.List {
					for _, id := range field.Names {
						if id.IsExported() {
							writeConversion(out, name, id.Name, field.Type, reverse)
						}
					}
				}
				fmt.Fprintln(out, "return out\n}")
			}
		}
	}
	return nil
}

func definitionFieldType(expr ast.Expr) string {
	switch t := expr.(type) {
	case *ast.Ident:
		if definitionTypes[t.Name] {
			return t.Name + "Definition"
		}
		switch t.Name {
		case "bool", "int", "uint64", "float64", "string":
			return t.Name
		default:
			return "excelize." + t.Name
		}
	case *ast.StarExpr:
		return "*" + definitionFieldType(t.X)
	case *ast.ArrayType:
		if t.Len == nil {
			return "[]" + definitionFieldType(t.Elt)
		}
	}
	return expressionText(expr)
}

func writeConversion(out *bytes.Buffer, owner, field string, expr ast.Expr, reverse bool) {
	if enum, ok := enumFields[owner+"."+field]; ok {
		if pointer, ok := expr.(*ast.StarExpr); ok {
			cast := expressionText(pointer.X)
			if reverse {
				cast = strings.TrimPrefix(enum, "*")
			}
			fmt.Fprintf(out, "if in.%s != nil { value := %s(*in.%s); out.%s = &value }\n", field, cast, field, field)
			return
		}
		cast := expressionText(expr)
		if reverse {
			cast = enum
		}
		fmt.Fprintf(out, "out.%s = %s(in.%s)\n", field, cast, field)
		return
	}
	if id, ok := expr.(*ast.Ident); ok && definitionTypes[id.Name] {
		if reverse {
			fmt.Fprintf(out, "out.%s = %sDefinitionFromExcelize(in.%s)\n", field, id.Name, field)
		} else {
			fmt.Fprintf(out, "out.%s = in.%s.ToExcelize()\n", field, field)
		}
		return
	}
	if ptr, ok := expr.(*ast.StarExpr); ok {
		if id, ok := ptr.X.(*ast.Ident); ok && definitionTypes[id.Name] {
			if reverse {
				fmt.Fprintf(out, "if in.%s != nil { value := %sDefinitionFromExcelize(*in.%s); out.%s = &value }\n", field, id.Name, field, field)
			} else {
				fmt.Fprintf(out, "if in.%s != nil { value := in.%s.ToExcelize(); out.%s = &value }\n", field, field, field)
			}
			return
		}
	}
	if slice, ok := expr.(*ast.ArrayType); ok && slice.Len == nil {
		if id, ok := slice.Elt.(*ast.Ident); ok && definitionTypes[id.Name] {
			kind := "excelize." + id.Name
			if reverse {
				kind = id.Name + "Definition"
			}
			fmt.Fprintf(out, "if in.%s != nil { out.%s = make([]%s,len(in.%s)); for i, value := range in.%s {", field, field, kind, field, field)
			if reverse {
				fmt.Fprintf(out, "out.%s[i] = %sDefinitionFromExcelize(value)", field, id.Name)
			} else {
				fmt.Fprintf(out, "out.%s[i] = value.ToExcelize()", field)
			}
			fmt.Fprintln(out, "} }")
			return
		}
	}
	fmt.Fprintf(out, "out.%s = in.%s\n", field, field)
}
