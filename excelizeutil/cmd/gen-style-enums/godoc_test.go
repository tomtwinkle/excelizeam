package main

import (
	"go/ast"
	"go/parser"
	"go/token"
	"strings"
	"testing"
)

func TestEnumDocumentationCoversTypesAndConstants(t *testing.T) {
	dir, err := excelizeModuleDir()
	if err != nil {
		t.Fatal(err)
	}
	public, _, err := generatedSources(dir)
	if err != nil {
		t.Fatal(err)
	}
	file, err := parser.ParseFile(token.NewFileSet(), "generated.go", public, parser.ParseComments)
	if err != nil {
		t.Fatal(err)
	}
	enums := map[string]bool{}
	check := func(name string, doc *ast.CommentGroup) {
		t.Helper()
		if doc == nil {
			t.Errorf("%s has no GoDoc", name)
			return
		}
		text := doc.Text()
		if !strings.HasPrefix(text, name+" ") {
			t.Errorf("%s GoDoc must start with its name: %s", name, text)
		}
		if !strings.Contains(text, officialStyleDocs) {
			t.Errorf("%s has no official documentation reference", name)
		}
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
			if underlying.Name != "int" && underlying.Name != "string" && underlying.Name != "uint64" {
				continue
			}
			enums[spec.Name.Name] = true
			doc := spec.Doc
			if doc == nil {
				doc = group.Doc
			}
			check(spec.Name.Name, doc)
		}
	}
	count := 0
	for _, decl := range file.Decls {
		group, ok := decl.(*ast.GenDecl)
		if !ok || group.Tok != token.CONST {
			continue
		}
		kind := ""
		for _, raw := range group.Specs {
			spec := raw.(*ast.ValueSpec)
			if id, ok := spec.Type.(*ast.Ident); ok {
				kind = id.Name
			}
			if !enums[kind] {
				continue
			}
			for _, name := range spec.Names {
				check(name.Name, spec.Doc)
				count++
			}
		}
	}
	if len(enums) < 10 || count < 100 {
		t.Fatal("enum documentation inventory is unexpectedly small")
	}
	for _, fragment := range []string{"selects a thin solid border", "with format code", "path bounds", "identifies excelize.Font.Bold"} {
		if !strings.Contains(string(public), fragment) {
			t.Errorf("missing meaningful explanation %q", fragment)
		}
	}
}
