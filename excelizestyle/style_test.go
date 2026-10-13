package excelizestyle

import (
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestStyleDefinitionPreservesEveryField(t *testing.T) {
	var original excelize.Style
	populate(reflect.ValueOf(&original).Elem())
	converted := StyleDefinitionFromExcelize(original).ToExcelize()
	if !reflect.DeepEqual(original, converted) {
		t.Fatalf("typed conversion lost a field:\noriginal: %#v\nconverted: %#v", original, converted)
	}
	// Nil and empty slices are observably different and must survive conversion.
	empty := excelize.Style{Border: []excelize.Border{}}
	if got := StyleDefinitionFromExcelize(empty).ToExcelize(); !reflect.DeepEqual(empty, got) {
		t.Fatal("empty Border changed to nil")
	}
	if got := StyleDefinitionFromExcelize(excelize.Style{}).ToExcelize(); !reflect.DeepEqual(excelize.Style{}, got) {
		t.Fatal("zero Style changed")
	}
}

func populate(v reflect.Value) {
	switch v.Kind() {
	case reflect.Struct:
		for i := 0; i < v.NumField(); i++ {
			if v.Field(i).CanSet() {
				populate(v.Field(i))
			}
		}
	case reflect.Ptr:
		v.Set(reflect.New(v.Type().Elem()))
		populate(v.Elem())
	case reflect.Slice:
		v.Set(reflect.MakeSlice(v.Type(), 2, 2))
		for i := 0; i < v.Len(); i++ {
			populate(v.Index(i))
		}
	case reflect.Bool:
		v.SetBool(true)
	case reflect.Int, reflect.Int64:
		v.SetInt(7)
	case reflect.Uint64:
		v.SetUint(2)
	case reflect.Float64:
		v.SetFloat(0.25)
	case reflect.String:
		v.SetString("custom-value")
	}
}

func TestGeneratedFieldsMatchExcelize(t *testing.T) {
	inventories := []struct {
		value  interface{}
		fields []string
	}{
		{excelize.Style{}, stringsOf(StyleFieldValues())},
		{excelize.Border{}, stringsOf(BorderFieldValues())},
		{excelize.Fill{}, stringsOf(FillFieldValues())},
		{excelize.Font{}, stringsOf(FontFieldValues())},
		{excelize.Alignment{}, stringsOf(AlignmentFieldValues())},
		{excelize.Protection{}, stringsOf(ProtectionFieldValues())},
	}
	for _, inventory := range inventories {
		typ := reflect.TypeOf(inventory.value)
		var expected []string
		for i := 0; i < typ.NumField(); i++ {
			if typ.Field(i).PkgPath == "" {
				expected = append(expected, typ.Field(i).Name)
			}
		}
		if !reflect.DeepEqual(expected, inventory.fields) {
			t.Errorf("%s generated fields are stale; run go generate ./excelizeutil: want %v, got %v", typ.Name(), expected, inventory.fields)
		}
	}
}

func stringsOf[T ~string](values []T) []string {
	result := make([]string, len(values))
	for i, v := range values {
		result[i] = string(v)
	}
	return result
}

func TestGeneratedEnumsAndExistingAliases(t *testing.T) {
	if BorderStyleContinuous0 != BorderStyleHair || BorderStyleDash2 != BorderStyleMediumDashed || BorderStyleSlantDash != BorderStyleSlantDashDot {
		t.Fatal("existing border aliases changed")
	}
	if AlignmentVerticalBottom != "bottom" || BorderPositionDiagonalUp != "diagonalUp" || BorderPositionDiagonalDown != "diagonalDown" {
		t.Fatal("missing categorical values")
	}
	if code, ok := NumberFormatID14.Code(); !ok || code == "" {
		t.Fatal("missing builtin number format")
	}
	if _, ok := NumberFormat(999999).Code(); ok {
		t.Fatal("unknown number format accepted")
	}
	for _, pattern := range []FillPattern{FillPatternNone, FillPatternSolid, FillPatternGray0625} {
		got, ok := ParseFillPattern(pattern.String())
		if !ok || got != pattern {
			t.Fatal("fill enum conversion changed")
		}
	}
}
