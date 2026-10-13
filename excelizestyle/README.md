# Typed Excelize styles

`StyleDefinition` and its nested definitions represent every public field of the
selected Excelize `Style`, `Border`, `Fill`, `Font`, `Alignment`, and `Protection`
structs. Categorical fields use named enum types. Numeric, text, boolean, and
pointer values retain their original meaning and can express custom values.

```go
style := excelizestyle.StyleDefinition{
    Fill: excelizestyle.FillDefinition{
        Type:    excelizestyle.FillTypePattern,
        Pattern: excelizestyle.FillPatternSolid,
        Color:   []string{"#123456"},
    },
    Alignment: &excelizestyle.AlignmentDefinition{
        Horizontal:   excelizestyle.AlignmentHorizontalCenter,
        Vertical:     excelizestyle.AlignmentVerticalBottom,
        ReadingOrder: excelizestyle.AlignmentReadingOrderLeftToRight,
    },
    NumFmt: excelizestyle.NumberFormatID14,
}
styleID, err := workbook.NewStyle(style.ToExcelize())
```

Use `StyleDefinitionFromExcelize` to convert existing styles, or
`excelizeutil.GetStyleDefinition` to read a cell into this representation.
Conversion between a definition and an `excelize.Style` preserves all struct
fields, including nil pointers, custom strings, unknown numeric enum values,
and the distinction between nil and empty border slices. This does not expand
what Excelize can serialize or recover from a workbook. In particular, cell fill
transparency, arbitrary gradient shapes, both pattern colors, and the original
`Lang`/`DecimalPlaces`/`NegRed` inputs cannot always be recovered.

## Generation and dependency updates

After updating Excelize, run from the repository root:

```sh
go generate ./excelizeutil
```

The generator reads the selected dependency through Go AST:

- Fill pattern, border style and position, and fill type tables.
- Gradient preset IDs, geometry, and stop positions (legacy switches and modern
  preset literals). The extractor uses these generated geometries.
- Documented alignment choices, underline and font vertical-alignment choices.
- Built-in, currency, and language-specific number format IDs/codes and language
  tags. `NumberFormat.CodeForLanguage` resolves the corresponding source table;
  these codes describe Excelize inputs, not arbitrary workbook custom formats
  that may reuse the same ID.
- Every public field name and the complete typed definitions/conversions.

Reading order, theme slots, character sets, and alignment defaults missing from
older documentation use stable OOXML definitions. Indexed palette values are
read from Excelize's palette when available; older releases use the 64-slot
OOXML palette. These schema definitions are explicit: AST cannot discover value
domains that Excelize does not define in code. Arbitrary RGB, font family/size,
indentation, rotation, tint, custom format, and transparency values remain plain
values rather than exhaustive enum lists.

The test suite also compares the generated files with the selected dependency,
so adding an upstream enum value or field without regenerating fails tests.

Check generation without changing files:

```sh
(cd excelizeutil && go run ./cmd/gen-style-enums -check)
```

Changes to unsupported source shapes cause generation to fail. Regeneration
updates definitions and enum tables; newly added fields still require a review
of the workbook extractor and its coverage. Tests check field inventories,
conversion of every field, upstream additions, source-shape failures, and
round-trips for all generated patterns, borders, and gradient presets.

The existing `Border`, `BorderAround`, `FindBorder`, `Alignment`, and `Fill`
helpers remain available, along with the original `BorderStyleContinuous*`,
`BorderStyleDash*`, and related aliases.
