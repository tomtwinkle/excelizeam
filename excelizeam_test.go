package excelizeam_test

import (
	"bytes"
	"fmt"
	"io"
	"testing"

	"golang.org/x/sync/errgroup"

	"github.com/tomtwinkle/excelizeam"
	"github.com/tomtwinkle/excelizeam/excelizestyle"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

func TestExcelizeam_Sync(t *testing.T) {
	tests := map[string]struct {
		testFunc func(w excelizeam.Excelizeam) error
		wantErr  error
	}{
		"SetCellValue-with_not_style": {
			testFunc: func(w excelizeam.Excelizeam) error {
				return w.SetCellValue(1, 1, "test", nil, false, false)
			},
		},
		"SetCellValue-with_not_style_override": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(1, 1, "test1", nil, false, false); err != nil {
					return err
				}
				// can override value
				if err := w.SetCellValue(1, 1, "test2", nil, true, true); err != nil {
					return err
				}
				return nil
			},
		},
		"SetCellValue-with_not_style_override_style_error": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(1, 1, "test1", &excelize.Style{Font: &excelize.Font{Size: 12}}, false, false); err != nil {
					return err
				}
				// can override value
				if err := w.SetCellValue(1, 1, "test2", &excelize.Style{Font: &excelize.Font{Size: 13}}, true, false); err != nil {
					return err
				}
				return nil
			},
			wantErr: excelizeam.ErrOverrideCellStyle,
		},
		"SetCellValue-with_not_style_override_value_error": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(1, 1, "test1", &excelize.Style{Font: &excelize.Font{Size: 12}}, false, false); err != nil {
					return err
				}
				// can override value
				if err := w.SetCellValue(1, 1, "test2", &excelize.Style{Font: &excelize.Font{Size: 13}}, false, true); err != nil {
					return err
				}
				return nil
			},
			wantErr: excelizeam.ErrOverrideCellValue,
		},
		"SetCellValue-with_not_style_multiple_rows_cols_no_sort": {
			testFunc: func(w excelizeam.Excelizeam) error {
				for rowIdx := 1; rowIdx <= 10; rowIdx++ {
					for colIdx := 1; colIdx <= 10; colIdx++ {
						if err := w.SetCellValue(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), nil, false, false); err != nil {
							return err
						}
					}
				}
				return nil
			},
		},
		"SetCellValue-with_not_style_multiple_rows_cols_no_sort_odd": {
			testFunc: func(w excelizeam.Excelizeam) error {
				for rowIdx := 1; rowIdx <= 10; rowIdx++ {
					if rowIdx%2 == 0 {
						continue
					}
					for colIdx := 1; colIdx <= 10; colIdx++ {
						if colIdx%2 == 0 {
							continue
						}
						if err := w.SetCellValue(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), nil, false, false); err != nil {
							return err
						}
					}
				}
				return nil
			},
		},
		"SetCellValue-with_not_style_multiple_rows_cols_no_sort_odd_override_cross": {
			testFunc: func(w excelizeam.Excelizeam) error {
				for rowIdx := 1; rowIdx <= 7; rowIdx++ {
					if rowIdx%2 == 0 {
						continue
					}
					for colIdx := 1; colIdx <= 7; colIdx++ {
						if colIdx%2 == 0 {
							continue
						}
						if err := w.SetCellValue(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), &excelize.Style{
							Border: excelizestyle.BorderAround(excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
						}, false, false); err != nil {
							return err
						}
					}
				}
				for rowIdx := 5; rowIdx <= 10; rowIdx++ {
					if rowIdx%2 == 0 {
						continue
					}
					for colIdx := 5; colIdx <= 10; colIdx++ {
						if colIdx%2 == 0 {
							continue
						}
						if err := w.SetCellValue(colIdx, rowIdx, fmt.Sprintf("override-test%d-%d", rowIdx, colIdx), &excelize.Style{
							Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
						}, true, true); err != nil {
							return err
						}
					}
				}
				return nil
			},
		},
		"SetCellValue-with_not_style_multiple_rows_cols_sort": {
			testFunc: func(w excelizeam.Excelizeam) error {
				for colIdx := 1; colIdx <= 10; colIdx++ {
					for rowIdx := 1; rowIdx <= 10; rowIdx++ {
						if err := w.SetCellValue(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), nil, false, false); err != nil {
							return err
						}
					}
				}
				return nil
			},
		},
		"SetCellValue-with_style_border_fill_font_alignment": {
			testFunc: func(w excelizeam.Excelizeam) error {
				return w.SetCellValue(2, 2, "test", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
					Font: &excelize.Font{
						Bold:  true,
						Size:  8,
						Color: "#718DDC",
					},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				}, false, false)
			},
		},
		"SetCellValue-with_style_border_fill_font_alignment_override_border_top": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
					Font: &excelize.Font{
						Bold:  true,
						Size:  8,
						Color: "#718DDC",
					},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				}, false, false); err != nil {
					return err
				}
				if err := w.SetCellValue(2, 2, "test2", &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionTop, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, true, true); err != nil {
					return err
				}
				return nil
			},
		},
		"SetCellValue-with_style_border_fill_font_alignment_override_border_bottom": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
					Font: &excelize.Font{
						Bold:  true,
						Size:  8,
						Color: "#718DDC",
					},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				}, false, false); err != nil {
					return err
				}
				if err := w.SetCellValue(2, 2, "test2", &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionBottom, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, true, true); err != nil {
					return err
				}
				return nil
			},
		},
		"SetCellValue-with_style_border_fill_font_alignment_override_border_left": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
					Font: &excelize.Font{
						Bold:  true,
						Size:  8,
						Color: "#718DDC",
					},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				}, false, false); err != nil {
					return err
				}
				if err := w.SetCellValue(2, 2, "test2", &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionLeft, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, true, true); err != nil {
					return err
				}
				return nil
			},
		},
		"SetCellValue-with_style_border_fill_font_alignment_override_border_right": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
					Font: &excelize.Font{
						Bold:  true,
						Size:  8,
						Color: "#718DDC",
					},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				}, false, false); err != nil {
					return err
				}
				if err := w.SetCellValue(2, 2, "test2", &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionRight, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, true, true); err != nil {
					return err
				}
				return nil
			},
		},
		"SetCellValue-with_style_border_fill_font_alignment_override_value_error": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				}, false, false); err != nil {
					return err
				}
				if err := w.SetCellValue(2, 2, "", &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionRight, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, false, false); err != nil {
					return err
				}
				return nil
			},
			wantErr: excelizeam.ErrOverrideCellValue,
		},
		"SetCellValue-with_style_border_fill_font_alignment_override_style_error": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				}, false, false); err != nil {
					return err
				}
				if err := w.SetCellValue(2, 2, nil, &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionRight, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, false, false); err != nil {
					return err
				}
				return nil
			},
			wantErr: excelizeam.ErrOverrideCellStyle,
		},
		"SetBorderRange-border_odd_override_cross": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetBorderRange(1, 1, 7, 7, excelizeam.BorderRange{
					Top:    &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Bottom: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Left:   &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Right:  &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Inside: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
				}, false); err != nil {
					return err
				}
				if err := w.SetBorderRange(5, 5, 10, 10, excelizeam.BorderRange{
					Top:    &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Bottom: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Left:   &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Right:  &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Inside: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
				}, true); err != nil {
					return err
				}
				return nil
			},
		},
		"SetStyleCellRange-border_odd_override_cross": {
			testFunc: func(w excelizeam.Excelizeam) error {
				if err := w.SetStyleCellRange(1, 1, 7, 7, excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#BF00BF"),
				}, false); err != nil {
					return err
				}
				if err := w.SetStyleCellRange(5, 5, 10, 10, excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#CFA0FF"),
				}, true); err != nil {
					return err
				}
				if err := w.SetStyleCellRange(6, 6, 7, 7, excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#7FA04F"),
				}, true); err != nil {
					return err
				}
				return nil
			},
		},
	}

	for n, v := range tests {
		name := n
		tt := v
		t.Run(name, func(t *testing.T) {
			w, err := excelizeam.New("test")
			assert.NoError(t, err)
			err = tt.testFunc(w)
			if tt.wantErr != nil {
				assert.ErrorIs(t, err, tt.wantErr)
				return
			}
			assert.NoError(t, err)
			var buf bytes.Buffer
			err = w.Write(&buf)
			if !assert.NoError(t, err) {
				return
			}

			//f, err := os.Create("testdata/sync/" + name + ".xlsx")
			//assert.NoError(t, err)
			//_, err = f.Write(buf.Bytes())
			//assert.NoError(t, err)

			expected, err := excelize.OpenFile("testdata/sync/" + name + ".xlsx")
			if !assert.NoError(t, err) {
				return
			}
			actual, err := excelize.OpenReader(&buf)
			if !assert.NoError(t, err) {
				return
			}
			Assert(t, expected, actual)
		})
	}
}

func TestExcelizeam_Async(t *testing.T) {
	tests := map[string]struct {
		testFunc func(w excelizeam.Excelizeam)
		wantErr  error
	}{
		"SetCellValueAsync-with_not_style": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValueAsync(1, 1, "test", nil, false)
			},
		},
		"SetCellValueAsync-with_not_style_override_style": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValue(1, 1, "test1", &excelize.Style{Font: &excelize.Font{Size: 12}}, false, false)
				// can override value
				w.SetCellValueAsync(1, 1, nil, &excelize.Style{Font: &excelize.Font{Size: 13}}, true)
			},
		},
		"SetCellValueAsync-with_not_style_override_value_error": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValue(1, 1, "test1", &excelize.Style{Font: &excelize.Font{Size: 12}}, false, false)
				// can override value
				w.SetCellValueAsync(1, 1, "test2", &excelize.Style{Font: &excelize.Font{Size: 13}}, false)
			},
			wantErr: excelizeam.ErrOverrideCellValue,
		},
		"SetCellValueAsync-with_not_style_override_style_error": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValue(1, 1, "test1", &excelize.Style{Font: &excelize.Font{Size: 12}}, false, false)
				// can override value
				w.SetCellValueAsync(1, 1, nil, &excelize.Style{Font: &excelize.Font{Size: 13}}, false)
			},
			wantErr: excelizeam.ErrOverrideCellStyle,
		},
		"SetCellValueAsync-with_not_style_multiple_rows_cols_no_sort": {
			testFunc: func(w excelizeam.Excelizeam) {
				for rowIdx := 1; rowIdx <= 10; rowIdx++ {
					for colIdx := 1; colIdx <= 10; colIdx++ {
						w.SetCellValueAsync(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), nil, false)
					}
				}
			},
		},
		"SetCellValueAsync-with_not_style_multiple_rows_cols_no_sort_odd": {
			testFunc: func(w excelizeam.Excelizeam) {
				for rowIdx := 1; rowIdx <= 10; rowIdx++ {
					if rowIdx%2 == 0 {
						continue
					}
					for colIdx := 1; colIdx <= 10; colIdx++ {
						if colIdx%2 == 0 {
							continue
						}
						w.SetCellValueAsync(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), nil, false)
					}
				}
			},
		},
		"SetCellValueAsync-with_not_style_multiple_rows_cols_sort": {
			testFunc: func(w excelizeam.Excelizeam) {
				for colIdx := 1; colIdx <= 10; colIdx++ {
					for rowIdx := 1; rowIdx <= 10; rowIdx++ {
						w.SetCellValueAsync(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), nil, false)
					}
				}
			},
		},
		"SetCellValueAsync-with_style_border_fill_font_alignment": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValueAsync(2, 2, "test", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
					Font: &excelize.Font{
						Bold:  true,
						Size:  8,
						Color: "#718DDC",
					},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				}, false)
			},
		},
		"SetCellValueAsync-with_style_border_fill_font_alignment_odd_row": {
			testFunc: func(w excelizeam.Excelizeam) {
				for rowIdx := 1; rowIdx <= 10; rowIdx++ {
					if rowIdx%2 == 0 {
						continue
					}
					for colIdx := 1; colIdx <= 10; colIdx++ {
						if colIdx%2 == 0 {
							continue
						}
						w.SetCellValueAsync(colIdx, rowIdx, fmt.Sprintf("%d-%d", rowIdx, colIdx), &excelize.Style{
							Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
							Fill:   excelizestyle.Fill(excelizestyle.FillPatternSolid, "#315D3C"),
							Font: &excelize.Font{
								Bold:  true,
								Size:  8,
								Color: "#718DDC",
							},
							Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
						}, false)
					}
				}
			},
		},
		"SetCellValueAsync-with_style_border_fill_font_alignment_override_value_error": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				}, false, false)
				w.SetCellValueAsync(2, 2, "", &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionRight, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, false)
			},
			wantErr: excelizeam.ErrOverrideCellValue,
		},
		"SetCellValueAsync-with_style_border_fill_font_alignment_override_style_error": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetCellValue(2, 2, "test1", &excelize.Style{
					Border: excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				}, false, false)
				w.SetCellValueAsync(2, 2, nil, &excelize.Style{
					Border: []excelize.Border{
						excelizestyle.Border(excelizestyle.BorderPositionRight, excelizestyle.BorderStyleDash2, excelizestyle.BorderColorBlack),
					},
				}, false)
			},
			wantErr: excelizeam.ErrOverrideCellStyle,
		},
		"SetBorderRangeAsync-not_override_style": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetBorderRangeAsync(1, 1, 5, 5, excelizeam.BorderRange{
					Top:    &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Bottom: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Left:   &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Right:  &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Inside: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
				}, false)
				w.SetBorderRangeAsync(6, 6, 10, 10, excelizeam.BorderRange{
					Top:    &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Bottom: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Left:   &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Right:  &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Inside: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
				}, false)
			},
		},
		"SetBorderRangeAsync-override_error": {
			testFunc: func(w excelizeam.Excelizeam) {
				w.SetBorderRange(1, 1, 7, 7, excelizeam.BorderRange{
					Top:    &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Bottom: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Left:   &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Right:  &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
					Inside: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleDash2, Color: excelizestyle.BorderColorBlack},
				}, false)
				w.SetBorderRangeAsync(5, 5, 10, 10, excelizeam.BorderRange{
					Top:    &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Bottom: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Left:   &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Right:  &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
					Inside: &excelizeam.BorderItem{Style: excelizestyle.BorderStyleContinuous2, Color: excelizestyle.BorderColorBlack},
				}, false)
			},
			wantErr: excelizeam.ErrOverrideCellStyle,
		},
	}

	for n, v := range tests {
		name := n
		tt := v
		t.Run(name, func(t *testing.T) {
			w, err := excelizeam.New("test")
			assert.NoError(t, err)
			tt.testFunc(w)
			var buf bytes.Buffer
			err = w.Write(&buf)
			if tt.wantErr != nil {
				assert.ErrorIs(t, err, tt.wantErr)
				return
			}
			if !assert.NoError(t, err) {
				return
			}

			//f, err := os.Create("testdata/async/" + name + ".xlsx")
			//assert.NoError(t, err)
			//_, err = f.Write(buf.Bytes())
			//assert.NoError(t, err)

			expected, err := excelize.OpenFile("testdata/async/" + name + ".xlsx")
			if !assert.NoError(t, err) {
				return
			}
			actual, err := excelize.OpenReader(&buf)
			if !assert.NoError(t, err) {
				return
			}
			Assert(t, expected, actual)
		})
	}
}

func BenchmarkExcelizeam(b *testing.B) {
	b.Run("Excelize", func(b *testing.B) {
		var buf bytes.Buffer
		defer buf.Reset()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := benchExcelize(&buf); err != nil {
				b.Error(err)
			}
		}
	})
	b.Run("Excelize Async", func(b *testing.B) {
		var buf bytes.Buffer
		defer buf.Reset()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := benchExcelizeAsync(&buf); err != nil {
				b.Error(err)
			}
		}
	})
	b.Run("Excelize StreamWriter", func(b *testing.B) {
		var buf bytes.Buffer
		defer buf.Reset()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := benchStream(&buf); err != nil {
				b.Error(err)
			}
		}
	})
	b.Run("Excelizeam Sync", func(b *testing.B) {
		var buf bytes.Buffer
		defer buf.Reset()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := benchExcelizeam(&buf); err != nil {
				b.Error(err)
			}
		}
	})
	b.Run("Excelizeam Async", func(b *testing.B) {
		var buf bytes.Buffer
		defer buf.Reset()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := benchExcelizeamAsync(&buf); err != nil {
				b.Error(err)
			}
		}
	})
}

func benchExcelize(w io.Writer) error {
	f := excelize.NewFile()
	f.SetSheetName("Sheet1", "test")

	for rowIdx := 1; rowIdx <= 1000; rowIdx++ {
		for colIdx := 1; colIdx <= 10; colIdx++ {
			cell, err := excelize.CoordinatesToCellName(colIdx, rowIdx)
			if err != nil {
				return err
			}
			if err := f.SetCellValue("test", cell, fmt.Sprintf("test%d-%d", rowIdx, colIdx)); err != nil {
				return err
			}
			styleID, err := f.NewStyle(&excelize.Style{
				Border:    excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				Font:      &excelize.Font{Size: 12, Bold: true},
				Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
			})
			if err != nil {
				return err
			}
			if err := f.SetCellStyle("test", cell, cell, styleID); err != nil {
				return err
			}
		}
	}
	return f.Write(w)
}

func benchExcelizeAsync(w io.Writer) error {
	f := excelize.NewFile()
	f.SetSheetName("Sheet1", "test")

	var eg errgroup.Group

	for rowIdx := 1; rowIdx <= 1000; rowIdx++ {
		rowIdx := rowIdx
		for colIdx := 1; colIdx <= 10; colIdx++ {
			colIdx := colIdx
			eg.Go(func() error {
				cell, err := excelize.CoordinatesToCellName(colIdx, rowIdx)
				if err != nil {
					return err
				}
				if err := f.SetCellValue("test", cell, fmt.Sprintf("test%d-%d", rowIdx, colIdx)); err != nil {
					return err
				}
				styleID, err := f.NewStyle(&excelize.Style{
					Border:    excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
					Font:      &excelize.Font{Size: 12, Bold: true},
					Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
				})
				if err != nil {
					return err
				}
				if err := f.SetCellStyle("test", cell, cell, styleID); err != nil {
					return err
				}
				return nil
			})
		}
	}

	if err := eg.Wait(); err != nil {
		return err
	}
	return f.Write(w)
}

func benchStream(w io.Writer) error {
	f := excelize.NewFile()
	f.SetSheetName("Sheet1", "test")
	sw, err := f.NewStreamWriter("test")
	if err != nil {
		return err
	}

	for rowIdx := 1; rowIdx <= 1000; rowIdx++ {
		for colIdx := 1; colIdx <= 10; colIdx++ {
			cell, err := excelize.CoordinatesToCellName(colIdx, rowIdx)
			if err != nil {
				return err
			}
			styleID, err := f.NewStyle(&excelize.Style{
				Border:    excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				Font:      &excelize.Font{Size: 12, Bold: true},
				Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
			})
			if err != nil {
				return err
			}
			if err := sw.SetRow(cell, []interface{}{
				excelize.Cell{
					StyleID: styleID,
					Value:   fmt.Sprintf("test%d-%d", rowIdx, colIdx),
				},
			}); err != nil {
				return err
			}
		}
	}
	return f.Write(w)
}

func benchExcelizeam(w io.Writer) error {
	e, err := excelizeam.New("test")
	if err != nil {
		return err
	}

	for rowIdx := 1; rowIdx <= 1000; rowIdx++ {
		for colIdx := 1; colIdx <= 10; colIdx++ {
			if err := e.SetCellValue(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), &excelize.Style{
				Border:    excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				Font:      &excelize.Font{Size: 12, Bold: true},
				Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
			}, false, false); err != nil {
				return err
			}
		}
	}
	return e.Write(w)
}

func benchExcelizeamAsync(w io.Writer) error {
	e, err := excelizeam.New("test")
	if err != nil {
		return err
	}

	for rowIdx := 1; rowIdx <= 1000; rowIdx++ {
		for colIdx := 1; colIdx <= 10; colIdx++ {
			e.SetCellValueAsync(colIdx, rowIdx, fmt.Sprintf("test%d-%d", rowIdx, colIdx), &excelize.Style{
				Border:    excelizestyle.BorderAround(excelizestyle.BorderStyleContinuous2, excelizestyle.BorderColorBlack),
				Font:      &excelize.Font{Size: 12, Bold: true},
				Alignment: excelizestyle.Alignment(excelizestyle.AlignmentHorizontalCenter, excelizestyle.AlignmentVerticalCenter, true),
			}, false)
		}
	}

	return e.Write(w)
}

func Assert(t *testing.T, expected, actual *excelize.File) {
	for rowIdx := 1; rowIdx <= 10; rowIdx++ {
		for colIdx := 1; colIdx <= 10; colIdx++ {
			cell, err := excelize.CoordinatesToCellName(colIdx, rowIdx)
			if !assert.NoError(t, err) {
				return
			}

			// Assert Value
			expectedValue, err := expected.GetCellValue("test", cell)
			if !assert.NoError(t, err) {
				return
			}
			actualValue, err := actual.GetCellValue("test", cell)
			if !assert.NoError(t, err) {
				return
			}
			assert.Equal(t, expectedValue, actualValue)

			// Assert Style
			// TODO
		}
	}
}
