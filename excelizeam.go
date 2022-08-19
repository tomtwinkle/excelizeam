package excelizeam

import (
	"errors"
	"io"
	"sync"

	"golang.org/x/sync/errgroup"

	"github.com/mitchellh/hashstructure/v2"
	"github.com/tomtwinkle/excelizeam/excelizestyle"
	"github.com/xuri/excelize/v2"
)

var (
	ErrOverrideCellValue = errors.New("override cell value")
	ErrOverrideCellStyle = errors.New("override cell style")
)

type Excelizeam interface {
	// Preparations in advance

	// SetDefaultBorderStyle Set default cell border
	// For example, use when you want to paint the cell background white
	SetDefaultBorderStyle(style excelizestyle.BorderStyle, color excelizestyle.BorderColor) error

	// Excelize StreamWriter Wrapper

	SetPageMargins(options ...excelize.PageMarginsOptions) error
	SetColWidth(colIndex int, width float64) error
	SetColWidthRange(colIndexMin, colIndexMax int, width float64) error
	MergeCell(startColIndex, startRowIndex, endColIndex, endRowIndex int) error

	// SetCellValue Set value and style to cell
	SetCellValue(colIndex, rowIndex int, value interface{}, style *excelize.Style, override bool) error
	// SetCellValueAsync Set value and style to cell asynchronously
	SetCellValueAsync(colIndex, rowIndex int, value interface{}, style *excelize.Style)

	// SetStyleCell Set style to cell
	SetStyleCell(colIndex, rowIndex int, style excelize.Style, override bool) error
	// SetStyleCellAsync Set style to cell asynchronously
	SetStyleCellAsync(colIndex, rowIndex int, style excelize.Style)

	// SetStyleCellRange Set style to cell with range
	SetStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) error
	// SetStyleCellRangeAsync Set style to cell with range asynchronously
	SetStyleCellRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style)

	// SetBorderRange Set border around cell range
	SetBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) error
	// SetBorderRangeAsync Set border around cell range asynchronously
	SetBorderRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange)

	// Write StreamWriter
	Write(w io.Writer) error
}

type excelizeam struct {
	sw *excelize.StreamWriter

	eg errgroup.Group

	maxRow        int
	maxCol        int
	defaultBorder *DefaultBorders
	styleStore    sync.Map
	rowStore      sync.Map
}

type DefaultBorders struct {
	StyleID int

	Top    excelize.Border
	Bottom excelize.Border
	Left   excelize.Border
	Right  excelize.Border
}

type BorderItem struct {
	Style excelizestyle.BorderStyle
	Color excelizestyle.BorderColor
}

type BorderRange struct {
	Top    *BorderItem
	Bottom *BorderItem
	Left   *BorderItem
	Right  *BorderItem
	Inside *BorderItem
}

type StoredStyle struct {
	StyleID int
	Style   *excelize.Style
}

type StoredRow struct {
	Row
	Cols sync.Map
}

type Row struct {
	Index int
}

type Cell struct {
	StyleID int
	Value   interface{}
}

func New(sheetName string) (Excelizeam, error) {
	f := excelize.NewFile()
	f.SetSheetName("Sheet1", sheetName)
	sw, err := f.NewStreamWriter(sheetName)
	if err != nil {
		return nil, err
	}
	return &excelizeam{sw: sw}, nil
}

func (e *excelizeam) SetDefaultBorderStyle(style excelizestyle.BorderStyle, color excelizestyle.BorderColor) error {
	db := &DefaultBorders{
		Top:    excelizestyle.Border(excelizestyle.BorderPositionTop, style, color),
		Bottom: excelizestyle.Border(excelizestyle.BorderPositionBottom, style, color),
		Left:   excelizestyle.Border(excelizestyle.BorderPositionLeft, style, color),
		Right:  excelizestyle.Border(excelizestyle.BorderPositionRight, style, color),
	}
	styleID, err := e.getStyleID(&excelize.Style{
		Border: []excelize.Border{
			db.Top,
			db.Bottom,
			db.Left,
			db.Right,
		},
	})
	if err != nil {
		return err
	}
	db.StyleID = styleID
	e.defaultBorder = db
	return nil
}

func (e *excelizeam) SetColWidth(colIndex int, width float64) error {
	return e.sw.SetColWidth(colIndex, colIndex, width)
}

func (e *excelizeam) SetColWidthRange(colIndexMin, colIndexMax int, width float64) error {
	return e.sw.SetColWidth(colIndexMin, colIndexMax, width)
}

func (e *excelizeam) SetPageMargins(options ...excelize.PageMarginsOptions) error {
	return e.sw.File.SetPageMargins(
		e.sw.Sheet,
		options...,
	)
}

func (e *excelizeam) MergeCell(startColIndex, startRowIndex, endColIndex, endRowIndex int) error {
	startCell, err := excelize.CoordinatesToCellName(startColIndex, startRowIndex)
	if err != nil {
		return err
	}
	endCell, err := excelize.CoordinatesToCellName(endColIndex, endRowIndex)
	if err != nil {
		return err
	}
	return e.sw.MergeCell(startCell, endCell)
}

func (e *excelizeam) SetCellValueAsync(colIndex, rowIndex int, value interface{}, style *excelize.Style) {
	e.eg.Go(func() error {
		return e.setCellValue(colIndex, rowIndex, value, style, false)
	})
}

func (e *excelizeam) SetCellValue(colIndex, rowIndex int, value interface{}, style *excelize.Style, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setCellValue(colIndex, rowIndex, value, style, override)
}

func (e *excelizeam) setCellValue(colIndex, rowIndex int, value interface{}, style *excelize.Style, override bool) error {
	if e.maxCol < colIndex {
		e.maxCol = colIndex
	}
	if e.maxRow < rowIndex {
		e.maxRow = rowIndex
	}
	if _, ok := e.rowStore.Load(rowIndex); !ok {
		e.rowStore.Store(rowIndex, &StoredRow{
			Row: Row{
				Index: rowIndex,
			},
		})
	}
	if cacherow, ok := e.rowStore.Load(rowIndex); ok {
		r := cacherow.(*StoredRow)
		if cachec, ok := r.Cols.Load(colIndex); ok {
			c := cachec.(*Cell)
			if c.Value != nil && value != nil && !override {
				return ErrOverrideCellValue
			}
			if value != nil {
				c.Value = value
			}

			if style != nil {
				if c.StyleID > 0 {
					if !override {
						return ErrOverrideCellStyle
					}
					styleID, err := e.overrideStyle(c.StyleID, *style)
					if err != nil {
						return err
					}
					c.StyleID = styleID
				} else {
					styleID, err := e.getStyleID(style)
					if err != nil {
						return err
					}
					c.StyleID = styleID
				}
			}
			return nil
		}
		styleID, err := e.getStyleID(style)
		if err != nil {
			return err
		}
		r.Cols.Store(colIndex, &Cell{
			StyleID: styleID,
			Value:   value,
		})
	}
	return nil
}

func (e *excelizeam) SetStyleCellAsync(colIndex, rowIndex int, style excelize.Style) {
	e.eg.Go(func() error {
		return e.setStyleCell(colIndex, rowIndex, style, false)
	})
}

func (e *excelizeam) SetStyleCell(colIndex, rowIndex int, style excelize.Style, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setStyleCell(colIndex, rowIndex, style, override)
}

func (e *excelizeam) setStyleCell(colIndex, rowIndex int, style excelize.Style, override bool) error {
	if e.maxCol < colIndex {
		e.maxCol = colIndex
	}
	if e.maxRow < rowIndex {
		e.maxRow = rowIndex
	}

	if _, ok := e.rowStore.Load(rowIndex); !ok {
		e.rowStore.Store(rowIndex, &StoredRow{
			Row: Row{
				Index: rowIndex,
			},
		})
	}
	if cacherow, ok := e.rowStore.Load(rowIndex); ok {
		r := cacherow.(*StoredRow)
		if cachec, ok := r.Cols.Load(colIndex); ok {
			c := cachec.(*Cell)
			if c.StyleID > 0 {
				if !override {
					return ErrOverrideCellStyle
				}
				styleID, err := e.overrideStyle(c.StyleID, style)
				if err != nil {
					return err
				}
				c.StyleID = styleID
			} else {
				styleID, err := e.getStyleID(&style)
				if err != nil {
					return err
				}
				c.StyleID = styleID
			}
			return nil
		}
		styleID, err := e.getStyleID(&style)
		if err != nil {
			return err
		}
		r.Cols.Store(colIndex, &Cell{
			StyleID: styleID,
			Value:   nil,
		})
	}
	return nil
}

func (e *excelizeam) SetStyleCellRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style) {
	e.eg.Go(func() error {
		return e.setStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex, style, false)
	})
}

func (e *excelizeam) SetStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex, style, override)
}

func (e *excelizeam) setStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) error {
	if e.maxCol < endColIndex {
		e.maxCol = endColIndex
	}
	if e.maxRow < endRowIndex {
		e.maxRow = endRowIndex
	}

	for rowIdx := startRowIndex; rowIdx <= endRowIndex; rowIdx++ {
		for colIdx := startColIndex; colIdx <= endColIndex; colIdx++ {
			if _, ok := e.rowStore.Load(rowIdx); !ok {
				e.rowStore.Store(rowIdx, &StoredRow{
					Row: Row{
						Index: rowIdx,
					},
				})
			}
			if cacherow, ok := e.rowStore.Load(rowIdx); ok {
				r := cacherow.(*StoredRow)
				if cachec, ok := r.Cols.Load(colIdx); ok {
					c := cachec.(*Cell)
					if c.StyleID > 0 {
						if !override {
							return ErrOverrideCellStyle
						}
						styleID, err := e.overrideStyle(c.StyleID, style)
						if err != nil {
							return err
						}
						c.StyleID = styleID
					} else {
						styleID, err := e.getStyleID(&style)
						if err != nil {
							return err
						}
						c.StyleID = styleID
					}
					return nil
				}
				styleID, err := e.getStyleID(&style)
				if err != nil {
					return err
				}
				r.Cols.Store(colIdx, &Cell{
					StyleID: styleID,
					Value:   nil,
				})
			}
		}
	}
	return nil
}

func (e *excelizeam) SetBorderRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange) {
	e.eg.Go(func() error {
		return e.setBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex, borderRange, false)
	})
}

func (e *excelizeam) SetBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex, borderRange, override)
}

func (e *excelizeam) setBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) error {
	if e.maxCol < endColIndex {
		e.maxCol = endColIndex
	}
	if e.maxRow < endRowIndex {
		e.maxRow = endRowIndex
	}

	for rowIdx := startRowIndex; rowIdx <= endRowIndex; rowIdx++ {
		for colIdx := startColIndex; colIdx <= endColIndex; colIdx++ {
			borderStyles := make([]excelize.Border, 0, 4)
			switch {
			case rowIdx == startRowIndex && colIdx == startColIndex: // TopLeft
				if borderRange.Top != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Top.Style, borderRange.Top.Color),
					)
				}
				if borderRange.Left != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Left.Style, borderRange.Left.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx == startRowIndex && colIdx > startColIndex && colIdx < endColIndex: // TopMiddle
				if borderRange.Top != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Top.Style, borderRange.Top.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx == startRowIndex && colIdx == endColIndex: // TopRight
				if borderRange.Top != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Top.Style, borderRange.Top.Color),
					)
				}
				if borderRange.Right != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Right.Style, borderRange.Right.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx > startRowIndex && rowIdx < endRowIndex && colIdx == startColIndex: // MiddleLeft
				if borderRange.Left != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Left.Style, borderRange.Left.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx == endRowIndex && colIdx == startColIndex: // BottomLeft
				if borderRange.Bottom != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Bottom.Style, borderRange.Bottom.Color),
					)
				}
				if borderRange.Left != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Left.Style, borderRange.Left.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx == endRowIndex && colIdx > startColIndex && colIdx < endColIndex: // BottomMiddle
				if borderRange.Bottom != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Bottom.Style, borderRange.Bottom.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx == endRowIndex && colIdx == endColIndex: // BottomRight
				if borderRange.Bottom != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Bottom.Style, borderRange.Bottom.Color),
					)
				}
				if borderRange.Right != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Right.Style, borderRange.Right.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			case rowIdx > startRowIndex && rowIdx < endRowIndex && colIdx == endColIndex: // MiddleRight
				if borderRange.Right != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Right.Style, borderRange.Right.Color),
					)
				}
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			default: // InsideBorder
				if borderRange.Inside != nil {
					borderStyles = append(
						borderStyles,
						excelizestyle.Border(excelizestyle.BorderPositionTop, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionBottom, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionLeft, borderRange.Inside.Style, borderRange.Inside.Color),
						excelizestyle.Border(excelizestyle.BorderPositionRight, borderRange.Inside.Style, borderRange.Inside.Color),
					)
				}
			}

			if len(borderStyles) == 0 {
				continue
			}
			style := excelize.Style{Border: borderStyles}

			if _, ok := e.rowStore.Load(rowIdx); !ok {
				e.rowStore.Store(rowIdx, &StoredRow{
					Row: Row{
						Index: rowIdx,
					},
				})
			}
			if cacherow, ok := e.rowStore.Load(rowIdx); ok {
				r := cacherow.(*StoredRow)
				if cachec, ok := r.Cols.Load(colIdx); ok {
					c := cachec.(*Cell)
					if c.StyleID > 0 {
						if !override {
							return ErrOverrideCellStyle
						}
						styleID, err := e.overrideStyle(c.StyleID, style)
						if err != nil {
							return err
						}
						c.StyleID = styleID
					} else {
						styleID, err := e.getStyleID(&style)
						if err != nil {
							return err
						}
						c.StyleID = styleID
					}
					continue
				}
				styleID, err := e.getStyleID(&style)
				if err != nil {
					return err
				}
				r.Cols.Store(colIdx, &Cell{
					StyleID: styleID,
					Value:   nil,
				})
			}
		}
	}
	return nil
}

func (e *excelizeam) getStyleID(style *excelize.Style) (int, error) {
	var styl excelize.Style
	if style == nil {
		return 0, nil
	}
	if style != nil {
		styl = *style
	}
	hash, err := hashstructure.Hash(styl, hashstructure.FormatV2, nil)
	if err != nil {
		return 0, err
	}
	var styleID int
	if s, ok := e.styleStore.Load(hash); ok {
		styleID = s.(StoredStyle).StyleID
	} else {
		styleID, err = e.sw.File.NewStyle(&styl)
		if err != nil {
			return 0, err
		}
		e.styleStore.Store(hash, StoredStyle{
			StyleID: styleID,
			Style:   style,
		})
	}
	return styleID, nil
}

func (e *excelizeam) overrideStyle(originStyleID int, overrideStyle excelize.Style) (int, error) {
	var originStyle *excelize.Style
	e.styleStore.Range(func(_, value any) bool {
		if value.(StoredStyle).StyleID == originStyleID {
			originStyle = value.(StoredStyle).Style
			return false
		}
		return true
	})
	if originStyle == nil {
		return e.getStyleID(&overrideStyle)
	}

	style := new(excelize.Style)
	style.Fill = originStyle.Fill
	style.Alignment = originStyle.Alignment
	style.Font = originStyle.Font
	style.Lang = originStyle.Lang
	style.CustomNumFmt = originStyle.CustomNumFmt
	style.DecimalPlaces = originStyle.DecimalPlaces
	style.NegRed = originStyle.NegRed
	style.NumFmt = originStyle.NumFmt
	style.Protection = originStyle.Protection

	// Border
	borders := make([]excelize.Border, 0)
	if overrideBorder, ok := excelizestyle.FindBorder(overrideStyle.Border, excelizestyle.BorderPositionTop); ok {
		borders = append(borders, *overrideBorder)
	} else if originBorder, ok := excelizestyle.FindBorder(originStyle.Border, excelizestyle.BorderPositionTop); ok {
		borders = append(borders, *originBorder)
	}
	if overrideBorder, ok := excelizestyle.FindBorder(overrideStyle.Border, excelizestyle.BorderPositionBottom); ok {
		borders = append(borders, *overrideBorder)
	} else if originBorder, ok := excelizestyle.FindBorder(originStyle.Border, excelizestyle.BorderPositionBottom); ok {
		borders = append(borders, *originBorder)
	}
	if overrideBorder, ok := excelizestyle.FindBorder(overrideStyle.Border, excelizestyle.BorderPositionLeft); ok {
		borders = append(borders, *overrideBorder)
	} else if originBorder, ok := excelizestyle.FindBorder(originStyle.Border, excelizestyle.BorderPositionLeft); ok {
		borders = append(borders, *originBorder)
	}
	if overrideBorder, ok := excelizestyle.FindBorder(overrideStyle.Border, excelizestyle.BorderPositionRight); ok {
		borders = append(borders, *overrideBorder)
	} else if originBorder, ok := excelizestyle.FindBorder(originStyle.Border, excelizestyle.BorderPositionRight); ok {
		borders = append(borders, *originBorder)
	}
	style.Border = borders

	// Fill
	if overrideStyle.Fill.Type != "" {
		style.Fill = overrideStyle.Fill
	}

	// Alignment
	if overrideStyle.Alignment != nil {
		style.Alignment = overrideStyle.Alignment
	}

	// Font
	if overrideStyle.Font != nil {
		style.Font = overrideStyle.Font
	}

	// Lang
	if overrideStyle.Lang != "" {
		style.Lang = overrideStyle.Lang
	}

	// CustomNumFmt
	if overrideStyle.CustomNumFmt != nil {
		style.CustomNumFmt = overrideStyle.CustomNumFmt
	}

	// DecimalPlaces
	if overrideStyle.DecimalPlaces != 0 {
		style.DecimalPlaces = overrideStyle.DecimalPlaces
	}

	// NegRed
	if overrideStyle.NegRed {
		style.NegRed = overrideStyle.NegRed
	}

	// NumFmt
	if overrideStyle.NumFmt != 0 {
		style.NumFmt = overrideStyle.NumFmt
	}

	// Protection
	if overrideStyle.Protection != nil {
		style.Protection = overrideStyle.Protection
	}

	return e.getStyleID(style)
}

func (e *excelizeam) Write(w io.Writer) error {
	if err := e.writeStream(); err != nil {
		return err
	}
	if err := e.sw.Flush(); err != nil {
		return err
	}
	if err := e.sw.File.Write(w); err != nil {
		return err
	}
	return nil
}

func (e *excelizeam) writeStream() error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	defaultStyleCells := make([]interface{}, e.maxCol)
	if e.defaultBorder != nil {
		for i := 0; i < e.maxCol; i++ {
			defaultStyleCells[i] = excelize.Cell{StyleID: e.defaultBorder.StyleID, Value: ""}
		}
	}

	for rowIdx := 1; rowIdx <= e.maxRow; rowIdx++ {
		cacherow, rowOK := e.rowStore.Load(rowIdx)
		if !rowOK {
			// Value/Styleのない行はデフォルトStyleのみ設定
			if e.defaultBorder != nil {
				cell, err := excelize.CoordinatesToCellName(1, rowIdx)
				if err != nil {
					return err
				}
				if err := e.sw.SetRow(
					cell,
					defaultStyleCells,
				); err != nil {
					return err
				}
			}
			continue
		}

		r := cacherow.(*StoredRow)
		canWrite := false
		cellValues := make([]interface{}, e.maxCol)
		if e.defaultBorder != nil {
			canWrite = true
			copy(cellValues, defaultStyleCells)
		}
		r.Cols.Range(func(key, cachec any) bool {
			colIdx := key.(int)
			c := cachec.(*Cell)
			cellValues[colIdx-1] = excelize.Cell{StyleID: c.StyleID, Value: c.Value}
			canWrite = true
			return true
		})

		if canWrite {
			cell, err := excelize.CoordinatesToCellName(1, rowIdx)
			if err != nil {
				return err
			}
			if err := e.sw.SetRow(
				cell,
				cellValues,
			); err != nil {
				return err
			}
		}
	}
	return nil
}
