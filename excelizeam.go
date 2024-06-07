package excelizeam

import (
	"crypto/sha1"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"
	"sync"

	"golang.org/x/sync/errgroup"

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

	SetPageMargins(options *excelize.PageLayoutMarginsOptions) error
	SetPageLayout(options *excelize.PageLayoutOptions) error
	GetPageLayout() (excelize.PageLayoutOptions, error)
	SetColWidth(colIndex int, width float64) error
	SetColWidthRange(colIndexMin, colIndexMax int, width float64) error
	MergeCell(startColIndex, startRowIndex, endColIndex, endRowIndex int) error

	// SetCellValue Set value and style to cell
	SetCellValue(colIndex, rowIndex int, value interface{}, style *excelize.Style, overrideValue, overrideStyle bool) error
	// SetCellValueAsync Set value and style to cell asynchronously
	SetCellValueAsync(colIndex, rowIndex int, value interface{}, style *excelize.Style, overrideStyle bool)

	// SetStyleCell Set style to cell
	SetStyleCell(colIndex, rowIndex int, style excelize.Style, override bool) error
	// SetStyleCellAsync Set style to cell asynchronously
	SetStyleCellAsync(colIndex, rowIndex int, style excelize.Style, override bool)

	// SetStyleCellRange Set style to cell with range
	SetStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) error
	// SetStyleCellRangeAsync Set style to cell with range asynchronously
	SetStyleCellRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool)

	// SetBorderRange Set border around cell range
	SetBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) error
	// SetBorderRangeAsync Set border around cell range asynchronously
	SetBorderRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool)

	// Wait
	// Wait for all running asynchronous operations to finish
	Wait() error

	// Write StreamWriter
	Write(w io.Writer) error

	// File Get the original excelize.File
	File() (*excelize.File, error)

	// CSVRecords Make csv records
	CSVRecords() ([][]string, error)
}

type excelizeam struct {
	sw   *excelize.StreamWriter
	file *excelize.File

	eg errgroup.Group

	mu     sync.Mutex
	maxRow int
	maxCol int

	defaultBorder *DefaultBorders
	styleStore    sync.Map
	cellStore     sync.Map
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

type Cell struct {
	StyleID int
	Value   interface{}
}

func New(sheetName string) (Excelizeam, error) {
	f := excelize.NewFile()
	err := f.SetSheetName("Sheet1", sheetName)
	if err != nil {
		return nil, err
	}
	sw, err := f.NewStreamWriter(sheetName)
	if err != nil {
		return nil, err
	}
	return &excelizeam{sw: sw, file: f}, nil
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

func (e *excelizeam) SetPageMargins(options *excelize.PageLayoutMarginsOptions) error {
	return e.file.SetPageMargins(
		e.sw.Sheet,
		options,
	)
}

func (e *excelizeam) SetPageLayout(options *excelize.PageLayoutOptions) error {
	return e.file.SetPageLayout(e.sw.Sheet, options)
}

func (e *excelizeam) GetPageLayout() (excelize.PageLayoutOptions, error) {
	return e.file.GetPageLayout(e.sw.Sheet)
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

func (e *excelizeam) SetCellValueAsync(colIndex, rowIndex int, value interface{}, style *excelize.Style, overrideStyle bool) {
	e.eg.Go(func() error {
		return e.setCellValue(colIndex, rowIndex, value, style, false, overrideStyle)
	})
}

func (e *excelizeam) SetCellValue(colIndex, rowIndex int, value interface{}, style *excelize.Style, overrideValue bool, overrideStyle bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setCellValue(colIndex, rowIndex, value, style, overrideValue, overrideStyle)
}

func (e *excelizeam) setCellValue(colIndex, rowIndex int, value interface{}, style *excelize.Style, overrideValue bool, overrideStyle bool) error {
	e.checkMaxIndex(colIndex, rowIndex)
	key := e.getCacheKey(colIndex, rowIndex)

	styleID, err := e.getStyleID(style)
	if err != nil {
		return err
	}
	if cached, ok := e.cellStore.LoadOrStore(key, &Cell{
		StyleID: styleID,
		Value:   value,
	}); ok {
		cell := cached.(*Cell)
		if cell.Value != nil && value != nil && !overrideValue {
			return ErrOverrideCellValue
		}
		if value != nil {
			cell.Value = value
		}

		if style != nil {
			if cell.StyleID > 0 {
				if !overrideStyle {
					return ErrOverrideCellStyle
				}
				styleID, err = e.overrideStyle(cell.StyleID, *style)
				if err != nil {
					return err
				}
			}
			cell.StyleID = styleID
		}
	}
	return nil
}

func (e *excelizeam) SetStyleCellAsync(colIndex, rowIndex int, style excelize.Style, override bool) {
	e.eg.Go(func() error {
		err := e.setStyleCell(colIndex, rowIndex, style, override)
		return err
	})
}

func (e *excelizeam) SetStyleCell(colIndex, rowIndex int, style excelize.Style, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setStyleCell(colIndex, rowIndex, style, override)
}

func (e *excelizeam) setStyleCell(colIndex, rowIndex int, style excelize.Style, override bool) error {
	e.checkMaxIndex(colIndex, rowIndex)
	key := e.getCacheKey(colIndex, rowIndex)

	styleID, err := e.getStyleID(&style)
	if err != nil {
		return err
	}
	if cached, ok := e.cellStore.LoadOrStore(key, &Cell{
		StyleID: styleID,
		Value:   nil,
	}); ok {
		c := cached.(*Cell)
		if c.StyleID > 0 {
			if !override {
				return ErrOverrideCellStyle
			}
			styleID, err = e.overrideStyle(c.StyleID, style)
			if err != nil {
				return err
			}
			c.StyleID = styleID
		}
		return nil
	}
	return nil
}

func (e *excelizeam) SetStyleCellRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) {
	e.eg.Go(func() error {
		err := e.setStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex, style, override)
		return err
	})
}

func (e *excelizeam) SetStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex, style, override)
}

func (e *excelizeam) setStyleCellRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, style excelize.Style, override bool) error {
	e.checkMaxIndex(endColIndex, endRowIndex)
	for rowIdx := startRowIndex; rowIdx <= endRowIndex; rowIdx++ {
		for colIdx := startColIndex; colIdx <= endColIndex; colIdx++ {
			key := e.getCacheKey(colIdx, rowIdx)

			styleID, err := e.getStyleID(&style)
			if err != nil {
				return err
			}
			if cached, ok := e.cellStore.LoadOrStore(key, &Cell{
				StyleID: styleID,
				Value:   nil,
			}); ok {
				c := cached.(*Cell)
				if c.StyleID > 0 {
					if !override {
						return ErrOverrideCellStyle
					}
					styleID, err = e.overrideStyle(c.StyleID, style)
					if err != nil {
						return err
					}
				}
				c.StyleID = styleID
			}
		}
	}
	return nil
}

func (e *excelizeam) SetBorderRangeAsync(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) {
	e.eg.Go(func() error {
		err := e.setBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex, borderRange, override)
		return err
	})
}

func (e *excelizeam) SetBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) error {
	if err := e.eg.Wait(); err != nil {
		return err
	}
	return e.setBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex, borderRange, override)
}

func (e *excelizeam) setBorderRange(startColIndex, startRowIndex, endColIndex, endRowIndex int, borderRange BorderRange, override bool) error {
	e.checkMaxIndex(endColIndex, endRowIndex)
	for rowIdx := startRowIndex; rowIdx <= endRowIndex; rowIdx++ {
		for colIdx := startColIndex; colIdx <= endColIndex; colIdx++ {
			key := e.getCacheKey(colIdx, rowIdx)
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

			styleID, err := e.getStyleID(&style)
			if err != nil {
				return err
			}
			if cached, ok := e.cellStore.LoadOrStore(key, &Cell{
				StyleID: styleID,
				Value:   nil,
			}); ok {
				c := cached.(*Cell)
				if c.StyleID > 0 {
					if !override {
						return ErrOverrideCellStyle
					}
					styleID, err = e.overrideStyle(c.StyleID, style)
					if err != nil {
						return err
					}
				}
				c.StyleID = styleID
				continue
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
	hash := fmt.Sprintf("%x", sha1.Sum([]byte(fmt.Sprintf("%+v", styl))))
	var styleID int
	if s, ok := e.styleStore.Load(hash); ok {
		styleID = s.(StoredStyle).StyleID
	} else {
		var err error
		styleID, err = e.file.NewStyle(&styl)
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

	// CustomNumFmt
	if overrideStyle.CustomNumFmt != nil {
		style.CustomNumFmt = overrideStyle.CustomNumFmt
	}

	// DecimalPlaces
	if overrideStyle.DecimalPlaces != nil {
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

func (e *excelizeam) getCacheKey(colIndex, rowIndex int) string {
	return fmt.Sprintf("%d-%d", rowIndex, colIndex)
}

func (e *excelizeam) getCacheAddress(key string) (colIndex, rowIndex int) {
	indexes := strings.Split(key, "-")
	if len(indexes) == 2 {
		var err error
		if rowIndex, err = strconv.Atoi(indexes[0]); err != nil {
			return 0, 0
		}
		if colIndex, err = strconv.Atoi(indexes[1]); err != nil {
			return 0, 0
		}
	}
	return colIndex, rowIndex
}

func (e *excelizeam) checkMaxIndex(colIndex, rowIndex int) {
	e.mu.Lock()
	defer e.mu.Unlock()
	if e.maxCol < colIndex {
		e.maxCol = colIndex
	}
	if e.maxRow < rowIndex {
		e.maxRow = rowIndex
	}
}

func (e *excelizeam) Wait() error {
	return e.eg.Wait()
}

func (e *excelizeam) Write(w io.Writer) error {
	if err := e.writeStream(); err != nil {
		return err
	}
	if err := e.sw.Flush(); err != nil {
		return err
	}
	if err := e.file.Write(w); err != nil {
		return err
	}
	return nil
}

func (e *excelizeam) File() (*excelize.File, error) {
	if err := e.writeStream(); err != nil {
		return nil, err
	}
	if err := e.sw.Flush(); err != nil {
		return nil, err
	}
	return e.file, nil
}

func (e *excelizeam) CSVRecords() ([][]string, error) {
	if err := e.eg.Wait(); err != nil {
		return nil, err
	}
	records := make([][]string, e.maxRow)
	for i := 0; i < e.maxRow; i++ {
		records[i] = make([]string, e.maxCol)
	}

	e.cellStore.Range(func(k, cached any) bool {
		key := k.(string)
		c := cached.(*Cell)
		colIdx, rowIdx := e.getCacheAddress(key)
		if c.Value != nil {
			records[rowIdx-1][colIdx-1] = fmt.Sprintf("%v", c.Value)
		}
		return true
	})
	return records, nil
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

	type writeCols struct {
		Cols     []interface{}
		CanWrite bool
	}

	writeRows := make([]writeCols, e.maxRow)
	for i := 0; i < e.maxRow; i++ {
		rowIdx := i + 1
		writeRows[i] = writeCols{
			Cols: make([]interface{}, e.maxCol),
		}
		if e.defaultBorder != nil {
			copy(writeRows[i].Cols, defaultStyleCells)
			writeRows[i].CanWrite = true
		}
		for ii := 0; ii < e.maxCol; ii++ {
			colIdx := ii + 1
			cached, ok := e.cellStore.Load(e.getCacheKey(colIdx, rowIdx))
			if !ok {
				continue
			}
			c := cached.(*Cell)
			writeRows[i].Cols[ii] = excelize.Cell{StyleID: c.StyleID, Value: c.Value}
			writeRows[i].CanWrite = true
		}
	}

	for i, row := range writeRows {
		if !row.CanWrite {
			continue
		}

		cell, err := excelize.CoordinatesToCellName(1, i+1)
		if err != nil {
			return err
		}
		if err := e.sw.SetRow(
			cell,
			row.Cols,
		); err != nil {
			return err
		}
	}
	return nil
}
