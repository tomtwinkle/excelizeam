package excelizestyle

import (
	"github.com/xuri/excelize/v2"
)

// Preserve the existing border style names as aliases of the generated enums.
const (
	BorderStyleContinuous0 = BorderStyleHair
	BorderStyleContinuous1 = BorderStyleThin
	BorderStyleContinuous2 = BorderStyleMedium
	BorderStyleContinuous3 = BorderStyleThick
	BorderStyleDash1       = BorderStyleDashed
	BorderStyleDash2       = BorderStyleMediumDashed
	BorderStyleDot         = BorderStyleDotted
	BorderStyleDashDot1    = BorderStyleDashDot
	BorderStyleDashDot2    = BorderStyleMediumDashDot
	BorderStyleDashDotDot1 = BorderStyleDashDotDot
	BorderStyleDashDotDot2 = BorderStyleMediumDashDotDot
	BorderStyleSlantDash   = BorderStyleSlantDashDot
)

type BorderColor string

const (
	BorderColorBlack BorderColor = "#000000"
	BorderColorWhite BorderColor = "#FFFFFF"
)

func Border(position BorderPosition, style BorderStyle, color BorderColor) excelize.Border {
	return excelize.Border{
		Type:  string(position),
		Color: string(color),
		Style: int(style),
	}
}

func BorderAround(style BorderStyle, color BorderColor) []excelize.Border {
	return []excelize.Border{
		{
			Type:  string(BorderPositionTop),
			Color: string(color),
			Style: int(style),
		},
		{
			Type:  string(BorderPositionLeft),
			Color: string(color),
			Style: int(style),
		},
		{
			Type:  string(BorderPositionRight),
			Color: string(color),
			Style: int(style),
		},
		{
			Type:  string(BorderPositionBottom),
			Color: string(color),
			Style: int(style),
		},
	}
}

func FindBorder(borders []excelize.Border, position BorderPosition) (*excelize.Border, bool) {
	for _, b := range borders {
		if BorderPosition(b.Type) == position {
			return &b, true
		}
	}
	return nil, false
}

func ExistsBorder(borders []excelize.Border, position BorderPosition) bool {
	for _, b := range borders {
		if BorderPosition(b.Type) == position {
			return true
		}
	}
	return false
}

func Alignment(horizontal AlignmentHorizontal, vertical AlignmentVertical, wrapText bool) *excelize.Alignment {
	return &excelize.Alignment{
		Horizontal: string(horizontal),
		Vertical:   string(vertical),
		WrapText:   wrapText,
	}
}

func Fill(pattern FillPattern, color string) excelize.Fill {
	return excelize.Fill{
		Type:    string(FillTypePattern),
		Pattern: int(pattern),
		Color:   []string{color},
	}
}
