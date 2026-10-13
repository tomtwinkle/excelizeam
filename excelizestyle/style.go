package excelizestyle

import (
	"github.com/xuri/excelize/v2"
)

// Preserve the existing border style names as aliases of the generated enums.
const (
	// BorderStyleContinuous0 is the compatibility alias for BorderStyleHair.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleContinuous0 = BorderStyleHair
	// BorderStyleContinuous1 is the compatibility alias for BorderStyleThin.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleContinuous1 = BorderStyleThin
	// BorderStyleContinuous2 is the compatibility alias for BorderStyleMedium.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleContinuous2 = BorderStyleMedium
	// BorderStyleContinuous3 is the compatibility alias for BorderStyleThick.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleContinuous3 = BorderStyleThick
	// BorderStyleDash1       is the compatibility alias for BorderStyleDashed.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDash1 = BorderStyleDashed
	// BorderStyleDash2       is the compatibility alias for BorderStyleMediumDashed.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDash2 = BorderStyleMediumDashed
	// BorderStyleDot         is the compatibility alias for BorderStyleDotted.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDot = BorderStyleDotted
	// BorderStyleDashDot1    is the compatibility alias for BorderStyleDashDot.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDashDot1 = BorderStyleDashDot
	// BorderStyleDashDot2    is the compatibility alias for BorderStyleMediumDashDot.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDashDot2 = BorderStyleMediumDashDot
	// BorderStyleDashDotDot1 is the compatibility alias for BorderStyleDashDotDot.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDashDotDot1 = BorderStyleDashDotDot
	// BorderStyleDashDotDot2 is the compatibility alias for BorderStyleMediumDashDotDot.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleDashDotDot2 = BorderStyleMediumDashDotDot
	// BorderStyleSlantDash   is the compatibility alias for BorderStyleSlantDashDot.
	// See https://xuri.me/excelize/en/style.html#border
	BorderStyleSlantDash = BorderStyleSlantDashDot
)

// BorderColor specifies an RGB border color, including custom hexadecimal values.
//
// See https://xuri.me/excelize/en/style.html#border
type BorderColor string

const (
	// BorderColorBlack selects the RGB color black (#000000).
	// See https://xuri.me/excelize/en/style.html#border
	BorderColorBlack BorderColor = "#000000"
	// BorderColorWhite selects the RGB color white (#FFFFFF).
	// See https://xuri.me/excelize/en/style.html#border
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
