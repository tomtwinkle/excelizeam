package excelizestyle

import (
	"github.com/xuri/excelize/v2"
)

type BorderPosition string

const (
	BorderPositionTop    BorderPosition = "top"
	BorderPositionLeft   BorderPosition = "left"
	BorderPositionRight  BorderPosition = "right"
	BorderPositionBottom BorderPosition = "bottom"
)

// BorderStyle https://xuri.me/excelize/ja/style.html#border
type BorderStyle int

const (
	BorderStyleNone        BorderStyle = 0  // Weight: 0, Style:
	BorderStyleContinuous0 BorderStyle = 7  // Weight: 0, Style: -----------
	BorderStyleContinuous1 BorderStyle = 1  // Weight: 1, Style: -----------
	BorderStyleContinuous2 BorderStyle = 2  // Weight: 2, Style: -----------
	BorderStyleContinuous3 BorderStyle = 5  // Weight: 3, Style: -----------
	BorderStyleDash1       BorderStyle = 3  // Weight: 1, Style: - - - - - -
	BorderStyleDash2       BorderStyle = 8  // Weight: 2, Style: - - - - - -
	BorderStyleDot         BorderStyle = 4  // Weight: 1, Style: . . . . . .
	BorderStyleDouble      BorderStyle = 6  // Weight: 3, Style: ===========
	BorderStyleDashDot1    BorderStyle = 9  // Weight: 1, Style: - . - . - .
	BorderStyleDashDot2    BorderStyle = 10 // Weight: 2, Style: - . - . - .
	BorderStyleDashDotDot1 BorderStyle = 11 // Weight: 1, Style: - . . - . .
	BorderStyleDashDotDot2 BorderStyle = 12 // Weight: 2, Style: - . . - . .
	BorderStyleSlantDash   BorderStyle = 13 // Weight: 2, Style: / - . / - .
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

// AlignmentHorizontal https://xuri.me/excelize/ja/style.html#align
type AlignmentHorizontal string

const (
	AlignmentHorizontalLeft             AlignmentHorizontal = "left"             // Left (indented)
	AlignmentHorizontalCenter           AlignmentHorizontal = "center"           // Centered
	AlignmentHorizontalRight            AlignmentHorizontal = "right"            // Right (indented)
	AlignmentHorizontalFill             AlignmentHorizontal = "fill"             // Filling
	AlignmentHorizontalJustify          AlignmentHorizontal = "justify"          // Justified
	AlignmentHorizontalCenterContinuous AlignmentHorizontal = "centerContinuous" // Cross-column centered
	AlignmentHorizontalDistributed      AlignmentHorizontal = "distributed"      // Decentralized alignment (indented)
)

// AlignmentVertical https://xuri.me/excelize/ja/style.html#align
type AlignmentVertical string

const (
	AlignmentVerticalTop         AlignmentVertical = "top"         // Top alignment
	AlignmentVerticalCenter      AlignmentVertical = "center"      // Centered
	AlignmentVerticalJustify     AlignmentVertical = "justify"     // Justified
	AlignmentVerticalDistributed AlignmentVertical = "distributed" // Decentralized alignment
)

func Alignment(horizontal AlignmentHorizontal, vertical AlignmentVertical, wrapText bool) *excelize.Alignment {
	return &excelize.Alignment{
		Horizontal: string(horizontal),
		Vertical:   string(vertical),
		WrapText:   wrapText,
	}
}

// FillPattern https://xuri.me/excelize/ja/style.html#pattern
type FillPattern int

const (
	FillPatternNone            FillPattern = 0
	FillPatternSolid           FillPattern = 1
	FillPatternMediumGray      FillPattern = 2
	FillPatternDarkGray        FillPattern = 3
	FillPatternLightGray       FillPattern = 4
	FillPatternDarkHorizontal  FillPattern = 5
	FillPatternDarkVertical    FillPattern = 6
	FillPatternDarkDown        FillPattern = 7
	FillPatternDarkUp          FillPattern = 8
	FillPatternDarkGrid        FillPattern = 9
	FillPatternDarkTrellis     FillPattern = 10
	FillPatternLightHorizontal FillPattern = 11
	FillPatternLightVertical   FillPattern = 12
	FillPatternLightDown       FillPattern = 13
	FillPatternLightUp         FillPattern = 14
	FillPatternLightGrid       FillPattern = 15
	FillPatternLightTrellis    FillPattern = 16
	FillPatternGray125         FillPattern = 17
	FillPatternGray0625        FillPattern = 18
)

func Fill(pattern FillPattern, color string) excelize.Fill {
	return excelize.Fill{
		Type:    "pattern",
		Pattern: int(pattern),
		Color:   []string{color},
	}
}
