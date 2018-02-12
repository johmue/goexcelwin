package model

import "log"

type Format struct {
	xb   *Book
	self uintptr
}

const (
	COLOR_BLACK              = 8
	COLOR_WHITE              = 9
	COLOR_RED                = 10
	COLOR_BRIGHTGREEN        = 11
	COLOR_BLUE               = 12
	COLOR_YELLOW             = 13
	COLOR_PINK               = 14
	COLOR_TURQUOISE          = 15
	COLOR_DARKRED            = 16
	COLOR_GREEN              = 17
	COLOR_DARKBLUE           = 18
	COLOR_DARKYELLOW         = 19
	COLOR_VIOLET             = 20
	COLOR_TEAL               = 21
	COLOR_GRAY25             = 22
	COLOR_GRAY50             = 23
	COLOR_PERIWINKLE_CF      = 24
	COLOR_PLUM_CF            = 25
	COLOR_IVORY_CF           = 26
	COLOR_LIGHTTURQUOISE_CF  = 27
	COLOR_DARKPURPLE_CF      = 28
	COLOR_CORAL_CF           = 29
	COLOR_OCEANBLUE_CF       = 30
	COLOR_ICEBLUE_CF         = 31
	COLOR_DARKBLUE_CL        = 32
	COLOR_PINK_CL            = 33
	COLOR_YELLOW_CL          = 34
	COLOR_TURQUOISE_CL       = 35
	COLOR_VIOLET_CL          = 36
	COLOR_DARKRED_CL         = 37
	COLOR_TEAL_CL            = 38
	COLOR_BLUE_CL            = 39
	COLOR_SKYBLUE            = 40
	COLOR_LIGHTTURQUOISE     = 41
	COLOR_LIGHTGREEN         = 42
	COLOR_LIGHTYELLOW        = 43
	COLOR_PALEBLUE           = 44
	COLOR_ROSE               = 45
	COLOR_LAVENDER           = 46
	COLOR_TAN                = 47
	COLOR_LIGHTBLUE          = 48
	COLOR_AQUA               = 49
	COLOR_LIME               = 50
	COLOR_GOLD               = 51
	COLOR_LIGHTORANGE        = 52
	COLOR_ORANGE             = 53
	COLOR_BLUEGRAY           = 54
	COLOR_GRAY40             = 55
	COLOR_DARKTEAL           = 56
	COLOR_SEAGREEN           = 57
	COLOR_DARKGREEN          = 58
	COLOR_OLIVEGREEN         = 59
	COLOR_BROWN              = 60
	COLOR_PLUM               = 61
	COLOR_INDIGO             = 62
	COLOR_GRAY80             = 63
	COLOR_DEFAULT_FOREGROUND = 64
	COLOR_DEFAULT_BACKGROUND = 65

	AS_DATE           = 1
	AS_FORMULA        = 2
	AS_NUMERIC_STRING = 3

	NUMFORMAT_GENERAL                 = 0
	NUMFORMAT_NUMBER                  = 1
	NUMFORMAT_NUMBER_D2               = 2
	NUMFORMAT_NUMBER_SEP              = 3
	NUMFORMAT_NUMBER_SEP_D2           = 4
	NUMFORMAT_CURRENCY_NEGBRA         = 5
	NUMFORMAT_CURRENCY_NEGBRARED      = 6
	NUMFORMAT_CURRENCY_D2_NEGBRA      = 7
	NUMFORMAT_CURRENCY_D2_NEGBRARED   = 8
	NUMFORMAT_PERCENT                 = 9
	NUMFORMAT_PERCENT_D2              = 10
	NUMFORMAT_SCIENTIFIC_D2           = 11
	NUMFORMAT_FRACTION_ONEDIG         = 12
	NUMFORMAT_FRACTION_TWODIG         = 13
	NUMFORMAT_DATE                    = 14
	NUMFORMAT_CUSTOM_DMON_YY          = 15
	NUMFORMAT_CUSTOM_DMON             = 16
	NUMFORMAT_CUSTOM_MON_YY           = 17
	NUMFORMAT_CUSTOM_HMM_AM           = 18
	NUMFORMAT_CUSTOM_HMMSS_AM         = 19
	NUMFORMAT_CUSTOM_HMM              = 20
	NUMFORMAT_CUSTOM_HMMSS            = 21
	NUMFORMAT_CUSTOM_MDYYYY_HMM       = 22
	NUMFORMAT_NUMBER_SEP_NEGBRA       = 37
	NUMFORMAT_NUMBER_SEP_NEGBRARED    = 38
	NUMFORMAT_NUMBER_D2_SEP_NEGBRA    = 39
	NUMFORMAT_NUMBER_D2_SEP_NEGBRARED = 40
	NUMFORMAT_ACCOUNT                 = 41
	NUMFORMAT_ACCOUNTCUR              = 42
	NUMFORMAT_ACCOUNT_D2              = 43
	NUMFORMAT_ACCOUNT_D2_CUR          = 44
	NUMFORMAT_CUSTOM_MMSS             = 45
	NUMFORMAT_CUSTOM_H0MMSS           = 46
	NUMFORMAT_CUSTOM_MMSS0            = 47
	NUMFORMAT_CUSTOM000P0E_PLUS0      = 48
	NUMFORMAT_TEXT                    = 49

	ALIGNH_GENERAL     = 0
	ALIGNH_LEFT        = 1
	ALIGNH_CENTER      = 2
	ALIGNH_RIGHT       = 3
	ALIGNH_FILL        = 4
	ALIGNH_JUSTIFY     = 5
	ALIGNH_MERGE       = 6
	ALIGNH_DISTRIBUTED = 7

	ALIGNV_TOP         = 0
	ALIGNV_CENTER      = 1
	ALIGNV_BOTTOM      = 2
	ALIGNV_JUSTIFY     = 3
	ALIGNV_DISTRIBUTED = 4

	BORDERSTYLE_NONE             = 0
	BORDERSTYLE_THIN             = 1
	BORDERSTYLE_MEDIUM           = 2
	BORDERSTYLE_DASHED           = 3
	BORDERSTYLE_DOTTED           = 4
	BORDERSTYLE_THICK            = 5
	BORDERSTYLE_DOUBLE           = 6
	BORDERSTYLE_HAIR             = 7
	BORDERSTYLE_MEDIUMDASHED     = 8
	BORDERSTYLE_DASHDOT          = 9
	BORDERSTYLE_MEDIUMDASHDOT    = 10
	BORDERSTYLE_DASHDOTDOT       = 11
	BORDERSTYLE_MEDIUMDASHDOTDOT = 12
	BORDERSTYLE_SLANTDASHDOT     = 13
	BORDERDIAGONAL_NONE          = 0
	BORDERDIAGONAL_DOWN          = 1
	BORDERDIAGONAL_UP            = 2
	BORDERDIAGONAL_BOTH          = 3

	FILLPATTERN_NONE                = 0
	FILLPATTERN_SOLID               = 1
	FILLPATTERN_GRAY50              = 2
	FILLPATTERN_GRAY75              = 3
	FILLPATTERN_GRAY25              = 4
	FILLPATTERN_HORSTRIPE           = 5
	FILLPATTERN_VERSTRIPE           = 6
	FILLPATTERN_REVDIAGSTRIPE       = 7
	FILLPATTERN_DIAGSTRIPE          = 8
	FILLPATTERN_DIAGCROSSHATCH      = 9
	FILLPATTERN_THICKDIAGCROSSHATCH = 10
	FILLPATTERN_THINHORSTRIPE       = 11
	FILLPATTERN_THINVERSTRIPE       = 12
	FILLPATTERN_THINREVDIAGSTRIPE   = 13
	FILLPATTERN_THINDIAGSTRIPE      = 14
	FILLPATTERN_THINHORCROSSHATCH   = 15
	FILLPATTERN_THINDIAGCROSSHATCH  = 16
	FILLPATTERN_GRAY12P5            = 17
	FILLPATTERN_GRAY6P25            = 18
)

func (xf *Format) GetSelf() uintptr {
	return xf.self
}

func (xf *Format) GetBook() uintptr {
	return xf.xb.GetSelf()
}

// FontHandle xlFormatFontW(FormatHandle handle);
func (xf *Format) Font() *Font {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatFontW").
		Call(xf.self)

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Font in Format.Font()")
	}

	fh := Font{}
	fh.xb = xf.xb
	fh.self = tmp
	return &fh
}

// int xlFormatSetFontW(FormatHandle handle, FontHandle fontHandle);
func (xf *Format) SetFont(fontHandle *Font) int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatSetFontW").
		Call(xf.self, uintptr(fontHandle.self))
	return int(tmp)
}

// int xlFormatNumFormatW(FormatHandle handle);
func (xf *Format) NumFormat() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatNumFormatW").
		Call(xf.self)

	return int(tmp)
}

// void xlFormatSetNumFormatW(FormatHandle handle, int numFormat);
func (xf *Format) SetNumFormat(numFormat int) {
	xf.xb.lib.NewProc("xlFormatSetNumFormatW").
		Call(xf.self, I(numFormat))
}

// int xlFormatAlignHW(FormatHandle handle);
func (xf *Format) AlignH() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatAlignHW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetAlignHW(FormatHandle handle, int align);
func (xf *Format) SetAlignH(align int) {
	xf.xb.lib.NewProc("xlFormatSetAlignHW").
		Call(xf.self, I(align))
}

// int xlFormatAlignVW(FormatHandle handle);
func (xf *Format) AlignV() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatAlignVW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetAlignVW(FormatHandle handle, int align);
func (xf *Format) SetAlignV(align int) {
	xf.xb.lib.NewProc("xlFormatSetAlignVW").
		Call(xf.self, I(align))
}

// int xlFormatWrapW(FormatHandle handle);
func (xf *Format) Wrap() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatWrapW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetWrapW(FormatHandle handle, int wrap);
func (xf *Format) SetWrap(wrap int) {
	xf.xb.lib.NewProc("xlFormatSetWrapW").
		Call(xf.self, I(wrap))
}

// int xlFormatRotationW(FormatHandle handle);
func (xf *Format) Rotation() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatRotationW").
		Call(xf.self)
	return int(tmp)
}

// int xlFormatSetRotationW(FormatHandle handle, int rotation);
func (xf *Format) SetRotation(rotation int) int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatSetRotationW").
		Call(xf.self, I(rotation))
	return int(tmp)
}

// int xlFormatIndentW(FormatHandle handle);
func (xf *Format) Indent() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatIndentW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetIndentW(FormatHandle handle, int indent);
func (xf *Format) SetIndent(indent int) {
	xf.xb.lib.NewProc("xlFormatSetIndentW").
		Call(xf.self, I(indent))
}

// int xlFormatShrinkToFitW(FormatHandle handle);
func (xf *Format) ShrinkToFit() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatShrinkToFitW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetShrinkToFitW(FormatHandle handle, int shrinkToFit);
func (xf *Format) SetShrinkToFit(shrinkToFit int) {
	xf.xb.lib.NewProc("xlFormatSetShrinkToFitW").
		Call(xf.self, I(shrinkToFit))
}

// void xlFormatSetBorderW(FormatHandle handle, int style);
func (xf *Format) SetBorder(style int) {
	xf.xb.lib.NewProc("xlFormatSetBorderW").
		Call(xf.self, I(style))
}

// void xlFormatSetBorderColorW(FormatHandle handle, int color);
func (xf *Format) SetBorderColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetBorderColorW").
		Call(xf.self, I(color))
}

// int xlFormatBorderLeftW(FormatHandle handle);
func (xf *Format) BorderLeft() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderLeftW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderLeftW(FormatHandle handle, int style);
func (xf *Format) SetBorderLeft(style int) {
	xf.xb.lib.NewProc("xlFormatSetBorderLeftW").
		Call(xf.self, I(style))
}

// int xlFormatBorderRightW(FormatHandle handle);
func (xf *Format) BorderRight() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderRightW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderRightW(FormatHandle handle, int style);
func (xf *Format) SetBorderRight(style int) {
	xf.xb.lib.NewProc("xlFormatSetBorderRightW").
		Call(xf.self, I(style))
}

// int xlFormatBorderTopW(FormatHandle handle);
func (xf *Format) BorderTop() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderTopW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderTopW(FormatHandle handle, int style);
func (xf *Format) SetBorderTop(style int) {
	xf.xb.lib.NewProc("xlFormatSetBorderTopW").
		Call(xf.self, I(style))
}

// int xlFormatBorderBottomW(FormatHandle handle);
func (xf *Format) BorderBottom() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderBottomW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderBottomW(FormatHandle handle, int style);
func (xf *Format) SetBorderBottom(style int) {
	xf.xb.lib.NewProc("xlFormatSetBorderBottomW").
		Call(xf.self, I(style))
}

// int xlFormatBorderLeftColorW(FormatHandle handle);
func (xf *Format) BorderLeftColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderLeftColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderLeftColorW(FormatHandle handle, int color);
func (xf *Format) SetBorderLeftColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetBorderLeftColorW").
		Call(xf.self, I(color))
}

// int xlFormatBorderRightColorW(FormatHandle handle);
func (xf *Format) BorderRightColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderRightColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderRightColorW(FormatHandle handle, int color);
func (xf *Format) SetBorderRightColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetBorderRightColorW").
		Call(xf.self, I(color))
}

// int xlFormatBorderTopColorW(FormatHandle handle);
func (xf *Format) BorderTopColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderTopColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderTopColorW(FormatHandle handle, int color);
func (xf *Format) SetBorderTopColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetBorderTopColorW").
		Call(xf.self, I(color))
}

// int xlFormatBorderBottomColorW(FormatHandle handle);
func (xf *Format) BorderBottomColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderBottomColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderBottomColorW(FormatHandle handle, int color);
func (xf *Format) SetBorderBottomColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetBorderBottomColorW").
		Call(xf.self, I(color))
}

// int xlFormatBorderDiagonalW(FormatHandle handle);
func (xf *Format) BorderDiagonal() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderDiagonalW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderDiagonalW(FormatHandle handle, int border);
func (xf *Format) SetBorderDiagonal(border int) {
	xf.xb.lib.NewProc("xlFormatSetBorderDiagonalW").
		Call(xf.self, I(border))
}

// int xlFormatBorderDiagonalStyleW(FormatHandle handle);
func (xf *Format) BorderDiagonalStyle() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderDiagonalStyleW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderDiagonalStyleW(FormatHandle handle, int style);
func (xf *Format) SetBorderDiagonalStyle(style int) {
	xf.xb.lib.NewProc("xlFormatSetBorderDiagonalStyleW").
		Call(xf.self, I(style))
}

// int xlFormatBorderDiagonalColorW(FormatHandle handle);
func (xf *Format) BorderDiagonalColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatBorderDiagonalColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetBorderDiagonalColorW(FormatHandle handle, int color);
func (xf *Format) SetBorderDiagonalColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetBorderDiagonalColorW").
		Call(xf.self, I(color))
}

// int xlFormatFillPatternW(FormatHandle handle);
func (xf *Format) FillPattern() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatFillPatternW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetFillPatternW(FormatHandle handle, int pattern);
func (xf *Format) SetFillPattern(pattern int) {
	xf.xb.lib.NewProc("xlFormatSetFillPatternW").
		Call(xf.self, I(pattern))
}

// int xlFormatPatternForegroundColorW(FormatHandle handle);
func (xf *Format) PatternForegroundColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatPatternForegroundColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetPatternForegroundColorW(FormatHandle handle, int color);
func (xf *Format) SetPatternForegroundColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetPatternForegroundColorW").
		Call(xf.self, I(color))
}

// int xlFormatPatternBackgroundColorW(FormatHandle handle);
func (xf *Format) PatternBackgroundColor() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatPatternBackgroundColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetPatternBackgroundColorW(FormatHandle handle, int color);
func (xf *Format) SetPatternBackgroundColor(color int) {
	xf.xb.lib.NewProc("xlFormatSetPatternBackgroundColorW").
		Call(xf.self, I(color))
}

// int xlFormatLockedW(FormatHandle handle);
func (xf *Format) Locked() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatLockedW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetLockedW(FormatHandle handle, int locked);
func (xf *Format) SetLocked(locked int) {
	xf.xb.lib.NewProc("xlFormatSetLockedW").
		Call(xf.self, I(locked))
}

// int xlFormatHiddenW(FormatHandle handle);
func (xf *Format) Hidden() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFormatHiddenW").
		Call(xf.self)
	return int(tmp)
}

// void xlFormatSetHiddenW(FormatHandle handle, int hidden);
func (xf *Format) SetHidden(hidden int) {
	xf.xb.lib.NewProc("xlFormatSetHiddenW").
		Call(xf.self, I(hidden))
}
