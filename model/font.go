package model

import "github.com/johmue/goexcelwin/helper"

type Font struct {
	xb   *Book
	self uintptr
}

const (
	NORMAL      = 0
	SUPERSCRIPT = 1
	SUBSCRIPT   = 2

	UNDERLINE_NONE      = 0
	UNDERLINE_SINGLE    = 1
	UNDERLINE_DOUBLE    = 2
	UNDERLINE_SINGLEACC = 33
	UNDERLINE_DOUBLEACC = 34
)

// int xlFontSizeW(FontHandle handle);
func (xf *Font) Size() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontSizeW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetSizeW(FontHandle handle, int size);
func (xf *Font) SetSize(size int) {
	xf.xb.lib.NewProc("xlFontSetSizeW").
		Call(xf.self, I(size))
}

// int xlFontItalicW(FontHandle handle);
func (xf *Font) Italic() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontItalicW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetItalicW(FontHandle handle, int italic);
func (xf *Font) SetItalic(italic int) {
	xf.xb.lib.NewProc("xlFontSetItalicW").
		Call(xf.self, I(italic))
}

// int xlFontStrikeOutW(FontHandle handle);
func (xf *Font) StrikeOut() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontStrikeOutW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetStrikeOutW(FontHandle handle, int strikeOut);
func (xf *Font) SetStrikeOut(strikeOut int) {
	xf.xb.lib.NewProc("xlFontSetStrikeOutW").
		Call(xf.self, I(strikeOut))
}

// int xlFontColorW(FontHandle handle);
func (xf *Font) Color() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontColorW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetColorW(FontHandle handle, int color);
func (xf *Font) SetColor(color int) {
	xf.xb.lib.NewProc("xlFontSetColorW").
		Call(xf.self, I(color))
}

// int xlFontBoldW(FontHandle handle);
func (xf *Font) Bold() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontBoldW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetBoldW(FontHandle handle, int bold);
func (xf *Font) SetBold(bold int) {
	xf.xb.lib.NewProc("xlFontSetBoldW").
		Call(xf.self, I(bold))
}

// int xlFontScriptW(FontHandle handle);
func (xf *Font) Script() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontScriptW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetScriptW(FontHandle handle, int script);
func (xf *Font) SetScript(script int) {
	xf.xb.lib.NewProc("xlFontSetScriptW").
		Call(xf.self, I(script))
}

// int xlFontUnderlineW(FontHandle handle);
func (xf *Font) Underline() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontUnderlineW").
		Call(xf.self)
	return int(tmp)
}

// void xlFontSetUnderlineW(FontHandle handle, int underline);
func (xf *Font) SetUnderline(underline int) {
	xf.xb.lib.NewProc("xlFontSetUnderlineW").
		Call(xf.self, I(underline))
}

// string xlFontNameW(FontHandle handle);
func (xf *Font) FontName(underline int) string {
	tmp, _, _ := xf.xb.lib.NewProc("xlFontNameW").
		Call(xf.self)
	return helper.UIntPtrToString(tmp)
}

// void xlFontSetNameW(FontHandle handle, const wchar_t* name);
func (xf *Font) SetName(name string) {
	xf.xb.lib.NewProc("xlFontSetNameW").
		Call(xf.self, S(name))
}
