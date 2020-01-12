package model

import (
	"github.com/johmue/goexcelwin/helper"
	"log"
	"math"
	"os"
	"syscall"
)

type Book struct {
	lib  *syscall.LazyDLL
	self uintptr
}

const (
	PICTURETYPE_DIB  = 3
	PICTURETYPE_EMF  = 4
	PICTURETYPE_JPEG = 1
	PICTURETYPE_PICT = 5
	PICTURETYPE_PNG  = 0
	PICTURETYPE_TIFF = 6
	PICTURETYPE_WMF  = 2

	SCOPE_UNDEFINED = -2
	SCOPE_WORKBOOK  = -1

	SHEETTYPE_CHART   = 1
	SHEETTYPE_CHEET   = 0
	SHEETTYPE_UNKNOWN = 2
)

func (xb *Book) GetSelf() uintptr {
	return xb.self
}

func (xb *Book) CreateXLSX(libxlDllFilePath string) {
	// check if libxl.dll exist at given location
	if _, err := os.Stat(libxlDllFilePath); os.IsNotExist(err) {
		log.Fatal("[Error] ", libxlDllFilePath, " not found\n")
	}
	xb.lib = syscall.NewLazyDLL(libxlDllFilePath)
	xb.self, _, _ = xb.lib.NewProc("xlCreateXMLBookCW").Call()
}

func (xb *Book) CreateXLS(libxlDllFilePath string) {
	// check if libxl.dll exist at given location
	if _, err := os.Stat(libxlDllFilePath); os.IsNotExist(err) {
		log.Fatal("[Error] ", libxlDllFilePath, " not found\n")
	}
	xb.lib = syscall.NewLazyDLL(libxlDllFilePath)
	xb.self, _, _ = xb.lib.NewProc("xlCreateBookCW").Call()
}

// int xlBookLoadW(BookHandle handle, const wchar_t* filename);
func (xb *Book) Load(filename string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadW").
		Call(xb.self, S(filename))
	return int(tmp)
}

// int xlBookSaveW(BookHandle handle, const wchar_t* filename);
func (xb *Book) Save(filename string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookSaveW").
		Call(xb.self, S(filename))
	return int(tmp)
}

// int xlBookLoadUsingTempFileW(BookHandle handle, const wchar_t* filename, const wchar_t* tempFile);
func (xb *Book) LoadUsingTempFile(filename string, tempFile string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadUsingTempFileW").
		Call(xb.self, S(filename), S(tempFile))
	return int(tmp)
}

// int xlBookSaveUsingTempFileW(BookHandle handle, const wchar_t* filename, int useTempFile);
func (xb *Book) SaveUsingTempFile(filename string, useTempFile int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookSaveUsingTempFileW").
		Call(xb.self, S(filename), I(useTempFile))
	return int(tmp)
}

// int xlBookLoadPartiallyW(BookHandle handle, const wchar_t* filename, int sheetIndex, int firstRow, int lastRow);
func (xb *Book) LoadPartially(filename string, sheetIndex int, firstRow int, lastRow int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadPartiallyW").
		Call(xb.self, S(filename), I(sheetIndex), I(firstRow), I(lastRow))
	return int(tmp)
}

// int xlBookLoadPartiallyUsingTempFileW(BookHandle handle, const wchar_t* filename, int sheetIndex, int firstRow, int lastRow, const wchar_t* tempFile);
func (xb *Book) LoadPartiallyUsingTempFile(filename string, sheetIndex int, firstRow int, lastRow int, tempFile string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadPartiallyUsingTempFileW").
		Call(xb.self, S(filename), I(sheetIndex), I(firstRow), I(lastRow), S(tempFile))
	return int(tmp)
}

// int xlBookLoadWithoutEmptyCellsW(BookHandle handle, const wchar_t* filename);
func (xb *Book) LoadWithoutEmptyCells(filename string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadWithoutEmptyCellsW").
		Call(xb.self, S(filename))
	return int(tmp)
}

// int xlBookLoadRawW(BookHandle handle, const char* data, unsigned size);
func (xb *Book) LoadRaw(data string, size int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadRawW").
		Call(xb.self, S(data), I(len(data)))
	return int(tmp)
}

// int xlBookLoadRawPartiallyW(BookHandle handle, const char* data, unsigned size, int sheetIndex, int firstRow, int lastRow);
func (xb *Book) LoadRawPartially(data string, size int, sheetIndex int, firstRow int, lastRow int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadRawPartiallyW").
		Call(xb.self, S(data), I(len(data)), I(sheetIndex), I(firstRow), I(lastRow))
	return int(tmp)
}

// int xlBookSaveRawW(BookHandle handle, const char** data, unsigned* size);
// @deprecated
func (xb *Book) SaveRaw() string {
	// @todo not implemented
	// @see Book.save()
	return ""
}

// SheetHandle xlBookAddSheetW(BookHandle handle, const wchar_t* name, SheetHandle initSheet);
func (xb *Book) AddSheet(name string, initSheet *Sheet) *Sheet {
	tmp := uintptr(0)
	if nil == initSheet {
		tmp, _, _ = xb.lib.NewProc("xlBookAddSheetW").
			Call(xb.self, S(name), 0)
	} else {
		tmp, _, _ = xb.lib.NewProc("xlBookAddSheetW").
			Call(xb.self, S(name), initSheet.self)
	}

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Sheet in Book.AddSheet()")
	}

	xs := Sheet{}
	xs.xb = xb
	xs.self = tmp
	return &xs
}

// SheetHandle xlBookInsertSheetW(BookHandle handle, int index, const wchar_t* name, SheetHandle initSheet);
func (xb *Book) InsertSheet(index int, name string, initSheet *Sheet) *Sheet {
	tmp := uintptr(0)

	if nil == initSheet {
		tmp, _, _ = xb.lib.NewProc("xlBookInsertSheetW").
			Call(xb.self, I(index), S(name), 0)
	} else {
		tmp, _, _ = xb.lib.NewProc("xlBookInsertSheetW").
			Call(xb.self, I(index), S(name), initSheet.self)
	}

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Sheet in Book.InsertSheet()")
	}

	xs := Sheet{}
	xs.xb = xb
	xs.self = tmp
	return &xs
}

// SheetHandle xlBookGetSheetW(BookHandle handle, int index);
func (xb *Book) GetSheet(index int) *Sheet {
	tmp, _, _ := xb.lib.NewProc("xlBookGetSheetW").
		Call(xb.self, I(index))

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Sheet in Book.GetSheet()")
	}

	xs := Sheet{}
	xs.xb = xb
	xs.self = tmp
	return &xs
}

// int xlBookSheetTypeW(BookHandle handle, int index);
func (xb *Book) SheetType(index int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookSheetTypeW").
		Call(xb.self, I(index))
	return int(tmp)
}

// int xlBookMoveSheetW(BookHandle handle, int srcIndex, int dstIndex);
func (xb *Book) MoveSheet(srcIndex int, dstIndex int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookMoveSheetW").
		Call(xb.self, I(srcIndex), I(dstIndex))
	return int(tmp)
}

// int xlBookDelSheetW(BookHandle handle, int index);
func (xb *Book) DelSheet(index int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookDelSheetW").
		Call(xb.self, I(index))
	return int(tmp)
}

// int xlBookSheetCountW(BookHandle handle);
func (xb *Book) SheetCount() int {
	tmp, _, _ := xb.lib.NewProc("xlBookSheetCountW").
		Call(xb.self)
	return int(tmp)
}

// FormatHandle xlBookAddFormatW(BookHandle handle, FormatHandle initFormat);
func (xb *Book) AddFormat() *Format {
	tmp, _, _ := xb.lib.NewProc("xlBookAddFormatW").
		Call(xb.self, 0)

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Format in Book.AddFormat()")
	}

	fh := Format{}
	fh.xb = xb
	fh.self = tmp
	return &fh
}

// FontHandle xlBookAddFontW(BookHandle handle, FontHandle initFont);
func (xb *Book) AddFont() *Font {
	tmp, _, _ := xb.lib.NewProc("xlBookAddFontW").
		Call(xb.self, 0)

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Font in Book.AddFont()")
	}

	fh := Font{}
	fh.xb = xb
	fh.self = tmp
	return &fh
}

// int xlBookAddCustomNumFormatW(BookHandle handle, const wchar_t* customNumFormat);
func (xb *Book) AddCustomNumFormat(customNumFormat string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookAddCustomNumFormatW").
		Call(xb.self, S(customNumFormat))
	return int(tmp)
}

// string xlBookCustomNumFormatW(BookHandle handle, int fmt);
func (xb *Book) CustomNumFormat(fmt int) string {
	tmp, _, _ := xb.lib.NewProc("xlBookCustomNumFormatW").
		Call(xb.self, I(fmt))
	return helper.UIntPtrToString(tmp)
}

// FormatHandle xlBookFormatW(BookHandle handle, int index);
func (xb *Book) Format(index int) *Format {
	tmp, _, _ := xb.lib.NewProc("xlBookFormatW").
		Call(xb.self, I(index))

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Format in Book.Format()")
	}

	format := Format{}
	format.xb = xb
	format.self = tmp
	return &format
}

// int xlBookFormatSizeW(BookHandle handle);
func (xb *Book) FormatSize() int {
	tmp, _, _ := xb.lib.NewProc("xlBookFormatSizeW").
		Call(xb.self)
	return int(tmp)
}

// FontHandle xlBookFontW(BookHandle handle, int index);
func (xb *Book) Font(index int) *Font {
	tmp, _, _ := xb.lib.NewProc("xlBookFontW").
		Call(xb.self, I(index))

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Font in Book.Font()")
	}

	fh := Font{}
	fh.xb = xb
	fh.self = tmp
	return &fh
}

// int xlBookFontSizeW(BookHandle handle);
func (xb *Book) FontSize() int {
	tmp, _, _ := xb.lib.NewProc("xlBookFontSizeW").
		Call(xb.self)
	return int(tmp)
}

// double xlBookDatePackW(BookHandle handle, int year, int month, int day, int hour, int min, int sec, int msec);
func (xb *Book) DatePack(year int, month int, day int, hour int, min int, sec int, msec int) float64 {
	tmp, _, _ := xb.lib.NewProc("xlBookDatePackW").
		Call(xb.self, I(year), I(month), I(day), I(hour), I(min), I(sec), I(msec))
	return float64(tmp)
}

// int xlBookDateUnpackW(BookHandle handle, double value, int* year, int* month, int* day, int* hour, int* min, int* sec, int* msec);
func (xb *Book) DateUnpack(value float64, year int, month int, day int, hour int, min int, sec int, msec int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookDateUnpackW").
		Call(xb.self, uintptr(math.Float64bits(value)), I_P(&year), I_P(&month), I_P(&day), I_P(&hour), I_P(&min), I_P(&sec), I_P(&msec))
	return int(tmp)
}

// int xlBookColorPackW(BookHandle handle, int red, int green, int blue);
func (xb *Book) ColorPack(red int, green int, blue int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookColorPackW").
		Call(xb.self, I_P(&red), I_P(&green), I_P(&blue))
	return int(tmp)
}

// void xlBookColorUnpackW(BookHandle handle, int color, int* red, int* green, int* blue);
func (xb *Book) ColorUnpack(color int) (red int, green int, blue int) {
	xb.lib.NewProc("xlBookColorUnpackW").
		Call(xb.self, I(color), I_P(&red), I_P(&green), I_P(&blue))
	return int(red), int(green), int(blue)
}

// int xlBookActiveSheetW(BookHandle handle);
func (xb *Book) ActiveSheet() int {
	tmp, _, _ := xb.lib.NewProc("xlBookActiveSheetW").
		Call(xb.self)
	return int(tmp)
}

// void xlBookSetActiveSheetW(BookHandle handle, int index);
func (xb *Book) SetActiveSheet(index int) {
	xb.lib.NewProc("xlBookSetActiveSheetW").
		Call(xb.self, I(index))
}

// int xlBookPictureSizeW(BookHandle handle);
func (xb *Book) PictureSize() int {
	tmp, _, _ := xb.lib.NewProc("xlBookPictureSizeW").
		Call(xb.self)
	return int(tmp)
}

// int xlBookGetPictureW(BookHandle handle, int index, const char** data, unsigned* size);
func (xb *Book) GetPicture(index int) string {
	var data string
	var dataLen int
	xb.lib.NewProc("xlBookGetPictureW").
		Call(xb.self, I(index), S_P(&data), I_P(&dataLen)) // @todo
	return string(data)
}

// int xlBookAddPictureW(BookHandle handle, const wchar_t* filename);
func (xb *Book) AddPicture(filename string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookAddPictureW").
		Call(xb.self, S(filename))
	return int(tmp)
}

// int xlBookAddPicture2W(BookHandle handle, const char* data, unsigned size);
func (xb *Book) AddPicture2(data string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookAddPicture2W").
		Call(xb.self, S(data), I(len(data)))
	return int(tmp)
}

// int xlBookAddPictureAsLinkW(BookHandle handle, const wchar_t* filename, int insert);
func (xb *Book) AddPictureAsLink(filename string, insert int) int {
	tmp, _, _ := xb.lib.NewProc("xlBookAddPictureAsLinkW").
		Call(xb.self, S(filename), I(insert))
	return int(tmp)
}

// string xlBookDefaultFontW(BookHandle handle, int* fontSize);
func (xb *Book) DefaultFont() (fontName string, fontSize int) {
	tmp, _, _ := xb.lib.NewProc("xlBookDefaultFontW").
		Call(xb.self, I_P(&fontSize))
	return helper.UIntPtrToString(tmp), int(fontSize)
}

// void xlBookSetDefaultFontW(BookHandle handle, const wchar_t* fontName, int fontSize);
func (xb *Book) SetDefaultFont(fontName string, fontSize int) {
	xb.lib.NewProc("xlBookSetDefaultFontW").
		Call(xb.self, S(fontName), I(fontSize))
}

// int xlBookRefR1C1W(BookHandle handle);
func (xb *Book) RefR1C1() int {
	tmp, _, _ := xb.lib.NewProc("xlBookRefR1C1W").
		Call(xb.self)
	return int(tmp)
}

// void xlBookSetRefR1C1W(BookHandle handle, int refR1C1);
func (xb *Book) SetRefR1C1(refR1C1 int) {
	xb.lib.NewProc("xlBookSetRefR1C1W").
		Call(xb.self, I(refR1C1))
}

// void xlBookSetKeyW(BookHandle handle, const wchar_t* name, const wchar_t* key);
func (xb *Book) SetKey(name string, key string) {
	xb.lib.NewProc("xlBookSetKeyW").
		Call(xb.self, S(name), S(key))
}

// int xlBookRgbModeW(BookHandle handle);
func (xb *Book) RgbMode() int {
	tmp, _, _ := xb.lib.NewProc("xlBookRgbModeW").
		Call(xb.self)
	return int(tmp)
}

// void xlBookSetRgbModeW(BookHandle handle, int rgbMode);
func (xb *Book) SetRgbMode(rgbMode int) {
	xb.lib.NewProc("xlBookSetRgbModeW").
		Call(xb.self, I(rgbMode))
}

// int xlBookVersionW(BookHandle handle);
func (xb *Book) Version() int {
	tmp, _, _ := xb.lib.NewProc("xlBookVersionW").
		Call(xb.self)
	return int(tmp)
}

// int xlBookBiffVersionW(BookHandle handle);
func (xb *Book) BiffVersion() int {
	tmp, _, _ := xb.lib.NewProc("xlBookBiffVersionW").
		Call(xb.self)
	return int(tmp)
}

// int xlBookIsDate1904W(BookHandle handle);
func (xb *Book) IsDate1904() int {
	tmp, _, _ := xb.lib.NewProc("xlBookIsDate1904W").
		Call(xb.self)
	return int(tmp)
}

// void xlBookSetDate1904W(BookHandle handle, int date1904);
func (xb *Book) SetDate1904(date1904 int) {
	xb.lib.NewProc("xlBookSetDate1904W").
		Call(xb.self, I(date1904))
}

// int xlBookIsTemplateW(BookHandle handle);
func (xb *Book) IsTemplate() int {
	tmp, _, _ := xb.lib.NewProc("xlBookIsTemplateW").
		Call(xb.self)
	return int(tmp)
}

// void xlBookSetTemplateW(BookHandle handle, int tmpl);
func (xb *Book) SetTemplate(tmpl int) {
	xb.lib.NewProc("xlBookSetTemplateW").
		Call(xb.self, I(tmpl))
}

// int xlBookSetLocaleW(BookHandle handle, const char* locale);
func (xb *Book) SetLocale(locale string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookSetLocaleW").
		Call(xb.self, S(locale))
	return int(tmp)
}

// string xlBookErrorMessageW(BookHandle handle);
func (xb *Book) ErrorMessage() string {
	tmp, _, _ := xb.lib.NewProc("xlBookErrorMessageW").
		Call(xb.self)
	return helper.UIntPtrToString(tmp)
}

// void xlBookReleaseW(BookHandle handle);
func (xb *Book) Release() {
	xb.lib.NewProc("xlBookReleaseW").
		Call(xb.self)
}

// int xlBookLoadInfo(BookHandle handle, const wchar_t* filename)
func (xb *Book) LoadInfo(filename string) int {
	tmp, _, _ := xb.lib.NewProc("xlBookLoadInfoW").
		Call(xb.self, S(filename))
	return int(tmp)
}

// const wchar_t* xlBookGetSheetName(BookHandle handle, int index)
func (xb *Book) GetSheetName(index int) string {
	tmp, _, _ := xb.lib.NewProc("xlBookGetSheetNameW").
		Call(xb.self, I(index))
	return helper.UIntPtrToString(tmp)
}
