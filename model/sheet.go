package model

import (
    "github.com/johmue/goexcelwin/helper"
	"log"
	"unsafe"
)

type Sheet struct {
	xb   *Book
	self uintptr
}

const (
	PAPER_DEFAULT              = 0
	PAPER_LETTER               = 1
	PAPER_LETTERSMALL          = 2
	PAPER_TABLOID              = 3
	PAPER_LEDGER               = 4
	PAPER_LEGAL                = 5
	PAPER_STATEMENT            = 6
	PAPER_EXECUTIVE            = 7
	PAPER_A3                   = 8
	PAPER_A4                   = 9
	PAPER_A4SMALL              = 10
	PAPER_A5                   = 11
	PAPER_B4                   = 12
	PAPER_B5                   = 13
	PAPER_FOLIO                = 14
	PAPER_QUATRO               = 15
	PAPER10X14                 = 16
	PAPER10X17                 = 17
	PAPER_NOTE                 = 18
	PAPER_ENVELOPE9            = 19
	PAPER_ENVELOPE10           = 20
	PAPER_ENVELOPE11           = 21
	PAPER_ENVELOPE12           = 22
	PAPER_ENVELOPE14           = 23
	PAPER_CSIZE                = 24
	PAPER_DSIZE                = 25
	PAPER_ESIZE                = 26
	PAPER_ENVELOPE_DL          = 27
	PAPER_ENVELOPE_C5          = 28
	PAPER_ENVELOPE_C3          = 29
	PAPER_ENVELOPE_C4          = 30
	PAPER_ENVELOPE_C6          = 31
	PAPER_ENVELOPE_C65         = 32
	PAPER_ENVELOPE_B4          = 33
	PAPER_ENVELOPE_B5          = 34
	PAPER_ENVELOPE_B6          = 35
	PAPER_ENVELOPE             = 36
	PAPER_ENVELOPE_MONARCH     = 37
	PAPER_US_ENVELOPE          = 38
	PAPER_FANFOLD              = 39
	PAPER_GERMAN_STD_FANFOLD   = 40
	PAPER_GERMAN_LEGAL_FANFOLD = 41

	CELLTYPE_EMPTY   = 0
	CELLTYPE_NUMBER  = 1
	CELLTYPE_STRING  = 2
	CELLTYPE_BOOLEAN = 3
	CELLTYPE_BLANK   = 4
	CELLTYPE_ERROR   = 5

	ERRORTYPE_NULL    = 0
	ERRORTYPE_DIV0    = 7
	ERRORTYPE_VALUE   = 15
	ERRORTYPE_REF     = 23
	ERRORTYPE_NAME    = 29
	ERRORTYPE_NUM     = 36
	ERRORTYPE_NA      = 42
	ERRORTYPE_NOERROR = 255

	LEFT_TO_RIGHT = 0
	RIGHT_TO_LEFT = 1

	IERR_EVAL_ERROR            = 1
	IERR_EMPTY_CELLREF         = 2
	IERR_NUMBER_STORED_AS_TEXT = 4
	IERR_INCONSIST_RANGE       = 8
	IERR_INCONSIST_FMLA        = 16
	IERR_TWODIG_TEXTYEAR       = 32
	IERR_UNLOCK_FMLA           = 64
	IERR_DATA_VALIDATION       = 128

	PROT_DEFAULT            = -1
	PROT_ALL                = 0
	PROT_OBJECTS            = 1
	PROT_SCENARIOS          = 2
	PROT_FORMAT_CELLS       = 4
	PROT_FORMAT_COLUMNS     = 8
	PROT_FORMAT_ROWS        = 16
	PROT_INSERT_COLUMNS     = 32
	PROT_INSERT_ROWS        = 64
	PROT_INSERT_HYPERLINKS  = 128
	PROT_DELETE_COLUMNS     = 256
	PROT_DELETE_ROWS        = 512
	PROT_SEL_LOCKED_CELLS   = 1024
	PROT_SORT               = 2048
	PROT_AUTOFILTER         = 4096
	PROT_PIVOTTABLES        = 8192
	PROT_SEL_UNLOCKED_CELLS = 16384

	SHEETSTATE_VISIBLE    = 0
	SHEETSTATE_HIDDEN     = 1
	SHEETSTATE_VERYHIDDEN = 2

	VALIDATION_TYPE_NONE             = 0
	VALIDATION_TYPE_WHOLE            = 1
	VALIDATION_TYPE_DECIMAL          = 2
	VALIDATION_TYPE_LIST             = 3
	VALIDATION_TYPE_DATE             = 4
	VALIDATION_TYPE_TIME             = 5
	VALIDATION_TYPE_TEXTLENGTH       = 6
	VALIDATION_TYPE_CUSTOM           = 7
	VALIDATION_OP_BETWEEN            = 0
	VALIDATION_OP_NOTBETWEEN         = 1
	VALIDATION_OP_EQUAL              = 2
	VALIDATION_OP_NOTEQUAL           = 3
	VALIDATION_OP_LESSTHAN           = 4
	VALIDATION_OP_LESSTHANOREQUAL    = 5
	VALIDATION_OP_GREATERTHAN        = 6
	VALIDATION_OP_GREATERTHANOREQUAL = 7
	VALIDATION_ERRSTYLE_STOP         = 0 // STOP ICON IN THE ERROR ALERT
	VALIDATION_ERRSTYLE_WARNING      = 1 // WARNING ICON IN THE ERROR ALERT
	VALIDATION_ERRSTYLE_INFORMATION  = 2 // INFORMATION ICON IN THE ERROR ALERT
)

// int xlSheetCellTypeW(SheetHandle handle, int row, int col);
func (xs *Sheet) CellType(row int, col int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetCellTypeW").
		Call(xs.self, I(row), I(col))
	return int(tmp)
}

// int xlSheetIsFormulaW(SheetHandle handle, int row, int col);
func (xs *Sheet) IsFormula(row int, col int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetIsFormulaW").
		Call(xs.self, I(row), I(col))
	return int(tmp)
}

// FormatHandle xlSheetCellFormatW(SheetHandle handle, int row, int col);
func (xs *Sheet) CellFormat(row int, col int) *Format {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetCellFormatW").
		Call(xs.self, I(row), I(col))

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize Format in Sheet.CellFormat()")
	}

	fo := Format{}
	fo.xb = xs.xb
	fo.self = tmp
	return &fo
}

// void xlSheetSetCellFormatW(SheetHandle handle, int row, int col, FormatHandle format);
func (xs *Sheet) SetCellFormat(row int, col int, format *Format) {
	xs.xb.lib.NewProc("xlSheetSetCellFormatW").
		Call(xs.self, I(row), I(col), format.self)
}

// string xlSheetReadStrW(SheetHandle handle, int row, int col, FormatHandle* format)
func (xs *Sheet) ReadStr(row int, col int, format *Format) string {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadStrW").
		Call(xs.self, I(row), I(col), fo)

	return helper.UIntPtrToString(tmp)
}

// int xlSheetWriteStrW(SheetHandle handle, int row, int col, const wchar_t* value, FormatHandle format);
func (xs *Sheet) WriteStr(row int, col int, value string, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteStrW").
		Call(xs.self, I(row), I(col), S(value), fo)
	return int(tmp)
}

// double xlSheetReadNumW(SheetHandle handle, int row, int col, FormatHandle* format);
func (xs *Sheet) ReadNum(row int, col int) (float64, *Format) {
	var fo uintptr
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadNumW").
		Call(xs.self, I(row), I(col), uintptr(unsafe.Pointer(&fo)))

	format := Format{}
	format.xb = xs.xb
	format.self = fo
	return float64(tmp), &format
}

// int xlSheetWriteNumW(SheetHandle handle, int row, int col, double value, FormatHandle format);
func (xs *Sheet) WriteNum(row int, col int, value float64, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteNumW").
		Call(xs.self, I(row), I(col), F(value), fo)

	return int(tmp)
}

// int xlSheetReadBoolW(SheetHandle handle, int row, int col, FormatHandle* format);
func (xs *Sheet) ReadBool(row int, col int) (int, *Format) {
	var fo uintptr
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadBoolW").
		Call(xs.self, I(row), I(col), uintptr(unsafe.Pointer(&fo)))

	format := Format{}
	format.xb = xs.xb
	format.self = fo
	return int(tmp), &format
}

// int xlSheetWriteBoolW(SheetHandle handle, int row, int col, int value, FormatHandle format);
func (xs *Sheet) WriteBool(row int, col int, value int, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteBoolW").
		Call(xs.self, I(row), I(col), I(value), fo)

	return int(tmp)
}

// int xlSheetReadBlankW(SheetHandle handle, int row, int col, FormatHandle* format);
func (xs *Sheet) ReadBlank(row int, col int, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadBlankW").
		Call(xs.self, I(row), I(col), fo)

	return int(tmp)
}

// int xlSheetWriteBlankW(SheetHandle handle, int row, int col, FormatHandle format);
func (xs *Sheet) WriteBlank(row int, col int, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteBlankW").
		Call(xs.self, I(row), I(col), fo)

	return int(tmp)
}

// string xlSheetReadFormulaW(SheetHandle handle, int row, int col, FormatHandle* format);
func (xs *Sheet) ReadFormula(row int, col int, format *Format) string {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadFormulaW").
		Call(xs.self, I(row), I(col), fo)

	return helper.UIntPtrToString(tmp)
}

// int xlSheetWriteFormulaW(SheetHandle handle, int row, int col, const wchar_t* value, FormatHandle format);
func (xs *Sheet) WriteFormula(row int, col int, value string, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteFormulaW").
		Call(xs.self, I(row), I(col), S(value), fo)

	return int(tmp)
}

// int xlSheetWriteFormulaNumW(SheetHandle handle, int row, int col, const wchar_t* expr, double value, FormatHandle format);
func (xs *Sheet) WriteFormulaNum(row int, col int, expr string, value float64, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteFormulaNumW").
		Call(xs.self, I(row), I(col), S(expr), F(value), fo)

	return int(tmp)
}

// int xlSheetWriteFormulaStrW(SheetHandle handle, int row, int col, const wchar_t* expr, const wchar_t* value, FormatHandle format);
func (xs *Sheet) WriteFormulaStr(row int, col int, expr string, value string, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteFormulaStrW").
		Call(xs.self, I(row), I(col), S(expr), S(value), fo)

	return int(tmp)
}

// int xlSheetWriteFormulaBoolW(SheetHandle handle, int row, int col, const wchar_t* expr, int value, FormatHandle format);
func (xs *Sheet) WriteFormulaBool(row int, col int, expr string, value int, format *Format) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetWriteFormulaBoolW").
		Call(xs.self, I(row), I(col), S(expr), I(value), fo)

	return int(tmp)
}

// string xlSheetReadCommentW(SheetHandle handle, int row, int col);
func (xs *Sheet) ReadComment(row int, col int) string {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadCommentW").
		Call(xs.self, I(row), I(col))
	return helper.UIntPtrToString(tmp)
}

// void xlSheetWriteCommentW(SheetHandle handle, int row, int col, const wchar_t* value, const wchar_t* author, int width, int height);
func (xs *Sheet) WriteComment(row int, col int, value string, author string, width int, height int) {
	xs.xb.lib.NewProc("xlSheetWriteCommentW").
		Call(xs.self, I(row), I(col), S(value), S(author), I(width), I(height))
}

// void xlSheetRemoveCommentW(SheetHandle handle, int row, int col);
func (xs *Sheet) RemoveComment(row int, col int) {
	xs.xb.lib.NewProc("xlSheetRemoveCommentW").
		Call(xs.self, I(row), I(col))
}

// int xlSheetIsDateW(SheetHandle handle, int row, int col);
func (xs *Sheet) IsDate(row int, col int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetIsDateW").
		Call(xs.self, I(row), I(col))
	return int(tmp)
}

// int xlSheetReadErrorW(SheetHandle handle, int row, int col);
func (xs *Sheet) ReadError(row int, col int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetReadErrorW").
		Call(xs.self, I(row), I(col))
	return int(tmp)
}

// void xlSheetWriteErrorW(SheetHandle handle, int row, int col, int error, FormatHandle format);
func (xs *Sheet) WriteError(row int, col int, error int, format *Format) {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}
	xs.xb.lib.NewProc("xlSheetWriteErrorW").
		Call(xs.self, I(row), I(col), I(error), fo)
}

// double xlSheetColWidthW(SheetHandle handle, int col);
func (xs *Sheet) ColWidth(col int) float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetColWidthW").
		Call(xs.self, I(col))
	return float64(tmp)
}

// double xlSheetRowHeightW(SheetHandle handle, int row);
func (xs *Sheet) RowHeight(row int) float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetRowHeightW").
		Call(xs.self, I(row))
	return float64(tmp)
}

// int xlSheetSetColW(SheetHandle handle, int colFirst, int colLast, double width, FormatHandle format, int hidden);
func (xs *Sheet) SetCol(colFirst int, colLast int, width float64, format *Format, hidden int) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetColW").
		Call(xs.self, I(colFirst), I(colLast), F(width), fo, I(hidden))

	return int(tmp)
}

// int xlSheetSetRowW(SheetHandle handle, int row, double height, FormatHandle format, int hidden);
func (xs *Sheet) SetRow(row int, height float64, format *Format, hidden int) int {
	fo := uintptr(0)
	if nil != format {
		fo = format.self
	}

	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetRowW").
		Call(xs.self, I(row), F(height), fo, I(hidden))

	return int(tmp)
}

// int xlSheetRowHiddenW(SheetHandle handle, int row);
func (xs *Sheet) RowHidden(row int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetRowHiddenW").
		Call(xs.self, I(row))
	return int(tmp)
}

// int xlSheetSetRowHiddenW(SheetHandle handle, int row, int hidden);
func (xs *Sheet) SetRowHidden(row int, hidden int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetRowHiddenW").
		Call(xs.self, I(row), I(hidden))
	return int(tmp)
}

// int xlSheetColHiddenW(SheetHandle handle, int col);
func (xs *Sheet) ColHidden(col int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetColHiddenW").
		Call(xs.self, I(col))
	return int(tmp)
}

// int xlSheetSetColHiddenW(SheetHandle handle, int col, int hidden);
func (xs *Sheet) SetColHidden(col int, hidden int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetColHiddenW").
		Call(xs.self, I(col), I(hidden))
	return int(tmp)
}

// int xlSheetGetMergeW(SheetHandle handle, int row, int col, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
func (xs *Sheet) GetMerge(row int, col int) (rowFirst int, rowLast int, colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetGetMergeW").
		Call(xs.self, I(row), I(col), I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast))
	return int(rowFirst), int(rowLast), int(colFirst), int(colLast)
}

// int xlSheetSetMergeW(SheetHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);
func (xs *Sheet) SetMerge(rowFirst int, rowLast int, colFirst int, colLast int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetMergeW").
		Call(xs.self, I(rowFirst), I(rowLast), I(colFirst), I(colLast))
	return int(tmp)
}

// int xlSheetDelMergeW(SheetHandle handle, int row, int col);
func (xs *Sheet) DelMerge(row int, col int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetDelMergeW").
		Call(xs.self, I(row), I(col))
	return int(tmp)
}

// int xlSheetMergeSizeW(SheetHandle handle);
func (xs *Sheet) MergeSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetMergeSizeW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetMergeW(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
func (xs *Sheet) Merge(index int) (rowFirst int, rowLast int, colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetMergeW").
		Call(xs.self, I(index), I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast))
	return int(rowFirst), int(rowLast), int(colFirst), int(colLast)
}

// int xlSheetDelMergeByIndexW(SheetHandle handle, int index);
func (xs *Sheet) DelMergeByIndex(index int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetDelMergeByIndexW").
		Call(xs.self, I(index))
	return int(tmp)
}

// int xlSheetPictureSizeW(SheetHandle handle);
func (xs *Sheet) PictureSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetPictureSizeW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetGetPictureW(SheetHandle handle, int index, int* rowTop, int* colLeft, int* rowBottom, int* colRight,
// 		int* width, int* height, int* offset_x, int* offset_y);
func (xs *Sheet) GetPicture(index int) (success int, rowTop int, colLeft int, rowBottom int, colRight int, width int, height int, offsetX int, offsetY int) {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGetPictureW").
		Call(xs.self, I(index), I_P(&rowTop), I_P(&colLeft), I_P(&rowBottom), I_P(&colRight), I_P(&width), I_P(&height), I_P(&offsetX), I_P(&offsetY))
	return int(tmp), int(rowTop), int(colLeft), int(rowBottom), int(colRight), int(width), int(height), int(offsetX), int(offsetY)
}

// void xlSheetSetPictureW(SheetHandle handle, int row, int col, int pictureId, double scale, int offset_x, int offset_y, int pos);
func (xs *Sheet) SetPicture(row int, col int, pictureId int, scale float64, offsetX int, offsetY int, pos int) {
	xs.xb.lib.NewProc("xlSheetSetPictureW").
		Call(xs.self, I(row), I(col), I(pictureId), F(scale), I(offsetX), I(offsetY), I(pos))
}

// void xlSheetSetPicture2W(SheetHandle handle, int row, int col, int pictureId, int width, int height, int offset_x, int offset_y, int pos);
func (xs *Sheet) SetPicture2(row int, col int, pictureId int, width int, height int, offsetX int, offsetY int, pos int) {
	xs.xb.lib.NewProc("xlSheetSetPicture2W").
		Call(xs.self, I(row), I(col), I(pictureId), I(width), I(height), I(offsetX), I(offsetY), I(pos))
}

// int xlSheetGetHorPageBreakW(SheetHandle handle, int index);
func (xs *Sheet) GetHorPageBreak(index int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGetHorPageBreakW").
		Call(xs.self, I(index))
	return int(tmp)
}

// int xlSheetGetHorPageBreakSizeW(SheetHandle handle);
func (xs *Sheet) GetHorPageBreakSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGetHorPageBreakSizeW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetGetVerPageBreakW(SheetHandle handle, int index);
func (xs *Sheet) GetVerPageBreak(index int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGetVerPageBreakW").
		Call(xs.self, I(index))
	return int(tmp)
}

// int xlSheetGetVerPageBreakSizeW(SheetHandle handle);
func (xs *Sheet) GetVerPageBreakSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGetVerPageBreakSizeW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetSetHorPageBreakW(SheetHandle handle, int row, int pageBreak);
func (xs *Sheet) SetHorPageBreak(row int, pageBreak int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetHorPageBreakW").
		Call(xs.self, I(row), I(pageBreak))
	return int(tmp)
}

// int xlSheetSetVerPageBreakW(SheetHandle handle, int col, int pageBreak);
func (xs *Sheet) SetVerPageBreak(col int, pageBreak int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetVerPageBreakW").
		Call(xs.self, I(col), I(pageBreak))
	return int(tmp)
}

// void xlSheetSplitW(SheetHandle handle, int row, int col);
func (xs *Sheet) Split(row int, col int) {
	xs.xb.lib.NewProc("xlSheetSplitW").
		Call(xs.self, I(row), I(col))
}

// int xlSheetSplitInfoW(SheetHandle handle, int* row, int* col);
func (xs *Sheet) SplitInfo() (row int, col int) {
	xs.xb.lib.NewProc("xlSheetSplitInfoW").
		Call(xs.self, I_P(&row), I_P(&col))
	return int(row), int(col)
}

// int xlSheetGroupRowsW(SheetHandle handle, int rowFirst, int rowLast, int collapsed);
func (xs *Sheet) GroupRows(rowFirst int, rowLast int, collapsed int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGroupRowsW").
		Call(xs.self, I(rowFirst), I(rowLast), I(collapsed))
	return int(tmp)
}

// int xlSheetGroupColsW(SheetHandle handle, int colFirst, int colLast, int collapsed);
func (xs *Sheet) GroupCols(colFirst int, colLast int, collapsed int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGroupColsW").
		Call(xs.self, I(colFirst), I(colLast), I(collapsed))
	return int(tmp)
}

// int xlSheetGroupSummaryBelowW(SheetHandle handle);
func (xs *Sheet) GroupSummaryBelow() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGroupSummaryBelowW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetGroupSummaryBelowW(SheetHandle handle, int below);
func (xs *Sheet) SetGroupSummaryBelow(below int) {
	xs.xb.lib.NewProc("xlSheetSetGroupSummaryBelowW").
		Call(xs.self, I(below))
}

// int xlSheetGroupSummaryRightW(SheetHandle handle);
func (xs *Sheet) GroupSummaryRight() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetGroupSummaryRightW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetGroupSummaryRightW(SheetHandle handle, int right);
func (xs *Sheet) SetGroupSummaryRight(right int) {
	xs.xb.lib.NewProc("xlSheetSetGroupSummaryRightW").
		Call(xs.self, I(right))
}

// void xlSheetClearW(SheetHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);
func (xs *Sheet) Clear(rowFirst int, rowLast int, colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetClearW").
		Call(xs.self, I(rowFirst), I(rowLast), I(colFirst), I(colLast))
}

// int xlSheetInsertColW(SheetHandle handle, int colFirst, int colLast);
func (xs *Sheet) InsertCol(colFirst int, colLast int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetInsertColW").
		Call(xs.self, I(colFirst), I(colLast))
	return int(tmp)
}

// int xlSheetInsertRowW(SheetHandle handle, int rowFirst, int rowLast);
func (xs *Sheet) InsertRow(rowFirst int, rowLast int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetInsertRowW").
		Call(xs.self, I(rowFirst), I(rowLast))
	return int(tmp)
}

// int xlSheetRemoveColW(SheetHandle handle, int colFirst, int colLast);
func (xs *Sheet) RemoveCol(colFirst int, colLast int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetRemoveColW").
		Call(xs.self, I(colFirst), I(colLast))
	return int(tmp)
}

// int xlSheetRemoveRowW(SheetHandle handle, int rowFirst, int rowLast);
func (xs *Sheet) RemoveRow(rowFirst int, rowLast int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetRemoveRowW").
		Call(xs.self, I(rowFirst), I(rowLast))
	return int(tmp)
}

// int xlSheetCopyCellW(SheetHandle handle, int rowSrc, int colSrc, int rowDst, int colDst);
func (xs *Sheet) CopyCell(rowSrc int, colSrc int, rowDst int, colDst int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetCopyCellW").
		Call(xs.self, I(rowSrc), I(colSrc), I(rowDst), I(colDst))
	return int(tmp)
}

// int xlSheetFirstRowW(SheetHandle handle);
func (xs *Sheet) FirstRow() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetFirstRowW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetLastRowW(SheetHandle handle);
func (xs *Sheet) LastRow() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetLastRowW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetFirstColW(SheetHandle handle);
func (xs *Sheet) FirstCol() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetFirstColW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetLastColW(SheetHandle handle);
func (xs *Sheet) LastCol() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetLastColW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetDisplayGridlinesW(SheetHandle handle);
func (xs *Sheet) DisplayGridlines() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetDisplayGridlinesW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetDisplayGridlinesW(SheetHandle handle, int show);
func (xs *Sheet) SetDisplayGridlines(show int) {
	xs.xb.lib.NewProc("xlSheetSetDisplayGridlinesW").
		Call(xs.self, I(show))
}

// int xlSheetPrintGridlinesW(SheetHandle handle);
func (xs *Sheet) PrintGridlines() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetPrintGridlinesW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetPrintGridlinesW(SheetHandle handle, int print);
func (xs *Sheet) SetPrintGridlines(print int) {
	xs.xb.lib.NewProc("xlSheetSetPrintGridlinesW").
		Call(xs.self, I(print))
}

// int xlSheetZoomW(SheetHandle handle);
func (xs *Sheet) Zoom() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetZoomW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetZoomW(SheetHandle handle, int zoom);
func (xs *Sheet) SetZoom(zoom int) {
	xs.xb.lib.NewProc("xlSheetSetZoomW").
		Call(xs.self, I(zoom))
}

// int xlSheetPrintZoomW(SheetHandle handle);
func (xs *Sheet) PrintZoom() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetPrintZoomW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetPrintZoomW(SheetHandle handle, int zoom);
func (xs *Sheet) SetPrintZoom(zoom int) {
	xs.xb.lib.NewProc("xlSheetSetPrintZoomW").
		Call(xs.self, I(zoom))
}

// int xlSheetGetPrintFitW(SheetHandle handle, int* wPages, int* hPages);
func (xs *Sheet) GetPrintFit() (wPages int, hPages int) {
	xs.xb.lib.NewProc("xlSheetGetPrintFitW").
		Call(xs.self, I_P(&wPages), I_P(&hPages))
	return wPages, hPages
}

// void xlSheetSetPrintFitW(SheetHandle handle, int wPages, int hPages);
func (xs *Sheet) SetPrintFit(wPages int, hPages int) {
	xs.xb.lib.NewProc("xlSheetSetPrintFitW").
		Call(xs.self, I(wPages), I(hPages))
}

// int xlSheetLandscapeW(SheetHandle handle);
func (xs *Sheet) Landscape() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetLandscapeW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetLandscapeW(SheetHandle handle, int landscape);
func (xs *Sheet) SetLandscape(landscape int) {
	xs.xb.lib.NewProc("xlSheetSetLandscapeW").
		Call(xs.self, I(landscape))
}

// int xlSheetPaperW(SheetHandle handle);
func (xs *Sheet) Paper() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetPaperW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetPaperW(SheetHandle handle, int paper);
func (xs *Sheet) SetPaper(paper int) {
	xs.xb.lib.NewProc("xlSheetSetPaperW").
		Call(xs.self, I(paper))
}

// string xlSheetHeaderW(SheetHandle handle)
func (xs *Sheet) Header() string {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetHeaderW").
		Call(xs.self)

	return helper.UIntPtrToString(tmp)
}

// int xlSheetSetHeaderW(SheetHandle handle, const wchar_t* header, double margin);
func (xs *Sheet) SetHeader(header string, margin float64) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetHeaderW").
		Call(xs.self, S(header), F(margin))
	return int(tmp)
}

// double xlSheetHeaderMarginW(SheetHandle handle);
func (xs *Sheet) HeaderMargin() float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetHeaderMarginW").
		Call(xs.self)
	return float64(tmp)
}

// string xlSheetFooterW(SheetHandle handle);
func (xs *Sheet) Footer() string {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetFooterW").
		Call(xs.self)

	return helper.UIntPtrToString(tmp)
}

// int xlSheetSetFooterW(SheetHandle handle, const wchar_t* footer, double margin);
func (xs *Sheet) SetFooter(footer string, margin float64) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetFooterW").
		Call(xs.self, S(footer), F(margin))
	return int(tmp)
}

// double xlSheetFooterMarginW(SheetHandle handle);
func (xs *Sheet) FooterMargin() float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetFooterMarginW").
		Call(xs.self)
	return float64(tmp)
}

// int xlSheetHCenterW(SheetHandle handle);
func (xs *Sheet) HCenter() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetHCenterW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetHCenterW(SheetHandle handle, int hCenter);
func (xs *Sheet) SetHCenter(hCenter int) {
	xs.xb.lib.NewProc("xlSheetSetHCenterW").
		Call(xs.self, I(hCenter))
}

// int xlSheetVCenterW(SheetHandle handle);
func (xs *Sheet) VCenter() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetVCenterW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetVCenterW(SheetHandle handle, int vCenter);
func (xs *Sheet) SetVCenter(vCenter int) {
	xs.xb.lib.NewProc("xlSheetSetVCenterW").
		Call(xs.self, I(vCenter))
}

// double xlSheetMarginLeftW(SheetHandle handle);
func (xs *Sheet) MarginLeft() float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetMarginLeftW").
		Call(xs.self)
	return float64(tmp)
}

// void xlSheetSetMarginLeftW(SheetHandle handle, double margin);
func (xs *Sheet) SetMarginLeft(margin float64) {
	xs.xb.lib.NewProc("xlSheetSetMarginLeftW").
		Call(xs.self, F(margin))
}

// double xlSheetMarginRightW(SheetHandle handle);
func (xs *Sheet) MarginRight() float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetMarginRightW").
		Call(xs.self)
	return float64(tmp)
}

// void xlSheetSetMarginRightW(SheetHandle handle, double margin);
func (xs *Sheet) SetMarginRight(margin float64) {
	xs.xb.lib.NewProc("xlSheetSetMarginRightW").
		Call(xs.self, F(margin))
}

// double xlSheetMarginTopW(SheetHandle handle);
func (xs *Sheet) MarginTop() float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetMarginTopW").
		Call(xs.self)
	return float64(tmp)
}

// void xlSheetSetMarginTopW(SheetHandle handle, double margin);
func (xs *Sheet) SetMarginTop(margin float64) {
	xs.xb.lib.NewProc("xlSheetSetMarginTopW").
		Call(xs.self, F(margin))
}

// double xlSheetMarginBottomW(SheetHandle handle);
func (xs *Sheet) MarginBottom() float64 {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetMarginBottomW").
		Call(xs.self)
	return float64(tmp)
}

// void xlSheetSetMarginBottomW(SheetHandle handle, double margin);
func (xs *Sheet) SetMarginBottom(margin float64) {
	xs.xb.lib.NewProc("xlSheetSetMarginBottomW").
		Call(xs.self, F(margin))
}

// int xlSheetPrintRowColW(SheetHandle handle);
func (xs *Sheet) PrintRowCol() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetPrintRowColW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetPrintRowColW(SheetHandle handle, int print);
func (xs *Sheet) SetPrintRowCol(print int) {
	xs.xb.lib.NewProc("xlSheetSetPrintRowColW").
		Call(xs.self, I(print))
}

// int xlSheetPrintRepeatRowsW(SheetHandle handle, int* rowFirst, int* rowLast);
func (xs *Sheet) PrintRepeatRows(rowFirst int, rowLast int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetPrintRepeatRowsW").
		Call(xs.self, I_P(&rowFirst), I_P(&rowLast))
	return int(tmp)
}

// void xlSheetSetPrintRepeatRowsW(SheetHandle handle, int rowFirst, int rowLast);
func (xs *Sheet) SetPrintRepeatRows(rowFirst int, rowLast int) {
	xs.xb.lib.NewProc("xlSheetSetPrintRepeatRowsW").
		Call(xs.self, I(rowFirst), I(rowLast))
}

// int xlSheetPrintRepeatColsW(SheetHandle handle, int* colFirst, int* colLast);
func (xs *Sheet) PrintRepeatCols() (colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetPrintRepeatColsW").
		Call(xs.self, I_P(&colFirst), I_P(&colLast))
	return int(colFirst), int(colLast)
}

// void xlSheetSetPrintRepeatColsW(SheetHandle handle, int colFirst, int colLast);
func (xs *Sheet) SetPrintRepeatCols(colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetSetPrintRepeatColsW").
		Call(xs.self, I(colFirst), I(colLast))
}

// int xlSheetPrintAreaW(SheetHandle handle, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
func (xs *Sheet) PrintArea() (rowFirst int, rowLast int, colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetPrintAreaW").
		Call(xs.self, I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast))
	return int(rowFirst), int(rowLast), int(colFirst), int(colLast)
}

// void xlSheetSetPrintAreaW(SheetHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);
func (xs *Sheet) SetPrintArea(rowFirst int, rowLast int, colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetSetPrintAreaW").
		Call(xs.self, I(rowFirst), I(rowLast), I(colFirst), I(colLast))
}

// void xlSheetClearPrintRepeatsW(SheetHandle handle);
func (xs *Sheet) ClearPrintRepeats() {
	xs.xb.lib.NewProc("xlSheetClearPrintRepeatsW").
		Call(xs.self)
}

// void xlSheetClearPrintAreaW(SheetHandle handle);
func (xs *Sheet) ClearPrintArea() {
	xs.xb.lib.NewProc("xlSheetClearPrintAreaW").
		Call(xs.self)
}

// int xlSheetGetNamedRangeW(SheetHandle handle, const wchar_t* name, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int scopeId, int* hidden);
func (xs *Sheet) GetNamedRange(name string) (rowFirst int, rowLast int, colFirst int, colLast int, scopeId int, hidden int) {
	xs.xb.lib.NewProc("xlSheetGetNamedRangeW").
		Call(xs.self, S(name), I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast), I_P(&scopeId), I_P(&hidden))
	return int(rowFirst), int(rowLast), int(colFirst), int(colLast), int(scopeId), int(hidden)
}

// int xlSheetSetNamedRangeW(SheetHandle handle, const wchar_t* name, int rowFirst, int rowLast, int colFirst, int colLast, int scopeId);
func (xs *Sheet) SetNamedRange(name string, rowFirst int, rowLast int, colFirst int, colLast int, scopeId int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetNamedRangeW").
		Call(xs.self, S(name), I(rowFirst), I(rowLast), I(colFirst), I(colLast), I(scopeId))
	return int(tmp)
}

// int xlSheetDelNamedRangeW(SheetHandle handle, const wchar_t* name, int scopeId);
func (xs *Sheet) DelNamedRange(name string, scopeId int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetDelNamedRangeW").
		Call(xs.self, S(name), I(scopeId))
	return int(tmp)
}

// int xlSheetNamedRangeSizeW(SheetHandle handle);
func (xs *Sheet) NamedRangeSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetNamedRangeSizeW").
		Call(xs.self)
	return int(tmp)
}

// string xlSheetNamedRangeW(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int* scopeId, int* hidden);
func (xs *Sheet) NamedRange(index int) (name string, rowFirst int, rowLast int, colFirst int, colLast int, scopeId int, hidden int) {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetNamedRangeW").
		Call(xs.self, I(index), I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast), I_P(&scopeId), I_P(&hidden))
	return helper.UIntPtrToString(tmp), int(rowFirst), int(rowLast), int(colFirst), int(colLast), int(scopeId), int(hidden)
}

// int xlSheetTableSizeW(SheetHandle handle);
func (xs *Sheet) TableSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetTableSizeW").
		Call(xs.self)
	return int(tmp)
}

// string xlSheetTableW(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int* headerRowCount, int* totalsRowCount);
func (xs *Sheet) Table(index int) (tableName string, rowFirst int, rowLast int, colFirst int, colLast int, headerRowCount int, totalsRowCount int) {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetTableW").
		Call(xs.self, I(index), I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast), I_P(&headerRowCount), I_P(&totalsRowCount))
	return helper.UIntPtrToString(tmp), int(rowFirst), int(rowLast), int(colFirst), int(colLast), int(headerRowCount), int(totalsRowCount)
}

// string xlSheetHyperlinkW(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
func (xs *Sheet) Hyperlink(index int) (hyperlink string, rowFirst int, rowLast int, colFirst int, colLast int) {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetHyperlinkW").
		Call(xs.self, I(index), I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast))
	return helper.UIntPtrToString(tmp), int(rowFirst), int(rowLast), int(colFirst), int(colLast)
}

// int xlSheetHyperlinkSizeW(SheetHandle handle);
func (xs *Sheet) HyperlinkSize() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetHyperlinkSizeW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetDelHyperlinkW(SheetHandle handle, int index);
func (xs *Sheet) DelHyperlink(index int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetDelHyperlinkW").
		Call(xs.self, I(index))
	return int(tmp)
}

// void xlSheetAddHyperlinkW(SheetHandle handle, const wchar_t* hyperlink, int rowFirst, int rowLast, int colFirst, int colLast);
func (xs *Sheet) AddHyperlink(hyperlink string, rowFirst int, rowLast int, colFirst int, colLast int) {
	xs.xb.lib.NewProc("xlSheetAddHyperlinkW").
		Call(xs.self, S(hyperlink), I(rowFirst), I(rowLast), I(colFirst), I(colLast))
}

// AutoFilterHandle xlSheetAutoFilterW(SheetHandle handle);
func (xs *Sheet) AutoFilter() *AutoFilter {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetAutoFilterW").
		Call(xs.self)

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize AutoFilter in Sheet.AutoFilter()")
	}

	af := AutoFilter{}
	af.xb = xs.xb
	af.self = tmp
	return &af
}

// void xlSheetApplyFilterW(SheetHandle handle);
func (xs *Sheet) ApplyFilter() {
	xs.xb.lib.NewProc("xlSheetApplyFilterW").
		Call(xs.self)
}

// void xlSheetRemoveFilterW(SheetHandle handle);
func (xs *Sheet) RemoveFilter() {
	xs.xb.lib.NewProc("xlSheetRemoveFilterW").
		Call(xs.self)
}

// string xlSheetNameW();
func (xs *Sheet) Name() string {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetNameW").
		Call(xs.self)

	return helper.UIntPtrToString(tmp)
}

// void xlSheetSetNameW(SheetHandle handle, const wchar_t* name);
func (xs *Sheet) SetName(name string) {
	xs.xb.lib.NewProc("xlSheetSetNameW").
		Call(xs.self, S(name))
}

// int xlSheetProtectW(SheetHandle handle);
func (xs *Sheet) Protect() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetProtectW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetProtectW(SheetHandle handle, int protect);
func (xs *Sheet) SetProtect(protect int) {
	xs.xb.lib.NewProc("xlSheetSetProtectW").
		Call(xs.self, I(protect))
}

// void xlSheetSetProtectExW(SheetHandle handle, int protect, const wchar_t* password, int enhancedProtection);
func (xs *Sheet) SetProtectEx(protect int, password string, enhancedProtection int) {
	xs.xb.lib.NewProc("xlSheetSetProtectExW").
		Call(xs.self, I(protect), S(password), I(enhancedProtection))
}

// int xlSheetHiddenW(SheetHandle handle);
func (xs *Sheet) Hidden() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetHiddenW").
		Call(xs.self)
	return int(tmp)
}

// int xlSheetSetHiddenW(SheetHandle handle, int hidden);
func (xs *Sheet) SetHidden(hidden int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetSetHiddenW").
		Call(xs.self, I(hidden))
	return int(tmp)
}

// void xlSheetGetTopLeftViewW(SheetHandle handle, int* row, int* col);
func (xs *Sheet) GetTopLeftView() (row int, col int) {
	xs.xb.lib.NewProc("xlSheetGetTopLeftViewW").
		Call(xs.self, I_P(&row), I_P(&col))
	return row, col
}

// void xlSheetSetTopLeftViewW(SheetHandle handle, int row, int col);
func (xs *Sheet) SetTopLeftView(row int, col int) {
	xs.xb.lib.NewProc("xlSheetSetTopLeftViewW").
		Call(xs.self, I(row), I(col))
}

// int xlSheetRightToLeftW(SheetHandle handle);
func (xs *Sheet) RightToLeft() int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetRightToLeftW").
		Call(xs.self)
	return int(tmp)
}

// void xlSheetSetRightToLeftW(SheetHandle handle, int rightToLeft);
func (xs *Sheet) SetRightToLeft(rightToLeft int) {
	xs.xb.lib.NewProc("xlSheetSetRightToLeftW").
		Call(xs.self, I(rightToLeft))
}

// void xlSheetSetAutoFitAreaW(SheetHandle handle, int rowFirst, int colFirst, int rowLast, int colLast);
func (xs *Sheet) SetAutoFitArea(rowFirst int, colFirst int, rowLast int, colLast int) {
	xs.xb.lib.NewProc("xlSheetSetAutoFitAreaW").
		Call(xs.self, I(rowFirst), I(colFirst), I(rowLast), I(colLast))
}

// void xlSheetAddrToRowColW(SheetHandle handle, const wchar_t* addr, int* row, int* col, int* rowRelative, int* colRelative);
func (xs *Sheet) AddrToRowCol(addr string) (row int, col int, rowRelative int, colRelative int) {
	xs.xb.lib.NewProc("xlSheetAddrToRowColW").
		Call(xs.self, S(addr), I_P(&row), I_P(&col), I_P(&rowRelative), I_P(&colRelative))

	return row, col, rowRelative, colRelative
}

// string xlSheetRowColToAddrW(SheetHandle handle, int row, int col, int rowRelative, int colRelative);
func (xs *Sheet) RowColToAddr(row int, col int, rowRelative int, colRelative int) string {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetRowColToAddrW").
		Call(xs.self, I(row), I(col), I(rowRelative), I(colRelative))

	return helper.UIntPtrToString(tmp)
}

// void xlSheetSetTabColorW(SheetHandle handle, int color);
func (xs *Sheet) SetTabColor(color int) {
	xs.xb.lib.NewProc("xlSheetSetTabColorW").
		Call(xs.self, I(color))
}

// void xlSheetSetTabRgbColorW(SheetHandle handle, int red, int green, int blue);
func (xs *Sheet) SetTabRgbColor(red int, green int, blue int) {
	xs.xb.lib.NewProc("xlSheetSetTabRgbColorW").
		Call(xs.self, I(red), I(green), I(blue))
}

// int xlSheetAddIgnoredErrorW(SheetHandle handle, int rowFirst, int colFirst, int rowLast, int colLast, int iError);
func (xs *Sheet) AddIgnoredError(rowFirst int, colFirst int, rowLast int, colLast int, iError int) int {
	tmp, _, _ := xs.xb.lib.NewProc("xlSheetAddIgnoredErrorW").
		Call(xs.self, I(rowFirst), I(colFirst), I(rowLast), I(colLast), I(iError))
	return int(tmp)
}

// void xlSheetAddDataValidationW(SheetHandle handle, int type, int op, int rowFirst, int rowLast, int colFirst, int colLast, const wchar_t* value1, const wchar_t* value2);
func (xs *Sheet) AddDataValidation(typ int, op int, rowFirst int, rowLast int, colFirst int, colLast int, value1 string, value2 string) {
	xs.xb.lib.NewProc("xlSheetAddDataValidationW").
		Call(xs.self, I(typ), I(op), I(rowFirst), I(rowLast), I(colFirst), I(colLast), S(value1), S(value2))
}

// void xlSheetAddDataValidationExW(SheetHandle handle, int type, int op, int rowFirst, int rowLast, int colFirst, int colLast, const wchar_t* value1, const wchar_t* value2,
//      int allowBlank, int hideDropDown, int showInputMessage, int showErrorMessage, const wchar_t* promptTitle, const wchar_t* prompt,
//      const wchar_t* errorTitle, const wchar_t* error, int errorStyle);
func (xs *Sheet) AddDataValidationEx(vtype int, op int, rowFirst int, rowLast int, colFirst int, colLast int, value1 string, value2 string,
	allowBlank int, hideDropDown int, showInputMessage int, showErrorMessage int, promptTitle string, prompt string,
	errorTitle string, error string, errorStyle int) {
	xs.xb.lib.NewProc("xlSheetAddDataValidationExW").
		Call(xs.self, I(vtype), I(op), I(rowFirst), I(rowLast), I(colFirst), I(colLast), S(value1), S(value2),
			I(allowBlank), I(hideDropDown), I(showInputMessage), I(showErrorMessage), S(promptTitle), S(prompt),
			S(errorTitle), S(error), I(errorStyle))
}

// void xlSheetAddDataValidationDoubleW(SheetHandle handle, int type, int op, int rowFirst, int rowLast, int colFirst, int colLast, double value1, double value2);
func (xs *Sheet) AddDataValidationDouble(typ int, op int, rowFirst int, rowLast int, colFirst int, colLast int, value1 float64, value2 float64) {
	xs.xb.lib.NewProc("xlSheetAddDataValidationDoubleW").
		Call(xs.self, I(typ), I(op), I(rowFirst), I(rowLast), I(colFirst), I(colLast), F(value1), F(value2))
}

// void xlSheetAddDataValidationDoubleExW(SheetHandle handle, int type, int op, int rowFirst, int rowLast, int colFirst, int colLast, double value1, double value2,
//  	int allowBlank, int hideDropDown, int showInputMessage, int showErrorMessage, const wchar_t* promptTitle, const wchar_t* prompt,
//      const wchar_t* errorTitle, const wchar_t* error, int errorStyle);
func (xs *Sheet) AddDataValidationDoubleEx(vtype int, op int, rowFirst int, rowLast int, colFirst int, colLast int, value1 float64, value2 float64,
	allowBlank int, hideDropDown int, showInputMessage int, showErrorMessage int, promptTitle string, prompt string,
	errorTitle string, error string, errorStyle int) {
	xs.xb.lib.NewProc("xlSheetAddDataValidationDoubleExW").
		Call(xs.self, I(vtype), I(op), I(rowFirst), I(rowLast), I(colFirst), I(colLast), F(value1), F(value2),
			I(allowBlank), I(hideDropDown), I(showInputMessage), I(showErrorMessage), S(promptTitle), S(prompt),
			S(errorTitle), S(error), I(errorStyle))
}

// void xlSheetRemoveDataValidationsW(SheetHandle handle);
func (xs *Sheet) RemoveDataValidations() {
	xs.xb.lib.NewProc("xlSheetRemoveDataValidationsW").
		Call(xs.self)
}

func (xs *Sheet) GetStrColumn(col int) (columnRows []string) {
	firstRow := 0
	lastRow := xs.LastRow()

	for rowIndex := firstRow; rowIndex < lastRow; rowIndex++ {
		columnRows = append(columnRows, xs.ReadStr(rowIndex, col, nil))
	}

	return columnRows
}
