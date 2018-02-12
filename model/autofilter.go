package model

import "log"

type AutoFilter struct {
	xb   *Book
	self uintptr
}

func (xa *AutoFilter) GetSelf() uintptr {
	return xa.self
}

// int xlAutoFilterGetRefW(AutoFilterHandle handle, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
func (xa *AutoFilter) GetRef() (rowFirst int, rowLast int, colFirst int, colLast int) {
	xa.xb.lib.NewProc("xlAutoFilterGetRefW").
		Call(xa.self, I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast))
	return int(rowFirst), int(rowLast), int(colFirst), int(colLast)
}

// void xlAutoFilterSetRefW(AutoFilterHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);
func (xa *AutoFilter) SetRef(rowFirst int, rowLast int, colFirst int, colLast int) {
	xa.xb.lib.NewProc("xlAutoFilterSetRefW").
		Call(xa.self, I(rowFirst), I(rowLast), I(colFirst), I(colLast))
}

// FilterColumnHandle xlAutoFilterColumnW(AutoFilterHandle handle, int colId);
func (xa *AutoFilter) Column(colId int) *FilterColumn {
	tmp, _, _ := xa.xb.lib.NewProc("xlAutoFilterColumnW").
		Call(xa.self, I(colId))

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize FilterColumn in AutoFilter.Column()")
	}

	fc := FilterColumn{}
	fc.xb = xa.xb
	fc.self = tmp
	return &fc
}

// int xlAutoFilterColumnSizeW(AutoFilterHandle handle);
func (xa *AutoFilter) ColumnSize() int {
	tmp, _, _ := xa.xb.lib.NewProc("xlAutoFilterColumnSizeW").
		Call(xa.self)

	return int(tmp)
}

// FilterColumnHandle xlAutoFilterColumnByIndexW(AutoFilterHandle handle, int index);
func (xa *AutoFilter) ColumnByIndex(index int) *FilterColumn {
	tmp, _, _ := xa.xb.lib.NewProc("xlAutoFilterColumnByIndexW").
		Call(xa.self, I(index))

	if tmp == uintptr(0) {
		log.Fatal("failed to initialize FilterColumn in AutoFilter.ColumnByIndex()")
	}

	fc := FilterColumn{}
	fc.xb = xa.xb
	fc.self = tmp
	return &fc
}

// int xlAutoFilterGetSortRangeW(AutoFilterHandle handle, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
func (xa *AutoFilter) GetSortRange() (rowFirst int, rowLast int, colFirst int, colLast int) {
	xa.xb.lib.NewProc("xlAutoFilterGetSortRangeW").
		Call(xa.self, I_P(&rowFirst), I_P(&rowLast), I_P(&colFirst), I_P(&colLast))
	return int(rowFirst), int(rowLast), int(colFirst), int(colLast)
}

// int xlAutoFilterGetSortW(AutoFilterHandle handle, int* columnIndex, int* descending);
func (xa *AutoFilter) GetSort() (columnIndex int, descending int) {
	xa.xb.lib.NewProc("xlAutoFilterGetSortW").
		Call(xa.self, I_P(&columnIndex), I_P(&descending))
	return int(columnIndex), int(descending)
}

// int xlAutoFilterSetSortW(AutoFilterHandle handle, int columnIndex, int descending);
func (xa *AutoFilter) SetSort(columnIndex int, descending int) int {
	tmp, _, _ := xa.xb.lib.NewProc("xlAutoFilterSetSortW").
		Call(xa.self, I(columnIndex), I(descending))
	return int(tmp)
}
