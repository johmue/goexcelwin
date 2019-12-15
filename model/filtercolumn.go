package model

import (
	"github.com/johmue/goexcelwin/helper"
	"math"
)

type FilterColumn struct {
	xb   *Book
	self uintptr
}

const (
	FILTER_VALUE   = 0
	FILTER_TOP10   = 1
	FILTER_CUSTOM  = 2
	FILTER_DYNAMIC = 3
	FILTER_COLOR   = 4
	FILTER_ICON    = 5
	FILTER_EXT     = 6
	FILTER_NOT_SET = 7

	OPERATOR_EQUAL                 = 0
	OPERATOR_GREATER_THAN          = 1
	OPERATOR_GREATER_THAN_OR_EQUAL = 2
	OPERATOR_LESS_THAN             = 3
	OPERATOR_LESS_THAN_OR_EQUAL    = 4
	OPERATOR_NOT_EQUAL             = 5
)

func (xf *FilterColumn) GetSelf() uintptr {
	return xf.self
}

// int xlFilterColumnIndexW(FilterColumnHandle handle);
func (xf *FilterColumn) Index() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFilterColumnIndexW").
		Call(xf.self)
	return int(tmp)
}

// int xlFilterColumnFilterTypeW(FilterColumnHandle handle);
func (xf *FilterColumn) FilterType() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFilterColumnFilterTypeW").
		Call(xf.self)
	return int(tmp)
}

// int xlFilterColumnFilterSizeW(FilterColumnHandle handle);
func (xf *FilterColumn) FilterSize() int {
	tmp, _, _ := xf.xb.lib.NewProc("xlFilterColumnFilterSizeW").
		Call(xf.self)
	return int(tmp)
}

// string xlFilterColumnFilterW(FilterColumnHandle handle, int index);
func (xf *FilterColumn) Filter(index int) string {
	tmp, _, _ := xf.xb.lib.NewProc("xlFilterColumnFilterW").
		Call(xf.self, I(index))
	return helper.UIntPtrToString(tmp)
}

// void xlFilterColumnAddFilterW(FilterColumnHandle handle, const wchar_t* value);
func (xf *FilterColumn) AddFilter(value string) {
	xf.xb.lib.NewProc("xlFilterColumnAddFilterW").
		Call(xf.self, S(value))
}

// int xlFilterColumnGetTop10W(FilterColumnHandle handle, double* value, int* top, int* percent);
func (xf *FilterColumn) GetTop10() (value float64, top int, percent int) {
	xf.xb.lib.NewProc("xlFilterColumnGetTop10W").
		Call(xf.self, F_P(&value), I_P(&top), I_P(&percent))
	return float64(value), int(top), int(percent)
}

// void xlFilterColumnSetTop10W(FilterColumnHandle handle, double value, int top, int percent);
func (xf *FilterColumn) SetTop10(value float64, top int, percent int) {
	xf.xb.lib.NewProc("xlFilterColumnSetTop10W").
		Call(xf.self, uintptr(math.Float64bits(value)), I(top), I(percent))
}

// int xlFilterColumnGetCustomFilterW(FilterColumnHandle handle, int* op1, const wchar_t** v1, int* op2, const wchar_t** v2, int* andOp);
func (xf *FilterColumn) GetCustomFilter() (op1 int, v1 string, op2 int, v2 string, andOp int) {
	xf.xb.lib.NewProc("xlFilterColumnGetCustomFilterW").
		Call(xf.self, I_P(&op1), S_P(&v1), I_P(&op2), S_P(&v2), I_P(&andOp)) // @todo
	return int(op1), string(v1), int(op2), string(v2), int(andOp)
}

// void xlFilterColumnSetCustomFilterW(FilterColumnHandle handle, int op, const wchar_t* val);
func (xf *FilterColumn) SetCustomFilter(op int, val string) {
	xf.xb.lib.NewProc("xlFilterColumnSetCustomFilterW").
		Call(xf.self, I(op), S(val))
}

// void xlFilterColumnSetCustomFilterExW(FilterColumnHandle handle, int op1, const wchar_t* v1, int op2, const wchar_t* v2, int andOp);
func (xf *FilterColumn) SetCustomFilterEx(op1 int, v1 string, op2 int, v2 string, andOp int) {
	xf.xb.lib.NewProc("xlFilterColumnSetCustomFilterExW").
		Call(xf.self, I(op1), S(v1), I(op2), S(v2), I(andOp))
}

// void xlFilterColumnClearW(FilterColumnHandle handle);
func (xf *FilterColumn) Clear() {
	xf.xb.lib.NewProc("xlFilterColumnClearW").
		Call(xf.self)
}
