package model

import (
	"math"
	"syscall"
	"unsafe"
)

func S(text string) uintptr {
	ptrTxt, _ := syscall.UTF16PtrFromString(text)
	return uintptr(unsafe.Pointer(ptrTxt))
}

func S_P(text *string) uintptr {
	ptrTxt, _ := syscall.UTF16PtrFromString(*text)
	return uintptr(unsafe.Pointer(ptrTxt))
}

func I_P(number *int) uintptr {
	return uintptr(unsafe.Pointer(number))
}

func F(number float64) uintptr {
	return uintptr(math.Float64bits(number))
}

func F_P(number *float64) uintptr {
	return uintptr(unsafe.Pointer(number))
}

func I(number int) uintptr {
	return uintptr(number)
	// return uintptr(unsafe.Pointer(&number))
}
