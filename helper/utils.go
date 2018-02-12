package helper

import (
	"log"
	"time"
	"unicode/utf16"
	"unsafe"
)

// credits go to
// https://stackoverflow.com/questions/15323767/does-golang-have-if-x-in-construct-similar-to-python
func StringInSlice(a string, list []string) bool {
	for _, b := range list {
		if b == a {
			return true
		}
	}
	return false
}

// credits go to
// https://coderwall.com/p/cp5fya/measuring-execution-time-in-go
func TimeTrack(start time.Time, name string) {
	elapsed := time.Since(start)
	log.Printf("[Info] %s took %s", name, elapsed)
}

// converting char* from c (as uintptr) to a go string
// credits go to https://groups.google.com/forum/#!topic/golang-nuts/6Aq0sy5naGM
func UIntPtrToString(up uintptr) (s string) {
	if up != 0x0 {
		var us []uint16
		for p := up; ; p += 2 {
			u := *(*uint16)(unsafe.Pointer(p))
			if u == 0 {
				return string(utf16.Decode(us))
			}
			us = append(us, u)
		}
	}
	return ""
}
