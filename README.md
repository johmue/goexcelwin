# goexcelwin
Go wrapper (Win64 only) for commercial [LibXL](http://libxl.com/) excel library.

## Installation

```shell
go get github.com/johmue/goexcelwin
```

Copy the `libxl.dll` for Win64 into a directory `./bin` inside your project.

## Simple example

```go
// main.go
package main

import (
	"github.com/johmue/goexcelwin"
)

func main() {
	xb := goexcelwin.Book{}
	xb.CreateXLSX("./bin/libxl.dll")

	xb.SetKey("<License Name>", "<License Key>")

	xb.SetLocale("UTF-8")
	xb.SetRgbMode(1)

	xs := xb.AddSheet("Table1", nil)

	xs.WriteStr(1, 1, "Hello!", nil)
	xs.WriteNum(1, 2, 100, nil)

	xb.Save("test.xlsx")
}
```