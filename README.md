# goexcelwin
Go wrapper for LibXL excel library (Win64 only)

## Installation

```shell
go get github.com/johmue/goexcelwin
```

## Simple example

```go
// main.go
package main

import (
	"github.com/johmue/goexcelwin/model"
)

xb := model.Book{}
xb.CreateXLSX("./bin/libxl.dll")

xb.SetKey("<License Name>", "<License Key>")

xb.SetLocale("UTF-8")
xb.SetRgbMode(1)

xs := xb.AddSheet("Table1")

xs.WriteStr(1, 1, "Hello!")
xs.WriteNum(1, 2, 100)

xb.Save("test.xlsx")
```