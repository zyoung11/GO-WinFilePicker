# Windows Modern File and Folder Picker for Go

**GO-WinFilePicker** is a lightweight Go library that provides modern, native file and folder selection dialogs for Windows applications. This package uses Windows COM APIs through Golang's syscall to provide native dialogs without external dependencies.

[![Go Reference](https://pkg.go.dev/badge/github.com/zyoung11/GO-WinFilePicker.svg)](https://pkg.go.dev/github.com/zyoung11/GO-WinFilePicker)

## Installation

```bash
go get github.com/zyoung11/GO-WinFilePicker
```

## Quick Start

```go
package main

import (
    "fmt"
    "github.com/zyoung11/GO-WinFilePicker"
)

func main() {

	// Select File Example
	file, err := winfilepicker.SelectFile("Please select a file")
	if err != nil {
		fmt.Println("File selection error:", err)
	} else {
		fmt.Println("Selected:", file)
	}

	// Select Folder Example
	folder, err := winfilepicker.SelectFolder("Please select a folder")
	if err != nil {
		fmt.Println("Folder selection error:", err)
	} else {
		fmt.Println("Selected:", folder)
	}
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
