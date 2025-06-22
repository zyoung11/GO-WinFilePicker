# Windows Modern File and Folder Picker for Go

**winfilepicker** is a lightweight Go library that provides modern, native file and folder selection dialogs for Windows applications. This package uses Windows COM APIs through Golang's syscall to provide native dialogs without external dependencies.

## Installation

```bash
go get github.com/zyoung11/GO-WinFilePicker
```

## Quick Start

```go
package main

import (
    "fmt"
    "github.com/yourgithub/winfilepicker"
)

func main() {

	// Select File Example
	file, err := SelectFile()
	if err != nil {
		fmt.Println("File selection error:", err)
	} else {
		fmt.Println("Selected:", file)
	}

	// Select Folder Example
	folder, err := SelectFolder()
	if err != nil {
		fmt.Println("Folder selection error:", err)
	} else {
		fmt.Println("Selected:", folder)
	}
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
