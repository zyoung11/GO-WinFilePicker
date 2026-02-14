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

	// 1. Single file selection
	file, err := SelectFile("Please select an image", "jpg", "png", "gif")
	if err != nil {
		fmt.Println("[Single file] Cancelled or error:", err)
	} else {
		fmt.Println("[Single file] Result:", file)
	}

	// 2. Multiple file selection
	files, err := SelectFiles("Please select multiple images", "jpg", "png", "gif")
	if err != nil {
		fmt.Println("[Multiple files] Cancelled or error:", err)
	} else {
		fmt.Printf("[Multiple files] Results (%d total):\n", len(files))
		for i, f := range files {
			fmt.Printf("  %d: %s\n", i+1, f)
		}
	}

	// 3. Single folder selection
	folder, err := SelectFolder("Please select a folder")
	if err != nil {
		fmt.Println("[Single folder] Cancelled or error:", err)
	} else {
		fmt.Println("[Single folder] Result:", folder)
	}

	// 4. Multiple folder selection
	folders, err := SelectFolders("Please select multiple folders")
	if err != nil {
		fmt.Println("[Multiple folders] Cancelled or error:", err)
	} else {
		fmt.Printf("[Multiple folders] Results (%d total):\n", len(folders))
		for i, f := range folders {
			fmt.Printf("  %d: %s\n", i+1, f)
		}
	}
}

```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
