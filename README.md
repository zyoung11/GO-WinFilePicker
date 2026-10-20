# GO-WinFilePicker

[![Go Reference](https://pkg.go.dev/badge/github.com/zyoung11/GO-WinFilePicker.svg)](https://pkg.go.dev/github.com/zyoung11/GO-WinFilePicker)
[![Go Report Card](https://goreportcard.com/badge/github.com/zyoung11/GO-WinFilePicker)](https://goreportcard.com/report/github.com/zyoung11/GO-WinFilePicker)
[![License](https://img.shields.io/github/license/zyoung11/GO-WinFilePicker)](LICENSE)

**GO-WinFilePicker** is a lightweight Go library that provides modern, native file and folder selection dialogs for Windows applications. It talks to the `IFileOpenDialog` COM interface through `syscall`, without any external dependencies.

The dialog taskbar, title bar and Alt-Tab icon is the Windows File Explorer icon instead of the generic application icon, extracted at the icon size your display scale calls for.

> **Windows only.** The package is guarded by build constraints and does not build on other platforms. Dialogs render with per-monitor DPI awareness for the duration of the call, so they stay crisp on scaled displays without touching the process wide DPI settings of the host application.

## Installation

```bash
go get github.com/zyoung11/GO-WinFilePicker
```

Note: the module path contains uppercase letters, which Go tooling handles automatically.

## Quick Start

```go
package main

import (
	"errors"
	"fmt"

	winfilepicker "github.com/zyoung11/GO-WinFilePicker"
)

func main() {
	file, err := winfilepicker.SelectFile("Please select an image", []string{"jpg", "jpeg", "png", "gif"})
	if errors.Is(err, winfilepicker.ErrCancelled) {
		fmt.Println("cancelled")
	} else if err != nil {
		fmt.Println("error:", err)
	} else {
		fmt.Println("selected:", file)
	}
}
```

## API

| Function | Description |
|----------|-------------|
| `SelectFile(title string, extensions []string) (string, error)` | Single file selection |
| `SelectFiles(title string, extensions []string) ([]string, error)` | Multiple file selection |
| `SelectFolder(title string) (string, error)` | Single folder selection |
| `SelectFolders(title string) ([]string, error)` | Multiple folder selection |

`title` is shown in the dialog title bar, pass an empty string for the system default.

`extensions` filters the visible files and accepts three spellings per entry: `"jpg"`, `".jpg"` or `"*.jpg"`. All extensions are combined into a single filter, so a long list stays readable:

```go
images := []string{"jpg", "jpeg", "png", "gif", "bmp", "webp", "*.tif", "*.tiff"}
files, err := winfilepicker.SelectFiles("Please select images", images)
```

Pass `nil` or an empty slice to show all files.

Cancelling a dialog is reported as an error for which `errors.Is(err, winfilepicker.ErrCancelled)` returns `true`, so user cancellation can be distinguished from real failures.

## Full Example

```go
package main

import (
	"errors"
	"fmt"

	winfilepicker "github.com/zyoung11/GO-WinFilePicker"
)

func main() {
	images := []string{"jpg", "jpeg", "png", "gif"}

	file, err := winfilepicker.SelectFile("Please select an image", images)
	if errors.Is(err, winfilepicker.ErrCancelled) {
		fmt.Println("[Single file] Cancelled")
	} else if err != nil {
		fmt.Println("[Single file] Error:", err)
	} else {
		fmt.Println("[Single file] Result:", file)
	}

	files, err := winfilepicker.SelectFiles("Please select multiple images", images)
	if err == nil {
		fmt.Printf("[Multiple files] Results (%d total):\n", len(files))
		for i, f := range files {
			fmt.Printf("  %d: %s\n", i+1, f)
		}
	}

	folder, err := winfilepicker.SelectFolder("Please select a folder")
	if err == nil {
		fmt.Println("[Single folder] Result:", folder)
	}

	folders, err := winfilepicker.SelectFolders("Please select multiple folders")
	if err == nil {
		fmt.Printf("[Multiple folders] Results (%d total):\n", len(folders))
		for i, f := range folders {
			fmt.Printf("  %d: %s\n", i+1, f)
		}
	}
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
