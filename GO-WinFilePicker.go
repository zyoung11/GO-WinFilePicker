package winfilepicker

import (
	"fmt"
	"syscall"
	"unsafe"
)

var (
	user32             = syscall.NewLazyDLL("user32.dll")
	ole32              = syscall.NewLazyDLL("ole32.dll")
	setProcessDPIAware = user32.NewProc("SetProcessDPIAware")
	coInitializeEx     = ole32.NewProc("CoInitializeEx")
	coUninitialize     = ole32.NewProc("CoUninitialize")
	coCreateInstance   = ole32.NewProc("CoCreateInstance")
	coTaskMemFree      = ole32.NewProc("CoTaskMemFree")
)

var (
	CLSID_FileOpenDialog = GUID{0xDC1C5A9C, 0xE88A, 0x4DDE, [8]byte{0xA5, 0xA1, 0x60, 0xF8, 0x2A, 0x20, 0xAE, 0xF7}}
	IID_IFileOpenDialog  = GUID{0xD57C7288, 0xD4AD, 0x4768, [8]byte{0xBE, 0x02, 0x9D, 0x96, 0x95, 0x32, 0xD9, 0x60}}
	IID_IShellItem       = GUID{0x43826D1E, 0xE718, 0x42EE, [8]byte{0x85, 0x65, 0x73, 0x74, 0x71, 0x6C, 0x45, 0x52}}
)

type GUID struct {
	Data1 uint32
	Data2 uint16
	Data3 uint16
	Data4 [8]byte
}

type COMDLG_FILTERSPEC struct {
	pszName *uint16
	pszSpec *uint16
}

type IFileOpenDialogVtbl struct {
	QueryInterface      uintptr
	AddRef              uintptr
	Release             uintptr
	Show                uintptr
	SetFileTypes        uintptr
	SetFileTypeIndex    uintptr
	GetFileTypeIndex    uintptr
	Advise              uintptr
	Unadvise            uintptr
	SetOptions          uintptr
	GetOptions          uintptr
	SetDefaultFolder    uintptr
	SetFolder           uintptr
	GetFolder           uintptr
	GetCurrentSelection uintptr
	SetFileName         uintptr
	GetFileName         uintptr
	SetTitle            uintptr
	SetOkButtonLabel    uintptr
	SetFileNameLabel    uintptr
	GetResult           uintptr
	AddPlace            uintptr
	SetDefaultExtension uintptr
	Close               uintptr
	SetClientGuid       uintptr
	ClearClientData     uintptr
	SetFilter           uintptr
	GetResults          uintptr
	GetSelectedItems    uintptr
}

type IFileOpenDialog struct {
	lpVtbl *IFileOpenDialogVtbl
}

type IShellItemVtbl struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	BindToHandler  uintptr
	GetParent      uintptr
	GetDisplayName uintptr
	GetAttributes  uintptr
	Compare        uintptr
}

type IShellItem struct {
	lpVtbl *IShellItemVtbl
}

type IShellItemArrayVtbl struct {
	QueryInterface             uintptr
	AddRef                     uintptr
	Release                    uintptr
	BindToHandler              uintptr
	GetPropertyStore           uintptr
	GetPropertyDescriptionList uintptr
	GetAttributes              uintptr
	GetCount                   uintptr
	GetItemAt                  uintptr
	EnumItems                  uintptr
}

type IShellItemArray struct {
	lpVtbl *IShellItemArrayVtbl
}

const (
	FOS_PICKFOLDERS          = 0x00000020
	FOS_ALLOWMULTISELECT     = 0x00000200
	FOS_FILEMUSTEXIST        = 0x00001000
	COINIT_APARTMENTTHREADED = 0x2
	SIGDN_FILESYSPATH        = 0x80028000
)

func init() {
	setProcessDPIAware.Call()
}

func UTF16PtrToString(p *uint16) string {
	if p == nil {
		return ""
	}
	length := 0
	for ptr := unsafe.Pointer(p); *(*uint16)(ptr) != 0; ptr = unsafe.Pointer(uintptr(ptr) + 2) {
		length++
	}
	if length == 0 {
		return ""
	}
	s := make([]uint16, length)
	for i := 0; i < length; i++ {
		s[i] = *(*uint16)(unsafe.Pointer(uintptr(unsafe.Pointer(p)) + uintptr(i)*2))
	}
	return syscall.UTF16ToString(s)
}

func CoTaskMemFree(p unsafe.Pointer) {
	syscall.SyscallN(coTaskMemFree.Addr(), uintptr(p))
}

func FAILED(hr uintptr) bool {
	return int32(hr) < 0
}

func SelectFile(title string, extensions ...string) (string, error) {
	paths, err := showOpenDialog(title, false, false, extensions)
	if err != nil {
		return "", err
	}
	if len(paths) > 0 {
		return paths[0], nil
	}
	return "", fmt.Errorf("no file selected")
}

func SelectFiles(title string, extensions ...string) ([]string, error) {
	return showOpenDialog(title, true, false, extensions)
}

func SelectFolder(title string) (string, error) {
	paths, err := showOpenDialog(title, false, true, nil)
	if err != nil {
		return "", err
	}
	if len(paths) > 0 {
		return paths[0], nil
	}
	return "", fmt.Errorf("no folder selected")
}

func SelectFolders(title string) ([]string, error) {
	return showOpenDialog(title, true, true, nil)
}

func showOpenDialog(title string, multiSelect bool, pickFolders bool, extensions []string) ([]string, error) {
	hr, _, _ := syscall.SyscallN(coInitializeEx.Addr(), 0, COINIT_APARTMENTTHREADED)
	if FAILED(hr) {
		return nil, fmt.Errorf("COM initialization failed")
	}
	defer syscall.SyscallN(coUninitialize.Addr())

	var pDialog *IFileOpenDialog
	hr, _, _ = syscall.SyscallN(
		coCreateInstance.Addr(),
		uintptr(unsafe.Pointer(&CLSID_FileOpenDialog)),
		0,
		uintptr(1),
		uintptr(unsafe.Pointer(&IID_IFileOpenDialog)),
		uintptr(unsafe.Pointer(&pDialog)),
	)
	if FAILED(hr) {
		return nil, fmt.Errorf("failed to create FileOpenDialog")
	}
	defer pDialog.Release()

	var options uint32
	syscall.SyscallN(
		pDialog.lpVtbl.GetOptions,
		uintptr(unsafe.Pointer(pDialog)),
		uintptr(unsafe.Pointer(&options)),
	)

	options |= FOS_FILEMUSTEXIST
	if pickFolders {
		options |= FOS_PICKFOLDERS
	}
	if multiSelect {
		options |= FOS_ALLOWMULTISELECT
	}

	hr, _, _ = syscall.SyscallN(
		pDialog.lpVtbl.SetOptions,
		uintptr(unsafe.Pointer(pDialog)),
		uintptr(options),
	)
	if FAILED(hr) {
		return nil, fmt.Errorf("failed to set options")
	}

	if title != "" {
		titlePtr, _ := syscall.UTF16PtrFromString(title)
		syscall.SyscallN(
			pDialog.lpVtbl.SetTitle,
			uintptr(unsafe.Pointer(pDialog)),
			uintptr(unsafe.Pointer(titlePtr)),
		)
	}

	if len(extensions) > 0 && !pickFolders {
		filterSpec := buildFilterSpec(extensions)
		if len(filterSpec) > 0 {
			syscall.SyscallN(
				pDialog.lpVtbl.SetFileTypes,
				uintptr(unsafe.Pointer(pDialog)),
				uintptr(len(filterSpec)),
				uintptr(unsafe.Pointer(&filterSpec[0])),
			)
		}
	}

	hr, _, _ = syscall.SyscallN(
		pDialog.lpVtbl.Show,
		uintptr(unsafe.Pointer(pDialog)),
		0,
	)
	if FAILED(hr) {
		return nil, fmt.Errorf("user cancelled")
	}

	if multiSelect {
		var pItemsArray *IShellItemArray
		hr, _, _ = syscall.SyscallN(
			pDialog.lpVtbl.GetResults,
			uintptr(unsafe.Pointer(pDialog)),
			uintptr(unsafe.Pointer(&pItemsArray)),
		)
		if FAILED(hr) {
			return nil, fmt.Errorf("failed to get results array")
		}
		defer pItemsArray.Release()

		var count uint32
		hr, _, _ = syscall.SyscallN(
			pItemsArray.lpVtbl.GetCount,
			uintptr(unsafe.Pointer(pItemsArray)),
			uintptr(unsafe.Pointer(&count)),
		)
		if FAILED(hr) {
			return nil, fmt.Errorf("failed to get item count")
		}

		var results []string
		for i := uint32(0); i < count; i++ {
			var pItem *IShellItem
			hr, _, _ = syscall.SyscallN(
				pItemsArray.lpVtbl.GetItemAt,
				uintptr(unsafe.Pointer(pItemsArray)),
				uintptr(i),
				uintptr(unsafe.Pointer(&pItem)),
			)
			if FAILED(hr) {
				continue
			}
			path := getPathFromItem(pItem)
			if path != "" {
				results = append(results, path)
			}
			pItem.Release()
		}
		return results, nil
	} else {
		var pItem *IShellItem
		hr, _, _ = syscall.SyscallN(
			pDialog.lpVtbl.GetResult,
			uintptr(unsafe.Pointer(pDialog)),
			uintptr(unsafe.Pointer(&pItem)),
		)
		if FAILED(hr) {
			return nil, fmt.Errorf("failed to retrieve result")
		}
		defer pItem.Release()

		path := getPathFromItem(pItem)
		return []string{path}, nil
	}
}

func buildFilterSpec(extensions []string) []COMDLG_FILTERSPEC {
	combinedName := "Supported Files"
	var combinedSpec string
	for i, ext := range extensions {
		if i > 0 {
			combinedSpec += ";"
		}
		combinedSpec += "*." + ext
	}

	namePtr, _ := syscall.UTF16PtrFromString(combinedName)
	specPtr, _ := syscall.UTF16PtrFromString(combinedSpec)

	return []COMDLG_FILTERSPEC{
		{pszName: namePtr, pszSpec: specPtr},
	}
}

func getPathFromItem(pItem *IShellItem) string {
	var pszPath *uint16
	hr, _, _ := syscall.SyscallN(
		pItem.lpVtbl.GetDisplayName,
		uintptr(unsafe.Pointer(pItem)),
		SIGDN_FILESYSPATH,
		uintptr(unsafe.Pointer(&pszPath)),
	)
	if FAILED(hr) {
		return ""
	}
	defer CoTaskMemFree(unsafe.Pointer(pszPath))
	return UTF16PtrToString(pszPath)
}

func (obj *IFileOpenDialog) Release() uint32 {
	ret, _, _ := syscall.SyscallN(
		obj.lpVtbl.Release,
		uintptr(unsafe.Pointer(obj)),
	)
	return uint32(ret)
}

func (obj *IShellItem) Release() uint32 {
	ret, _, _ := syscall.SyscallN(
		obj.lpVtbl.Release,
		uintptr(unsafe.Pointer(obj)),
	)
	return uint32(ret)
}

func (obj *IShellItemArray) Release() uint32 {
	ret, _, _ := syscall.SyscallN(
		obj.lpVtbl.Release,
		uintptr(unsafe.Pointer(obj)),
	)
	return uint32(ret)
}

// func main() {

// 	// 1. Single file selection
// 	file, err := SelectFile("Please select an image", "jpg", "png", "gif")
// 	if err != nil {
// 		fmt.Println("[Single file] Cancelled or error:", err)
// 	} else {
// 		fmt.Println("[Single file] Result:", file)
// 	}

// 	// 2. Multiple files selection
// 	files, err := SelectFiles("Please select multiple images", "jpg", "png", "gif")
// 	if err != nil {
// 		fmt.Println("[Multiple files] Cancelled or error:", err)
// 	} else {
// 		fmt.Printf("[Multiple files] Results (%d total):\n", len(files))
// 		for i, f := range files {
// 			fmt.Printf("  %d: %s\n", i+1, f)
// 		}
// 	}

// 	// 3. Single folder selection
// 	folder, err := SelectFolder("Please select a folder")
// 	if err != nil {
// 		fmt.Println("[Single folder] Cancelled or error:", err)
// 	} else {
// 		fmt.Println("[Single folder] Result:", folder)
// 	}

// 	// 4. Multiple folders selection
// 	folders, err := SelectFolders("Please select multiple folders")
// 	if err != nil {
// 		fmt.Println("[Multiple folders] Cancelled or error:", err)
// 	} else {
// 		fmt.Printf("[Multiple folders] Results a(%d total):\n", len(folders))
// 		for i, f := range folders {
// 			fmt.Printf("  %d: %s\n", i+1, f)
// 		}
// 	}
// }
