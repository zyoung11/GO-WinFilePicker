package main

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
	IID_IFileDialog      = GUID{0x42F85136, 0xDB7E, 0x439C, [8]byte{0x85, 0xF1, 0xE4, 0x07, 0x5D, 0x13, 0x5F, 0xC8}}
	IID_IShellItem       = GUID{0x43826D1E, 0xE718, 0x42EE, [8]byte{0x85, 0x65, 0x73, 0x74, 0x71, 0x6C, 0x45, 0x52}}
)

type GUID struct {
	Data1 uint32
	Data2 uint16
	Data3 uint16
	Data4 [8]byte
}

type IFileDialogVtbl struct {
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
}

type IFileDialog struct {
	lpVtbl *IFileDialogVtbl
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

const (
	FOS_PICKFOLDERS          = 0x00000020
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

	end := unsafe.Pointer(p)
	n := 0
	for *(*uint16)(end) != 0 {
		end = unsafe.Pointer(uintptr(end) + unsafe.Sizeof(*p))
		n++
	}

	s := make([]uint16, n)
	ptr := unsafe.Pointer(p)
	for i := 0; i < n; i++ {
		s[i] = *(*uint16)(ptr)
		ptr = unsafe.Pointer(uintptr(ptr) + unsafe.Sizeof(*p))
	}

	return syscall.UTF16ToString(s)
}

func CoTaskMemFree(p unsafe.Pointer) {
	syscall.SyscallN(coTaskMemFree.Addr(), uintptr(p))
}

func FAILED(hr uintptr) bool {
	return int32(hr) < 0
}

func SelectFile() (string, error) {
	hr, _, _ := syscall.SyscallN(coInitializeEx.Addr(), 0, COINIT_APARTMENTTHREADED)
	if FAILED(hr) {
		return "", fmt.Errorf("COM initialization failed")
	}
	defer syscall.SyscallN(coUninitialize.Addr())

	var pFileDialog *IFileDialog
	hr, _, _ = syscall.SyscallN(
		coCreateInstance.Addr(),
		uintptr(unsafe.Pointer(&CLSID_FileOpenDialog)),
		0,
		uintptr(1),
		uintptr(unsafe.Pointer(&IID_IFileDialog)),
		uintptr(unsafe.Pointer(&pFileDialog)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to create FileOpenDialog")
	}
	defer pFileDialog.Release()

	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.SetOptions,
		uintptr(unsafe.Pointer(pFileDialog)),
		0,
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to set options")
	}

	title := "Select File"
	titlePtr, err := syscall.UTF16PtrFromString(title)
	if err != nil {
		return "", fmt.Errorf("failed to convert title: %v", err)
	}

	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.SetTitle,
		uintptr(unsafe.Pointer(pFileDialog)),
		uintptr(unsafe.Pointer(titlePtr)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to set title")
	}

	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.Show,
		uintptr(unsafe.Pointer(pFileDialog)),
		0,
	)
	if FAILED(hr) {
		return "", fmt.Errorf("user deselected")
	}

	var pItem *IShellItem
	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.GetResult,
		uintptr(unsafe.Pointer(pFileDialog)),
		uintptr(unsafe.Pointer(&pItem)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to retrieve results")
	}
	defer pItem.Release()

	var pszPath *uint16
	hr, _, _ = syscall.SyscallN(
		pItem.lpVtbl.GetDisplayName,
		uintptr(unsafe.Pointer(pItem)),
		SIGDN_FILESYSPATH,
		uintptr(unsafe.Pointer(&pszPath)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to retrieve the path")
	}
	defer CoTaskMemFree(unsafe.Pointer(pszPath))

	return UTF16PtrToString(pszPath), nil
}

func SelectFolder() (string, error) {
	hr, _, _ := syscall.SyscallN(coInitializeEx.Addr(), 0, COINIT_APARTMENTTHREADED)
	if FAILED(hr) {
		return "", fmt.Errorf("COM initialization failed")
	}
	defer syscall.SyscallN(coUninitialize.Addr())

	var pFileDialog *IFileDialog
	hr, _, _ = syscall.SyscallN(
		coCreateInstance.Addr(),
		uintptr(unsafe.Pointer(&CLSID_FileOpenDialog)),
		0,
		uintptr(1),
		uintptr(unsafe.Pointer(&IID_IFileDialog)),
		uintptr(unsafe.Pointer(&pFileDialog)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to create FileOpenDialog")
	}
	defer pFileDialog.Release()

	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.SetOptions,
		uintptr(unsafe.Pointer(pFileDialog)),
		FOS_PICKFOLDERS,
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to set options")
	}

	title := "Select folder"
	titlePtr, err := syscall.UTF16PtrFromString(title)
	if err != nil {
		return "", fmt.Errorf("failed to convert title:%v", err)
	}

	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.SetTitle,
		uintptr(unsafe.Pointer(pFileDialog)),
		uintptr(unsafe.Pointer(titlePtr)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to set title")
	}

	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.Show,
		uintptr(unsafe.Pointer(pFileDialog)),
		0,
	)
	if FAILED(hr) {
		return "", fmt.Errorf("user deselected")
	}

	var pItem *IShellItem
	hr, _, _ = syscall.SyscallN(
		pFileDialog.lpVtbl.GetResult,
		uintptr(unsafe.Pointer(pFileDialog)),
		uintptr(unsafe.Pointer(&pItem)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to retrieve results")
	}
	defer pItem.Release()

	var pszPath *uint16
	hr, _, _ = syscall.SyscallN(
		pItem.lpVtbl.GetDisplayName,
		uintptr(unsafe.Pointer(pItem)),
		SIGDN_FILESYSPATH,
		uintptr(unsafe.Pointer(&pszPath)),
	)
	if FAILED(hr) {
		return "", fmt.Errorf("failed to retrieve the path")
	}
	defer CoTaskMemFree(unsafe.Pointer(pszPath))

	return UTF16PtrToString(pszPath), nil
}

func (obj *IFileDialog) Release() uint32 {
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
