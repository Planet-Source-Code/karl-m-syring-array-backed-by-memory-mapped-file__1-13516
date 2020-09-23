<div align="center">

## Array backed by memory mapped file


</div>

### Description

The memory for an array is allocated from a memory mapped file. This is an big advantage for huge arrays, as they will not fill the pagefile.
 
### More Info
 
The size of allocated memory is limited by the free virtual address space (1GB max on 9x, 2GB max on NTx)

Does IO to a file without using VB IO-functions


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Karl M\. Syring](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/karl-m-syring.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/karl-m-syring-array-backed-by-memory-mapped-file__1-13516/archive/master.zip)

### API Declarations

```
' Allocate a 2D array from a memory mapped file
' Autor: Karl M. Syring (syring@email.com)
' inspired by
' http://www.vbaccelerator.com/codelib/gfx/dibsect.htm
' http://www.vb2themax.com/HtmlDoc.asp?Table=Books&ID=1501&Page=3
' http://www.devx.com/premier/mgznarch/vbpj/2000/07jul00/bb0007/bb0007.asp
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
 lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
' Memory mapping API-function
Private Const GENERIC_READ = &H80000000, GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2, OPEN_ALWAYS = 4, FILE_ATTRIBUTE_NORMAL = &H80
Private Const PAGE_READWRITE = 4, FILE_MAP_ALL_ACCESS = &HF001F
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
```


### Source Code

```
Private Type SAFEARRAYBOUND
 cElements As Long
 lLbound As Long
End Type
Private Type SAFEARRAY2D
 cDims As Integer
 fFeatures As Integer
 cbElements As Long
 cLocks As Long
 pvData As Long
 Bounds(0 To 1) As SAFEARRAYBOUND
End Type
' keep it safe, be global
Dim mArray() As Double
Dim tSA As SAFEARRAY2D
Dim hFile As Long
Dim hFileMapping As Long
Dim lpFileBase As Long
Sub Create2DMMArray(Filename As String, ElemSize As Long, n As Long, m As Long)
 With tSA
 .cbElements = ElemSize
 .cDims = 2
 .Bounds(0).lLbound = 0
 .Bounds(0).cElements = m
 .Bounds(1).lLbound = 0
 .Bounds(1).cElements = n
 .fFeatures = &H10 Or &H2 ' FADF_FIXEDSIZE and FADF_STATIC
 .cLocks = 1
 GetViewOfFile Filename, ElemSize, n, m
 .pvData = lpFileBase
 End With
 If tSA.pvData = 0 Then
 Err.Raise 1243, "Create2DMMArray()", "Memory mapping failed"
 Else
 CopyMemory ByVal VarPtrArray(mArray()), VarPtr(tSA), 4
 End If
End Sub
Function GetViewOfFile(Filename As String, ElemSize As Long, n As Long, m As Long) As Long
 hFile = CreateFile(Filename, GENERIC_READ Or GENERIC_WRITE, 0, 0, _
    CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, vbEmpty)
 If hFile = -1 Then Err.Raise Err.LastDllError, "GetViewOfFile()", "Could not open file " & Filename
 Dim FileSize As Long
 FileSize = ElemSize * m * n
 hFileMapping = CreateFileMapping(hFile, 0, PAGE_READWRITE, 0, FileSize, vbEmpty)
 lpFileBase = MapViewOfFile(hFileMapping, FILE_MAP_ALL_ACCESS, 0, 0, 0 * FileSize)
 GetViewOfFile = lpFileBase
End Function
Function FreeViewOfFile() As Long
Dim ret As Long
 ' Clear the temporary array descriptor
 ' This may be necessary under NT4.
 CopyMemory ByVal VarPtrArray(mArray), 0&, 4
 FreeViewOfFile = UnmapViewOfFile(lpFileBase)
 If FreeViewOfFile = 0 Then Debug.Print "Error: ", Err.LastDllError
' If FreeViewOfFile = 0 Then Err.Raise Err.LastDllError, "FreeViewOfFile()", "Memory unmapping failed"
 ret = CloseHandle(hFileMapping)
 ret = CloseHandle(hFile)
End Function
Function checkMMA()
Dim n As Long, m As Long, i As Long, j As Long
Dim Filename As String, ElemSize As Long
 Filename = "c:\kill.me"
 n = 10 ^ 6: m = 10
 ElemSize = 8 ' size of Double is 8
 'Create 2D Array(m,n) of Double,
 Create2DMMArray Filename, ElemSize, n, m
 'random acess to our file
 For i = 0 To 1000
 mArray(Rnd * n Mod n, Rnd * m Mod m) = i
 Next i
' close down, destroy array
' this MUST be called
FreeViewOfFile
End Function
```

