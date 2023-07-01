Attribute VB_Name = "Module1"
Option Explicit

' Required data structures
Private Type POINTAPI
    x As Long
    y As Long
End Type

' Clipboard Manager Functions
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' Other required Win32 APIs
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Predefined Clipboard Formats
Private Const CF_TEXT         As Long = 1
Private Const CF_BITMAP       As Long = 2
Private Const CF_METAFILEPICT As Long = 3
Private Const CF_SYLK         As Long = 4
Private Const CF_DIF          As Long = 5
Private Const CF_TIFF         As Long = 6
Private Const CF_OEMTEXT      As Long = 7
Private Const CF_DIB          As Long = 8
Private Const CF_PALETTE      As Long = 9
Private Const CF_PENDATA      As Long = 10
Private Const CF_RIFF         As Long = 11
Private Const CF_WAVE         As Long = 12
Private Const CF_UNICODETEXT  As Long = 13
Private Const CF_ENHMETAFILE  As Long = 14
Private Const CF_HDROP        As Long = 15
Private Const CF_LOCALE       As Long = 16
Private Const CF_MAX          As Long = 17

' New shell-oriented clipboard formats
Private Const CFSTR_SHELLIDLIST       As String = "Shell IDList Array"
Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
Private Const CFSTR_NETRESOURCES      As String = "Net Resource"
Private Const CFSTR_FILEDESCRIPTOR    As String = "FileGroupDescriptor"
Private Const CFSTR_FILECONTENTS      As String = "FileContents"
Private Const CFSTR_FILENAME          As String = "FileName"
Private Const CFSTR_PRINTERGROUP      As String = "PrinterFriendlyName"
Private Const CFSTR_FILENAMEMAP       As String = "FileNameMap"

' Global Memory Flags
Private Const GMEM_FIXED          As Long = &H0
Private Const GMEM_MOVEABLE       As Long = &H2
Private Const GMEM_NOCOMPACT      As Long = &H10
Private Const GMEM_NODISCARD      As Long = &H20
Private Const GMEM_ZEROINIT       As Long = &H40
Private Const GMEM_MODIFY         As Long = &H80
Private Const GMEM_DISCARDABLE    As Long = &H100
Private Const GMEM_NOT_BANKED     As Long = &H1000
Private Const GMEM_SHARE          As Long = &H2000
Private Const GMEM_DDESHARE       As Long = &H2000
Private Const GMEM_NOTIFY         As Long = &H4000
Private Const GMEM_LOWER          As Long = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS    As Long = &H7F72
Private Const GMEM_INVALID_HANDLE As Long = &H8000
Private Const GHND                As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR                As Long = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Type DROPFILES
    pFiles As Long
    pt     As POINTAPI
    fNC    As Long
    fWide  As Long
End Type

Public Function ClipboardCopyFiles(Files() As String) As Boolean

    Dim data As String
    Dim df As DROPFILES
    Dim hGlobal As Long
    Dim lpGlobal As Long
    Dim i As Long

    ' Open and clear existing crud off clipboard.
    If OpenClipboard(0&) Then
        Call EmptyClipboard

        ' Build double-null terminated list of files.
        For i = LBound(Files) To UBound(Files)
            data = data & Files(i) & vbNullChar
        Next
        data = data & vbNullChar

        ' Allocate and get pointer to global memory,
        ' then copy file list to it.
        hGlobal = GlobalAlloc(GHND, Len(df) + Len(data))
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)

            ' Build DROPFILES structure in global memory.
            df.pFiles = Len(df)
            Call CopyMem(ByVal lpGlobal, df, Len(df))
            Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal data, Len(data))
            Call GlobalUnlock(hGlobal)

            ' Copy data to clipboard, and return success.
            If SetClipboardData(CF_HDROP, hGlobal) Then
                ClipboardCopyFiles = True
            End If
        End If

        ' Clean up
        Call CloseClipboard
    End If

End Function

Public Function ClipboardPasteFiles(Files() As String) As Long

    Dim hDrop As Long
    Dim nFiles As Long
    Dim i As Long
    Dim desc As String
    Dim filename As String
    Dim pt As POINTAPI
    Const MAX_PATH As Long = 260

    ' Insure desired format is there, and open clipboard.
    If IsClipboardFormatAvailable(CF_HDROP) Then
        If OpenClipboard(0&) Then

            ' Get handle to Dropped Filelist data, and number of files.
            hDrop = GetClipboardData(CF_HDROP)
            nFiles = DragQueryFile(hDrop, -1&, "", 0)

            ' Allocate space for return and working variables.
            ReDim Files(0 To nFiles - 1) As String
            filename = Space(MAX_PATH)

            ' Retrieve each filename in Dropped Filelist.
            For i = 0 To nFiles - 1
                Call DragQueryFile(hDrop, i, filename, Len(filename))
                Files(i) = TrimNull(filename)
            Next

            ' Clean up
            Call CloseClipboard
        End If

        ' Assign return value equal to number of files dropped.
        ClipboardPasteFiles = nFiles
    End If

End Function

Private Function TrimNull(ByVal sTmp As String) As String

    Dim nNul As Long

    '
    ' Truncate input sTmpg at first Null.
    ' If no Nulls, perform ordinary Trim.
    '
    nNul = InStr(sTmp, vbNullChar)
    Select Case nNul
    Case Is > 1
        TrimNull = Left(sTmp, nNul - 1)
    Case 1
        TrimNull = ""
    Case 0
        TrimNull = Trim(sTmp)
    End Select

End Function
