VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Global Memory Flags
Private Const GMEM_FIXED          As Long = &H0
Private Const GMEM_MOVEABLE       As Long = &H2&
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

Public Enum ClipboardFormat
                           ' Das Objekt in der Zwischenablage ist ein . . .
    CF_TEXT = 1                     ' Handle auf einen unformatierten Ansi-Text (UTF-8).
    CF_BITMAP = 2                   ' Handle auf eine Windows Bitmap-Grafik
    CF_METAFILEPICT = 3             ' Handle auf eine Windows Metafile-Grafik
    CF_SYLK = 4                     ' Handle auf einen Microsoft Symbolic-Link
    CF_DIF = 5                      ' "Software Arts' Data Interchange Format"
    CF_TIFF = 6                     ' Handle zu einem Tiff-Bitmap
    CF_OEMTEXT = 7                  ' Handle zu einem OEM-String
    CF_DIB = 8                      ' Handle zu einer %BITMAPINFO%-Struktur
    CF_BOTTOMUP_DIB = CF_DIB        '
    CF_PALETTE = 9                  ' Handle zu einer Palette
    CF_PENDATA = 10                 ' sind Daten zu einem Microsoft Pen Extensions
    CF_RIFF = 11                    ' Handle zu einer Audiodatei
    CF_WAVE = 12                    ' Handle zu Wavedatei
    CF_UNICODETEXT = 13             ' Handle zu einem Unicode-String (UTF-16)
    CF_ENHMETAFILE = 14             ' Handle zu einer Enhanced Metadatei
    CF_HDROP = 15                   ' Liste von Dateihandles im Zusammenhang mit Drag-And-Drop
    CF_LOCALE = 16                  ' Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde
    CF_DIBV5 = 17                   ' Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)
    
    CF_JPEG = 19
    CF_TOPDOWN_DIB = 20
    
    CF_MULTI_TIFF = 22
    
    CF_OWNERDISPLAY = &H80&         '   128 ' benutzerdefinierter Anzeigetyp
    CF_DSPTEXT = &H81&              '   129 ' Text, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_DSPBITMAP = &H82&            '   130 ' Bitmap, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_DSPMETAFILEPICT = &H83&      '   131 ' Zwischendatei, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_PRIVATEFIRST = &H200&        '   512 ' privates Handle
    CF_PRIVATELAST = &H2FF&         '   767 ' privates Handle
    CF_GDIOBJFIRST = &H300&         '   768 ' Dient in der Zwischenablage dem Setzen von privaten Datenformate
    CF_GDIOBJLAST = &H3FF&          '  1023 ' Dient in der Zwischenablage dem Setzen von privaten Datenformaten
    CF_FileName = &HC006&           ' 49158 ' Dateiname
    CF_FileNameW = &HC007&          ' 49159 ' Dateiname
    CF_DataObject = &HC009&         ' 49161 ' DataObject
    'von mir selber hinzugefügt, herausgefunden durch Excel-Zelle in Zwischenablage:
    CF_RTF = &HC09A&                ' 49306 ' Richt Text Format
    
    CF_HTML = &HC0E2&
    
    CF_HTML_xls1 = &HC108&               ' 49416 ' HTML Format
    CF_HTML_xls2 = &HC12F&               ' 49455 ' HTML Format
    'da will mich bei Microsoft wohl jemand ärgern!!
    'jetzt hat sich plötzlich die Konstante geändert schon sehr merkwürdig
    'man braucht eine Funktion die zum Text HTML Format die Konstante raussucht!
    
    CF_PICTURE = &HC20A&            ' 49674 ' Handle auf ein Objekt vom Datentyp Picture
    CF_OBJECT = &HC215&             ' 49685 ' Handle auf ein beliebiges Objekt
    CF_XML = &HC308&                ' 49928 ' XML Spreadheet
' da gibt es noch mehr Konstanten...
'Shell IDList Array
'DataObjectAttributes
'DataObjectAttributesRequiringElevation
'Shell Object Offsets
'Preferred DropEffect
'AsyncFlag
'FileDrop
'FileNameW
'FileName
End Enum

' New shell-oriented clipboard formats
Private Const CFSTR_FILECONTENTS      As String = "FileContents"
Private Const CFSTR_FILEDESCRIPTOR    As String = "FileGroupDescriptor"
Private Const CFSTR_FILENAME          As String = "FileName"
Private Const CFSTR_FILENAMEMAP       As String = "FileNameMap"
Private Const CFSTR_NETRESOURCES      As String = "Net Resource"
Private Const CFSTR_PRINTERGROUP      As String = "PrinterFriendlyName"
Private Const CFSTR_SHELLIDLIST       As String = "Shell IDList Array"
Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"

' Required data structures
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type DROPFILES
    pFiles As Long
    pt     As POINTAPI
    fNC    As Long
    fWide  As Long
End Type

Private Declare Function OpenClipboard Lib _
    "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumClipboardFormats Lib _
    "user32" (ByVal wFormat As Long) As Long
Private Declare Function CountClipboardFormats Lib _
    "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib _
    "user32" (ByVal wFormat As Long) As Long

Private Declare Function GetClipboardData Lib _
    "user32" (ByVal wFormat As Long) As Long

Private Declare Function GetClipboardDataObj Lib _
    "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Object

Private Declare Function SetClipboardData Lib _
    "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardFormatNameA Lib _
    "user32" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib _
    "user32" () As Long
Private Declare Function CloseClipboard Lib _
    "user32" () As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDst As Long, ByVal pSrc As Long, ByVal bLen As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
                    
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" ( _
    ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

'https://msdn.microsoft.com/de-de/library/windows/desktop/aa366574(v=vs.85).aspx
'GlobalAlloc: Allocates the specified number of bytes from the heap.
'Note:
'The global functions have greater overhead and provide fewer features than other memory management functions.
'New applications should use the heap functions unless documentation states that a global function should be used.
'For more information, see Global and Local Functions.
'
'Global and Local Functions:
'https://msdn.microsoft.com/de-de/library/windows/desktop/aa366596(v=vs.85).aspx
'For example, some Windows functions allocate memory that must be freed with LocalFree, and the global
'functions are still used with Dynamic Data Exchange (DDE), the clipboard functions, and OLE data objects.


'Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' BitBlt dwRop-Konstante
Private Const SRCCOPY As Long = &HCC0020

Private m_CBFormats() As Long
Private m_hWnd        As Long

Private Sub Class_Initialize()
    ReDim m_CBFormats(0)
End Sub

Public Sub Clear()
    OpenClipboard m_hWnd
    EmptyClipboard
    CloseClipboard
    ReDim m_CBFormats(0)
End Sub
Public Property Get IsClear() As Boolean
    IsClear = UBound(m_CBFormats) = 0
End Property
Public Sub CBFormats_Clear()
    ReDim m_CBFormats(0)
End Sub

Public Property Get hwnd() As Long
    hwnd = m_hWnd
End Property
Public Property Let hwnd(aHWnd As Long)
    m_hWnd = aHWnd
End Property

Public Sub ReadClipboardFormats()
    If OpenClipboard(m_hWnd) Then
        Dim n As Long: n = CountClipboardFormats
        ReDim m_CBFormats(0 To n)
        Dim cbf As Long
        Dim i As Long
        Do
            i = i + 1
            ' Das nächste Format aus der Zwischenablage ermitteln
            cbf = EnumClipboardFormats(cbf)
            If cbf = 0 Then Exit Do
            m_CBFormats(i) = cbf
        Loop
        CloseClipboard
    End If
End Sub

Public Sub CBFormats_ToListBox(aLB As ListBox)
    Dim cfs() As Long: cfs = CBFormats
    Dim cf As Long
    Dim i As Long
    With aLB
        .Clear
        For i = 0 To UBound(cfs)
            cf = cfs(i)
            If cf <> 0 Then
                .AddItem CBFormat_ToStr(cf)
            End If
        Next
    End With
End Sub

Public Property Get CBFormats() As Long()
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats
    CBFormats = m_CBFormats
End Property

Public Function GetCBFormatForName(aName As String) As Long
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats
    Dim i As Long
    For i = 0 To UBound(m_CBFormats)
        Dim n1 As String: n1 = CBFormat_ToStr(m_CBFormats(i))
        Dim n2 As String: n2 = "*" & aName & "*"
        If n1 Like n2 Then
            GetCBFormatForName = m_CBFormats(i)
            Exit Function
        End If
    Next
End Function

Public Property Get HasFormat(ByVal aCBFormat As ClipboardFormat) As Boolean
    HasFormat = IsClipboardFormatAvailable(aCBFormat) '<> 0
    'Dim i As Long
    'For i = 0 To UBound(m_CBFormats)
    '    HasCBFormat = m_CBFormats(i) = aCBFormat
    '    If HasFormat Then
    '        Exit Property
    '    End If
    'Next
End Property

Public Property Get StrData(ByVal aCBFormat As Long) As String
    If OpenClipboard(m_hWnd) Then
        'Dim pMem As LongPtr:         pMem = GetClipboardData(aCBFormat)
        Dim pMem As Long:         pMem = GetClipboardData(aCBFormat)
        If pMem Then
            Dim bLen As Long:        bLen = GlobalSize(pMem)
            If bLen Then
                'Dim pLok As LongPtr: pLok = GlobalLock(pMem)
                Dim pLok As Long: pLok = GlobalLock(pMem)
                If pLok Then
                    If aCBFormat = CF_LOCALE Then
                        Dim localeID As Long
                        RtlMoveMemory VarPtr(localeID), pLok, bLen
                        StrData = CStr(localeID) 'StrConv(StrData, vbUnicode)
                    Else
                        Dim slen As Long: slen = IIf(aCBFormat = CF_UNICODETEXT, (bLen \ 2) - 1, bLen)
                        StrData = Space$(slen)
                        RtlMoveMemory StrPtr(StrData), pLok, bLen
                        If aCBFormat <> CF_UNICODETEXT Then
                            StrData = StrConv(StrData, vbUnicode)
                        End If
                    End If
                    GlobalUnlock pMem
                End If
            End If
        End If
        CloseClipboard
    End If
End Property

Public Property Let StrData(ByVal aCBFormat As Long, Value As String)
    If OpenClipboard(m_hWnd) Then
        Dim bLen As Long:            bLen = LenB(Value) + 2
        If bLen > 2 Then
            'Dim pMem As LongPtr:     pMem = GlobalAlloc(GMEM_MOVEABLE, bLen)
            Dim pMem As Long:     pMem = GlobalAlloc(GMEM_MOVEABLE, bLen)
            If pMem Then
                'Dim pLok As LongPtr: pLok = GlobalLock(pMem)
                Dim pLok As Long: pLok = GlobalLock(pMem)
                If pLok Then
                    RtlMoveMemory pLok, StrPtr(Value), bLen
                    GlobalUnlock pMem
                    SetClipboardData aCBFormat, pMem
                End If
            End If
        End If
        CloseClipboard
    End If
End Property

Public Property Get Data(ByVal aCBFormat As Long) As Long
'
End Property
Public Property Let Data(ByVal aCBFormat As Long, ByVal HndData As Long)
    SetClipboardData aCBFormat, HndData
End Property

Public Property Get DataObj(ByVal aCBFormat As Long) As Object
    Set DataObj = GetClipboardDataObj(aCBFormat)
End Property

Public Function GetData(Optional ByVal aCBFormat As Long)
    If OpenClipboard(m_hWnd) Then
        If aCBFormat = CF_BITMAP Then
            Set GetData = GetPicture
        End If
        CloseClipboard
    End If
End Function

Public Function GetPicture(Optional cf As Long = CF_BITMAP) As IPictureDisp
    If OpenClipboard(m_hWnd) Then
        Dim hnd As Long: hnd = GetClipboardData(cf)
        Set GetPicture = Picture_FromHandle(hnd, , CBFormat_ToPicType(cf))
        GDIDeleteObject hnd
        CloseClipboard
    End If
End Function

Private Function CBFormat_ToPicType(cf As Long) As PictureTypeConstants
    Dim ptc As PictureTypeConstants
    Select Case cf
    Case CF_BITMAP:          ptc = PictureTypeConstants.vbPicTypeBitmap
    Case CF_METAFILEPICT:    ptc = PictureTypeConstants.vbPicTypeMetafile
    Case CF_ENHMETAFILE:     ptc = PictureTypeConstants.vbPicTypeEMetafile
    Case CF_MULTI_TIFF:      ptc = PictureTypeConstants.vbPicTypeBitmap
    Case CF_JPEG:            ptc = PictureTypeConstants.vbPicTypeBitmap
    Case CF_DIBV5:           ptc = PictureTypeConstants.vbPicTypeBitmap
    Case CF_DSPBITMAP:       ptc = PictureTypeConstants.vbPicTypeBitmap
    Case CF_DSPMETAFILEPICT: ptc = PictureTypeConstants.vbPicTypeBitmap
    Case CF_PICTURE:         ptc = PictureTypeConstants.vbPicTypeBitmap
    Case Else:               ptc = PictureTypeConstants.vbPicTypeNone
    End Select
    CBFormat_ToPicType = ptc
End Function
Public Function CBFormat_ToStr(aCBFormat As Long) As String
    Dim s  As String: s = CLng(aCBFormat) & ", &H" & Hex(aCBFormat)
    Dim s2 As String: s2 = Space(256)
    Dim rv As Long: rv = GetClipboardFormatNameA(aCBFormat, s2, 256)
    s2 = Trim(s2)
    If Len(s2) Then
        s = s & "(api): " & s2
    Else
        Select Case aCBFormat
                    
        Case 0:
        Case CF_TEXT:         s = s & ": Handle zu String"
        
        Case CF_BITMAP:       s = s & ": Handle eines Bitmaps"
        
        Case CF_METAFILEPICT: s = s & ": Handle zu einem Metafile Bild"
        Case CF_SYLK:         s = s & ": CF_SYLK"
        Case CF_DIF:          s = s & ": Software Arts' Data Interchange Format"
        Case CF_TIFF:         s = s & ": Handle zu einem Tiff-Bitmap"
        Case CF_OEMTEXT:      s = s & ": Handle zu einem OEM-String"
        Case CF_DIB:          s = s & ": Handle zu einer %BITMAPINFO%-Struktur"
        Case CF_BOTTOMUP_DIB: s = s & ": CF_DIB"
        Case CF_PALETTE:      s = s & ": Handle zu einer Palette"
        Case CF_PENDATA:      s = s & ": sind Daten zu einem Microsoft Pen Extensions"
        Case CF_RIFF:         s = s & ": Handle zu einer Audiodatei"
        Case CF_WAVE:         s = s & ": Handle zu Wavedatei"
        Case CF_UNICODETEXT:  s = s & ": Handle zu einem Unicode-String"
        Case CF_ENHMETAFILE:  s = s & ": Handle zu einer Enhanced Metadatei"
        Case CF_HDROP:        s = s & ": Liste von Dateihandles"
        Case CF_LOCALE:       s = s & ": Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde"
        Case CF_DIBV5:        s = s & ": Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)"
        Case CF_JPEG:         s = s & ": CF_JPEG"
        Case CF_TOPDOWN_DIB:  s = s & ": CF_TOPDOWN_DIB"
    
        Case CF_MULTI_TIFF:   s = s & ": CF_MULTI_TIFF"
        Case CF_OWNERDISPLAY: s = s & ": benutzerdefinierter Anzeigetyp"
        
        Case CF_PRIVATEFIRST: s = s & ": CF_PRIVATEFIRST privates Handle"
        Case CF_PRIVATELAST:  s = s & ": CF_PRIVATELAST privates Handle"
        Case CF_GDIOBJFIRST:  s = s & ": CF_GDIOBJFIRST"
        Case CF_GDIOBJLAST:   s = s & ": CF_GDIOBJLAST"
    
        Case CF_FileName:     s = s & ": CF_FileName   Dateiname"
        Case CF_FileNameW:    s = s & ": CF_FileNameW  Dateiname"
        Case CF_DataObject:   s = s & ": CF_DataObject DataObject"
            
        Case CF_RTF:          s = s & ": CF_RTF Richt Text Format" ' = &HC09A& = 49306 ' Richt Text Format
        
        Case CF_HTML:         s = s & ": CF_HTML Hypertext Markup Language" ' = &HC0E2&
        '
        Case CF_HTML_xls1:    s = s & ": CF_HTML_xls1 Hypertext Markup Language mit Excel-Tabelle (1)"  ' = &HC108&               ' 49416 ' HTML Format
        Case CF_HTML_xls2:    s = s & ": CF_HTML_xls2 Hypertext Markup Language mit Excel-Tabelle (2)"  ' = &HC12F&               ' 49455 ' HTML Format
        
        Case CF_PICTURE:      s = s & ": CF_PICTURE  Picture" ' = &HC20A&   ' 49674 ' Handle auf ein Objekt vom Datentyp Picture
        Case CF_OBJECT:       s = s & ": CF_OBJECT  Object"   ' = &HC215&   ' 49685 ' Handle auf ein beliebiges Objekt
        Case CF_XML:          s = s & ": CF_XML  Xml-Daten"   ' = &HC308&   ' 49928 ' XML Spreadheet
        '' da gibt es noch mehr Konstanten...
        ''Shell IDList Array
        ''DataObjectAttributes
        ''DataObjectAttributesRequiringElevation
        ''Shell Object Offsets
        ''Preferred DropEffect
        ''AsyncFlag
        ''FileDrop
        
        Case Else:            s = s & ": unbekanntes Format"
        End Select
    End If
    CBFormat_ToStr = s
End Function

Public Function CopyFiles(Files() As String) As Boolean
    ' Open and clear existing crud off clipboard.
    If OpenClipboard(m_hWnd) Then
        Call EmptyClipboard
        ' Build double-null terminated list of files.
        Dim Data As String
        Dim i    As Long
        For i = LBound(Files) To UBound(Files)
            Data = Data & Files(i) & vbNullChar
        Next
        Data = Data & vbNullChar
        ' Allocate and get pointer to global memory,
        ' then copy file list to it.
        Dim df      As DROPFILES
        Dim hGlobal As Long: hGlobal = GlobalAlloc(GHND, Len(df) + Len(Data))
        If hGlobal Then
            Dim lpGlobal As Long:  lpGlobal = GlobalLock(hGlobal)
            ' Build DROPFILES structure in global memory.
            df.pFiles = Len(df)
            CopyMem ByVal lpGlobal, df, Len(df)
            CopyMem ByVal (lpGlobal + Len(df)), ByVal Data, Len(Data)
            Call GlobalUnlock(hGlobal)
            ' Copy data to clipboard, and return success.
            If SetClipboardData(CF_HDROP, hGlobal) Then
                CopyFiles = True
            End If
        End If
        ' Clean up
        Call CloseClipboard
    End If
End Function

'Public Function ClipboardCopyFiles(Files() As String) As Boolean
'
'    Dim data As String
'    Dim df As DROPFILES
'    Dim hGlobal As Long
'    Dim lpGlobal As Long
'    Dim i As Long
'    ' Open and clear existing crud off clipboard.
'    If OpenClipboard(0&) Then
'        Call EmptyClipboard
'        ' Build double-null terminated list of files.
'        For i = LBound(Files) To UBound(Files)
'            data = data & Files(i) & vbNullChar
'        Next
'        data = data & vbNullChar
'        ' Allocate and get pointer to global memory,
'        ' then copy file list to it.
'        hGlobal = GlobalAlloc(GHND, Len(df) + Len(data))
'        If hGlobal Then
'            lpGlobal = GlobalLock(hGlobal)
'            ' Build DROPFILES structure in global memory.
'            df.pFiles = Len(df)
'            Call CopyMem(ByVal lpGlobal, df, Len(df))
'            Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal data, Len(data))
'            Call GlobalUnlock(hGlobal)
'            ' Copy data to clipboard, and return success.
'            If SetClipboardData(CF_HDROP, hGlobal) Then
'                ClipboardCopyFiles = True
'            End If
'        End If
'        ' Clean up
'        Call CloseClipboard
'    End If
'End Function

'Public Function CopyFilesH(FileHs As Collection) As Boolean
'    ' Open and clear existing crud off clipboard.
'    If OpenClipboard(m_hWnd) Then
'        Call EmptyClipboard
'        ' Build double-null terminated list of files.
'        Dim data As String
'        Dim i    As Long
'        For i = LBound(Files) To UBound(Files)
'            data = data & Files(i) & vbNullChar
'        Next
'        data = data & vbNullChar
'        ' Allocate and get pointer to global memory,
'        ' then copy file list to it.
'        Dim df      As DROPFILES
'        Dim hGlobal As Long: hGlobal = GlobalAlloc(GHND, Len(df) + Len(data))
'        If hGlobal Then
'            Dim lpGlobal As Long:  lpGlobal = GlobalLock(hGlobal)
'            ' Build DROPFILES structure in global memory.
'            df.pFiles = Len(df)
'            RtlMoveMemory ByVal lpGlobal, VarPtr(df), Len(df)
'            RtlMoveMemory ByVal (lpGlobal + Len(df)), ByVal data, Len(data)
'            Call GlobalUnlock(hGlobal)
'            ' Copy data to clipboard, and return success.
'            If SetClipboardData(CF_HDROP, hGlobal) Then
'                CopyFiles = True
'            End If
'        End If
'        ' Clean up
'        Call CloseClipboard
'    End If
'End Function

Public Function PasteFiles(Files() As String) As Long
    'Dim desc     As String
    'Dim pt       As POINTAPI
    Const MAX_PATH As Long = 260
    ' Insure desired format is there, and open clipboard.
    If IsClipboardFormatAvailable(CF_HDROP) Then
        If OpenClipboard(m_hWnd) Then
            ' Get handle to Dropped Filelist data, and number of files.
            Dim hDrop    As Long:      hDrop = GetClipboardData(CF_HDROP)
            Dim nFiles   As Long:     nFiles = DragQueryFile(hDrop, -1&, "", 0)
            ' Allocate space for return and working variables.
            ReDim Files(0 To nFiles - 1) As String
            Dim filename As String: filename = Space(MAX_PATH)
            ' Retrieve each filename in Dropped Filelist.
            Dim i As Long
            For i = 0 To nFiles - 1
                Call DragQueryFile(hDrop, i, filename, Len(filename))
                Files(i) = TrimNull(filename)
            Next
            ' Clean up
            Call CloseClipboard
        End If
        ' Assign return value equal to number of files dropped.
        PasteFiles = nFiles
    End If
End Function

Private Function TrimNull(ByVal sTmp As String) As String
    ' Truncate input sTmp at first Null. If no Nulls, perform ordinary Trim.
    Dim nNul As Long: nNul = InStr(sTmp, vbNullChar)
    Select Case nNul
    Case Is > 1: TrimNull = Left(sTmp, nNul - 1)
    Case 1:      TrimNull = ""
    Case 0:      TrimNull = Trim(sTmp)
    End Select
End Function

