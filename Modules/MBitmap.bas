Attribute VB_Name = "MBitmap"
Option Explicit

Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PictDesc
    cbSize   As Long
    PicType  As Long
    hBmp     As Long
    hPal     As Long
    Reserved As Long
End Type

Public Declare Function GDIDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long

Public Declare Sub IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As IID)

Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    PicDesc As PictDesc, ByVal lpiid As Long, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'Private Declare Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
'
'Private Sub cmdCreateIPicture_Click()
'    ' *********************
'    ' You must set a Reference to:
'    ' "OLE Automation"
'    ' for this function to work.
'    ' Goto the Menu and select
'    ' Tools->References
'    ' Scroll down to:
'    ' Ole Automation
'    ' and click in the check box to select
'    ' this reference.
'
'    'Dim lngRet As Long
'    'Dim lngBytes As Long
'
'    'Dim hPicBox As StdPicture
'
'    Me.OLEBound19.SetFocus
'    'Me.OLEbound19.SizeMode = acOLESizeZoom
'    DoCmd.RunCommand acCmdCopy
'    Dim hBitmap As Long:   hBitmap = GetClipBoard
'    Dim hPix As IPicture: Set hPix = BitmapToPicture(hBitmap)
'
'    SavePicture hPix, "C:\ole.bmp"
'    apiDeleteObject (hBitmap)
'    Me.Image0.Picture = "C:\ole.bmp"
'
'    Set hPix = Nothing
'End Sub
'
'' Here's the code behind the code module
'
'
'Private Const vbPicTypeBitmap = 1


'''Windows API Function Declarations

'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long

'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'The API format types we're interested in
'Const CF_BITMAP = 2
'Const CF_PALETTE = 9
'Const CF_ENHMETAFILE = 14
'Const IMAGE_BITMAP = 0
'Const LR_COPYRETURNORG = &H4
'' Addded by SL Apr/2000
'Const xlPicture = CF_BITMAP
'Const xlBitmap = CF_BITMAP

'*******************************************
'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 VBA ONLY
'
'Copyright: Lebans Holdings 1999 Ltd.
' May not be resold in whole or part. Please feel
' free to use any/all of this code within your
' own application without cost or obligation.
' Please include the one line Copyright notice
' if you use this function in your own code.
'
'Name: BitmapToPicture &
' GetClipBoard
'
'Purpose: Provides a method to save the contents of a
' Bound or Unbound OLE Control to a Disk file.
' This version only handles BITMAP files.
' '
'Author: Stephen Lebans
'Email: Step...@lebans.com
'Web Site: www.lebans.com
'Date: Apr 10, 2000, 05:31:18 AM
'
'Called by: Any
'
'Inputs: Needs a Handle to a Bitmap.
' This must be a 24 bit bitmap for this release.
'
'Credits:
'As noted directly in Source :-)
'
'BUGS:
'To keep it simple this version only works with Bitmap files of 16 or 24 bits.
'I'll go back and add the
'code to allow any depth bitmaps and add support for
'metafiles as well.
'No serious bugs notices at this point in time.
'Please report any bugs to my email address.
'
'What's Missing:
'
'
'HOW TO USE:
'
'*******************************************

Public Function Picture_FromHandle(ByVal hBmp As Long, Optional ByVal hPal As Long = 0&, Optional aPicType As PictureTypeConstants = PictureTypeConstants.vbPicTypeBitmap) As IPicture

    ' The following code is adapted from
    ' Bruce McKinney's "Hardcore Visual Basic"
    ' And Code samples from:
    ' http://www.mvps.org/vbnet/code/bitma...screenole.htmv
    ' and examples posted on MSDN

    ' The handle to the Bitmap created by CreateDibSection
    ' cannot be passed directly as the PICTDESC.Bitmap element
    ' that get's passed to OleCreatePictureIndirect.
    ' We need to create a regular bitmap from our CreateDibSection
    'Dim hBmptemp As Long, hBmpOrig As Long
    'Dim hDCtemp As Long


    'hDCtemp = apiCreateCompatibleDC(0)
    'hBmptemp = apiCreateCompatibleBitmap _
    '(mhDCImage, lpBmih.bmiHeader.biWidth, _
    'lpBmih.bmiHeader.biHeight)

    'hBmpOrig = apiSelectObject(hDCtemp, hBmptemp)

    ' lngRet = apiBitBlt(hDCtemp, 0&, 0&, lpBmih.bmiHeader.biWidth, _
    ' lpBmih.bmiHeader.biHeight, mhDCImage, 0, 0, SRCCOPY)

    'hBmptemp = apiSelectObject(hDCtemp, hBmpOrig)
    'Call apiDeleteDC(hDCtemp)

    'Fill picture description
    ' No palette info here
    ' Everything is 24bit for now
    ' picdes.hPal = hPal
    Dim picdes As PictDesc: picdes = New_PictDesc(hBmp, hPal, aPicType)
    
    ' ' Fill in magic IPicture GUID
    '{7BF80980-BF32-101A-8BBB-00AA00300CAB}
    Dim iidIPicture As IID: iidIPicture = New_IID(IID_IPicture)
    
    'Debug.Print IID_ToStr(iidIPicture)
    
    Dim IPic As IPicture
    Dim lngRet As Long: lngRet = OleCreatePictureIndirect(picdes, VarPtr(iidIPicture), True, IPic)
    '' Result will be valid Picture or Nothing-either way set it
    Set Picture_FromHandle = IPic
    
    GDIDeleteObject hBmp
End Function
'Function GetClipBoard() As Long
'    ' Adapted from original Source Code by:
'    '* MODULE NAME: Paste Picture
'    '* AUTHOR & DATE: STEPHEN BULLEN, Business Modelling Solutions Ltd.
'    '* 15 November 1998
'    '*
'    '* CONTACT: Step...@BMSLtd.co.uk
'    '* WEB SITE: http://www.BMSLtd.co.uk
'
'    ' Handles for graphic Objects
'    Dim hClipBoard As Long
'    Dim hBitmap As Long
'    Dim hBitmap2 As Long
'
'    'Check if the clipboard contains the required format
'    'hPicAvail = IsClipboardFormatAvailable(lPicType)
'
'    ' Open the ClipBoard
'    hClipBoard = OpenClipboard(0&)
'
'    If hClipBoard <> 0 Then
'    ' Get a handle to the Bitmap
'    hBitmap = GetClipboardData(CF_BITMAP)
'
'    If hBitmap = 0 Then GoTo exit_error
'    ' Create our own copy of the image on the clipboard, in the appropriate format.
'    'If lPicType = CF_BITMAP Then
'    hBitmap2 = CopyImage(hBitmap, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
'    ' Else
'    ' hBitmap2 = CopyEnhMetaFile(hBitmap, vbNullString)
'    ' End If
'
'    'Release the clipboard to other programs
'    hClipBoard = CloseClipboard
'
'    GetClipBoard = hBitmap2
'    Exit Function
'
'    End If
'
'exit_error:
'    ' Return False
'    GetClipBoard = -1
'End Function
'
Private Function New_PictDesc(ByVal hBitmap As Long, _
                              Optional ByVal hPalette As Long = 0, _
                              Optional ByVal PicType As PictureTypeConstants = PictureTypeConstants.vbPicTypeBitmap) As PictDesc
    With New_PictDesc
        .cbSize = LenB(New_PictDesc)
        .hBmp = hBitmap
        .hPal = hPalette
        .PicType = PicType
    End With
End Function

Public Function New_IID(strIID As String) As IID
    IIDFromString StrPtr(strIID), New_IID
End Function

Public Function IID_ToStr(aIID As IID) As String
    Dim s As String: s = "{"
    With aIID
        s = s & Hex8(.Data1) & "-"
        s = s & Hex4(.Data2) & "-"
        s = s & Hex4(.Data3) & "-"
        s = s & HexBA(.Data4(), 0, 2) & "-"
        s = s & HexBA(.Data4(), 2, 6)
    End With
    IID_ToStr = s & "}"
End Function

'{7BF80980-BF32-101A-8BBB-00AA00300CAB}
Public Function Hex8(ByVal Value As Long) As String
    Hex8 = Hex(Value)
    Hex8 = String(8 - Len(Hex8), "0") & Hex8
End Function
Public Function Hex4(ByVal Value As Integer) As String
    Hex4 = Hex(Value)
    Hex4 = String(4 - Len(Hex4), "0") & Hex4
End Function
Public Function Hex2(ByVal Value As Byte) As String
    Hex2 = Hex(Value)
    Hex2 = String(2 - Len(Hex2), "0") & Hex2
End Function
Public Function HexBA(Values() As Byte, startB As Byte, nBytes As Long) As String
    Dim s As String
    Dim i As Long
    For i = startB To startB + nBytes - 1
        s = s & Hex2(Values(i))
    Next
    HexBA = s
End Function


