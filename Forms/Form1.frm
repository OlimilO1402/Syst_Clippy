VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Test3"
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test2"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test1"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5175
   End
   Begin VB.CommandButton BtnReadCBConstants 
      Caption         =   "Read ClipBoard Constants"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   6015
      Left            =   5400
      ScaleHeight     =   5955
      ScaleWidth      =   7515
      TabIndex        =   7
      Top             =   960
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      Height          =   6015
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   960
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "Click in List to show Text in TextBox"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clippy As New CClipBoard
Dim FileH  As New FileH
Dim sPFN As String

Dim ListOfFileH As Collection

Private Sub Form_Load()
    sPFN = "C:\TestDir\TestFile.txt"
End Sub

Private Sub BtnReadCBConstants_Click()
    Clippy.CBFormats_ToListBox List1
End Sub

Private Sub Command1_Click()
    
    FileH.Name = sPFN
    
End Sub

Private Sub Command2_Click()
    Clippy.Data(CF_HDROP) = FileH.handle
End Sub

Private Sub Command3_Click()
'in dotnet:
'    Dim DataObject As New DataObject
'    Dim tempFileArray(0) As String
'    'NOTE THAT IT MUST BE PASSED IN AN ARRAY!!!
'    tempFileArray(0) = activeListView.SelectedItems(0).Tag
'    'DataObject.SetData (DataFormats.FileDrop, False, tempFileArray)
'    DataObject.SetData tempFileArray, CF_HDROP
'    'Clipboard.SetData DataObject
    
    ReDim aFiles(0 To 2) As String
    
    aFiles(0) = "C:\TestDir\TestFile1.txt"   ' File
    aFiles(1) = "C:\TestDir\TestFile2.txt"   ' File
    aFiles(2) = "C:\TestDir\test1\"          ' Folder
    
    'Set ListOfFileH = New Collection
    'MClipBoard.ClipboardCopyFiles aFiles()
    ' copy to Clipboard
    Debug.Print Clippy.CopyFiles(aFiles)
    
End Sub

Private Sub List1_Click()
    Dim i As Integer: i = List1.ListIndex
    Dim s As String:  s = List1.List(i)
    Dim sa() As String: sa = Split(s, ",")
    Dim cf As Long: cf = CLng(sa(0))
    
    Select Case cf
    Case 0, 1, 7, 13, 16
        
        If Clippy.HasFormat(cf) Then
            Text1.Text = Clippy.StrData(cf)
            Text1.ZOrder 0
        End If
    
    Case CF_HTML, CF_HTML_xls1, CF_HTML_xls2
        
        Text1.Text = Clippy.StrData(cf)
        Text1.ZOrder 0
    
    Case CF_BITMAP, CF_METAFILEPICT, CF_ENHMETAFILE, CF_PICTURE
        
        'Set Picture1.Picture = Clippy.GetData(cf)
        'Set Picture1.Picture = Clipboard.GetData(cf)
        
        Set Picture1.Picture = Clippy.GetPicture(cf)
        Picture1.ZOrder 0
        
    Case Else
        Text1.Text = ""
        Text1.ZOrder 0
    End Select
End Sub


