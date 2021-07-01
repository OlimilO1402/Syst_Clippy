VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const GENERIC_READ      As Long = &H80000000
Private Const GENERIC_WRITE     As Long = &H40000000

Private Const FILE_SHARE_READ   As Long = &H1
Private Const FILE_SHARE_WRITE  As Long = &H2
Private Const FILE_SHARE_DELETE As Long = &H4

Private Const CREATE_ALWAYS     As Long = 2
Private Const CREATE_NEW        As Long = 1
Private Const OPEN_ALWAYS       As Long = 4
Private Const OPEN_EXISTING     As Long = 3
Private Const TRUNCATE_EXISTING As Long = 5

Private Const FILE_ATTRIBUTE_READONLY  As Long = &H1
Private Const FILE_ATTRIBUTE_HIDDEN    As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM    As Long = &H4
Private Const FILE_ATTRIBUTE_ARCHIVE   As Long = &H20
Private Const FILE_ATTRIBUTE_NORMAL    As Long = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Private Const FILE_ATTRIBUTE_OFFLINE   As Long = &H1000
Private Const FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000

Private Const FILE_FLAG_OPEN_NO_RECALL     As Long = &H100000
Private Const FILE_FLAG_OPEN_REPARSE_POINT As Long = &H200000
Private Const FILE_FLAG_SESSION_AWARE      As Long = &H800000
Private Const FILE_FLAG_POSIX_SEMANTICS    As Long = &H1000000
Private Const FILE_FLAG_BACKUP_SEMANTICS   As Long = &H2000000
Private Const FILE_FLAG_DELETE_ON_CLOSE    As Long = &H4000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN    As Long = &H8000000
Private Const FILE_FLAG_RANDOM_ACCESS      As Long = &H10000000
Private Const FILE_FLAG_NO_BUFFERING       As Long = &H20000000
Private Const FILE_FLAG_OVERLAPPED         As Long = &H40000000
Private Const FILE_FLAG_WRITE_THROUGH      As Long = &H80000000


Private Const SecurityAnonymous As Long = 1
Private Const SecurityIdentification As Long = 2
Private Const SECURITY_ANONYMOUS As Long = (SecurityAnonymous * 2 ^ 16)
Private Const SECURITY_CONTEXT_TRACKING As Long = &H40000
'Private Const SECURITY_DELEGATION As Long = (SecurityDelegation * 2 ^ 16)
Private Const SECURITY_EFFECTIVE_ONLY As Long = &H80000
Private Const SECURITY_IDENTIFICATION As Long = (SecurityIdentification * 2 ^ 16)
'Private Const SECURITY_IMPERSONATION As Long = (SecurityImpersonation * 2 ^ 16)

'Private Const INVALID_HANDLE_VALUE As Long = ((Handle) - 1)
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
     ByVal lpFileName As String, _
     ByVal dwDesiredAccess As Long, _
     ByVal dwShareMode As Long, _
     ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
     ByVal dwCreationDisposition As Long, _
     ByVal dwFlagsAndAttributes As Long, _
     ByVal hTemplateFile As Long) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
     ByVal hObject As Long) As Long
     
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private m_hnd  As Long
Private m_Name As String
Private m_sa   As SECURITY_ATTRIBUTES



Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal Value As String)
    m_Name = Value
    m_sa.nLength = LenB(m_sa)
    m_hnd = CreateFile(m_Name, GENERIC_READ, FILE_SHARE_READ, m_sa, OPEN_EXISTING, 0, 0)
    
End Property

Public Property Get Handle() As Long
    Handle = m_hnd
End Property

Private Sub Class_Terminate()
    If m_hnd <> 0 Then CloseHandle (m_hnd)
End Sub
