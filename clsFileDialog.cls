VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Owner As Object
Private m_Filter As String
Private m_Flags As FileDlgConstants
Private m_FileName As String

Public Enum FileDlgConstants
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHAREWARN = 0
    OFN_SHARENOWARN = 1
    OFN_SHOWHELP = &H10
    OFS_MAXPATHNAME = 128
End Enum

Private Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Private Const BaseFlags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Property Get FileName() As String
    FileName = m_FileName
End Property


Public Property Get Filter() As String
    Filter = m_Filter
End Property
Public Property Let Filter(NewFilter As String)
    m_Filter = NewFilter
End Property
Public Property Get Flags() As FileDlgConstants
    Flags = m_Flags
End Property
Public Property Let Flags(NewFlags As FileDlgConstants)
    m_Flags = NewFlags
End Property

Private Function GetWinFilter(VBFilter As String) As String
    GetWinFilter = Replace(VBFilter, "|", Chr(0)) & Chr(0)
End Function
Private Function NullTrim(NullStr As String) As String
    NullTrim = left(NullStr, InStr(NullStr, Chr(0)) - 1)
End Function
Public Sub ShowOpen()
    Dim RetVal As Long, OpenFile As OPENFILENAME
    
    With OpenFile
        .nStructSize = Len(OpenFile)
        .hwndOwner = m_Owner.hwnd
        .sFilter = GetWinFilter(m_Filter)
        
        .sFile = Chr(0)
        .sFile = Space(1024) & Chr(0)
        .nFileSize = Len(OpenFile.sFile)
        
        .sFileTitle = Space(512)
        .nTitleSize = Len(.sFileTitle)
        .Flags = BaseFlags Or m_Flags
    End With
    
    RetVal = GetOpenFileName(OpenFile)
    If (RetVal <> 0) Then
        m_FileName = NullTrim(OpenFile.sFile)
    End If
End Sub
Public Property Get Owner() As Object
    Set Owner = m_Owner
End Property

Public Property Let Owner(NewOwner As Object)
    Set m_Owner = NewOwner
End Property



