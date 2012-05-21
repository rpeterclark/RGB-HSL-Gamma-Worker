VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ===============================================================================================
' vbAccelerator
' http://vbaccelerator.com/
' © 1999 Steve McMahon (steve@vbaccelerator.com)
'
' cMemDC
' A simple wrapper around a Memory Device Context and
' a bitmap.
'
' ===============================================================================================


Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAP '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type


Private m_hDC As Long
Private m_hBmp As Long
Private m_hBmpOld As Long

Private m_iSizeX As Long
Private m_iSizeY As Long

Private m_bMono As Boolean

Private m_bInit As Boolean

Private Sub SetSize(ByVal x As Long, ByVal y As Long)
Dim bInit As Boolean
   If x > m_iSizeX Or y > m_iSizeY Then
      m_bInit = True
   End If
   m_iSizeX = x
   m_iSizeY = y
End Sub
Public Property Get Width() As Long
   Width = m_iSizeX
End Property
Public Property Get Height() As Long
   Height = m_iSizeY
End Property
Public Property Let Width(ByVal lWidth As Long)
   If m_iSizeX < lWidth Then
      m_bInit = True
   End If
   m_iSizeX = lWidth
End Property
Public Property Let Height(ByVal lHeight As Long)
   If m_iSizeY < lHeight Then
      m_bInit = True
   End If
   m_iSizeY = lHeight
End Property
Public Property Get hDC() As Long
   If m_bInit Then
      If Init Then
         m_bInit = False
      End If
   End If
   hDC = m_hDC
End Property
Public Property Get Mono() As Boolean
   Mono = m_bMono
End Property
Public Property Let Mono(ByVal bState As Boolean)
   If Not (m_bMono = bState) Then
      m_bInit = True
   End If
   m_bMono = bState
End Property
Private Function Init(Optional ByVal hBmp As Long = 0) As Boolean
Dim hDCDisp As Long
Dim lWidth As Long, lHeight As Long
   
   If m_bMono Then
      If m_hDC = 0 Then
         m_hDC = CreateCompatibleDC(0)
      End If
   Else
      hDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      
      If hDCDisp <> 0 Then
         If m_hDC = 0 Then
            m_hDC = CreateCompatibleDC(hDCDisp)
            If m_hDC = 0 Then
               ThrowError 1
               Exit Function
            End If
         End If
      Else
         ThrowError 1
         Exit Function
      End If
   End If
   
   If m_hDC <> 0 Then
      If m_hBmpOld <> 0 Then
         SelectObject m_hDC, m_hBmpOld
         m_hBmpOld = 0
      End If
      If m_hBmp <> 0 Then
         DeleteObject m_hBmp
         m_hBmp = 0
      End If
      If hBmp = 0 Then
         If m_bMono Then
            m_hBmp = CreateCompatibleBitmap(m_hDC, m_iSizeX, m_iSizeY)
         Else
            m_hBmp = CreateCompatibleBitmap(hDCDisp, m_iSizeX, m_iSizeY)
         End If
      Else
         m_hBmp = hBmp
      End If
      If m_hBmp = 0 Then
         DeleteDC hDCDisp
         hDCDisp = 0
         ThrowError 1
      Else
         m_hBmpOld = SelectObject(m_hDC, m_hBmp)
         ' ok.
         Init = True
      End If
   Else
      DeleteDC hDCDisp
      hDCDisp = 0
      ThrowError 1
   End If
   
   If Not (hDCDisp = 0) Then
      DeleteDC hDCDisp
   End If
   
   
End Function

Private Function ThrowError(ByVal lErr As Long)
Dim sMsg As String
   If lErr = 1 Then
      ' fatal
      Destroy
      Err.Raise 7
   Else
      Err.Raise 7
   End If
End Function
Private Function Destroy()
   If m_hBmpOld <> 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If m_hBmp <> 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If m_hDC <> 0 Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
   
End Function
Public Function ExtractBitmap() As Long
   If m_hBmpOld <> 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If m_hDC <> 0 Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
   ExtractBitmap = m_hBmp
   ' the responsibility for clearing up
   ' the bitmap now belongs to the caller:
   m_hBmp = 0
   m_bInit = True
End Function
Public Sub InjectBitmap(ByVal hBmp As Long)
Dim tBM As BITMAP
   Destroy
   GetObjectAPI hBmp, Len(tBM), tBM
   Width = tBM.bmWidth
   Height = tBM.bmHeight
   Init hBmp
   m_bInit = False
End Sub
Private Sub Class_Terminate()
   Destroy
End Sub


