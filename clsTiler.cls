VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   left As Long
   tOp As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
    Private Const OPAQUE = 2
    Private Const TRANSPARENT = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Const WM_DESTROY = &H2
Private Const WM_PAINT = &HF
Private Const WM_COMMAND = &H111
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_PRINT = &H317
Private Const WM_PRINTCLIENT = &H318
Private Const WM_ERASEBKGND = &H14

Private Const EN_CHANGE = &H300
Private Const EN_HSCROLL = &H601
Private Const EN_VSCROLL = &H602

Private Const PRF_CHECKVISIBLE = &H1&
Private Const PRF_NONCLIENT = &H2&
Private Const PRF_CLIENT = &H4&
Private Const PRF_ERASEBKGND = &H8&
Private Const PRF_CHILDREN = &H10&
Private Const PRF_OWNED = &H20&

Implements ISubclass
Private m_hWnd As Long
Private m_hWndParent As Long
Private m_bCE As Boolean
Private m_bMouseDown As Boolean
Private m_iScroll As Long
Private m_tR As RECT
Private m_bFixBackground As Boolean
Private m_xTileOffset As Long
Private m_yTileOffset As Long

Private m_cWorkDC As New clsMemDC
Private m_cBackDC As New clsMemDC

Public Sub SetBackdrop(pic As IPicture)
   Dim cTempDC As New clsMemDC
   ' Create DC and put bmp into it:
   cTempDC.InjectBitmap pic.Handle
   ' Set up BackDC:
   m_cBackDC.Width = cTempDC.Width
   m_cBackDC.Height = cTempDC.Height
   BitBlt m_cBackDC.hDC, 0, 0, m_cBackDC.Width, m_cBackDC.Height, cTempDC.hDC, 0, 0, vbSrcCopy
   ' Return bitmap
   cTempDC.ExtractBitmap
End Sub

Public Property Get TileOffsetX() As Long
   TileOffsetX = m_xTileOffset
End Property
Public Property Let TileOffsetX(ByVal xOffset As Long)
   m_xTileOffset = xOffset
End Property
Public Property Get TileOffsetY() As Long
   TileOffsetY = m_yTileOffset
End Property
Public Property Let TileOffsetY(ByVal yOffset As Long)
   m_yTileOffset = yOffset
End Property

Public Sub TileArea( _
        ByVal hdcTo As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long _
      )
Dim SrcWidth As Long
Dim SrcHeight As Long
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long
      
   SrcWidth = m_cBackDC.Width
   SrcHeight = m_cBackDC.Height
   
    lSrcStartX = ((x + m_xTileOffset) Mod SrcWidth)
    lSrcStartY = ((y + m_yTileOffset) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, _
               m_cBackDC.hDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub

Private Sub PostPaint(Optional ByVal bUseArea As Boolean = False)
Dim lHDC As Long
Dim tR As RECT

   'Debug.Print "PostPaint", Hex$(m_hWnd), bUseArea

   If m_bFixBackground Then
      bUseArea = False
   End If
   
   ' Get size of textbox:
   GetClientRect m_hWnd, tR
   
   ' Set work DC to right size:
   m_cWorkDC.Width = tR.Right - tR.left + 1
   m_cWorkDC.Height = tR.Bottom - tR.tOp + 1
   
   ' Copy background into workdc:
   If Not m_bFixBackground Then
      TileArea m_cWorkDC.hDC, _
         m_tR.left, m_tR.tOp, _
         m_tR.Right - m_tR.left + 1, m_tR.Bottom - m_tR.tOp + 1
   Else
      TileArea m_cWorkDC.hDC, _
         0, 0, m_cWorkDC.Width, m_cWorkDC.Height
   End If
   
   ' Ask text box to draw itself into the workdc:
   If bUseArea Then
   Else
      SendMessageLong m_hWnd, WM_PRINT, m_cWorkDC.hDC, PRF_CLIENT Or PRF_CHECKVISIBLE
   End If
   
   ' Draw the results into the textbox:
   lHDC = GetDC(m_hWnd)
   If Not m_bFixBackground Then
      BitBlt lHDC, m_tR.left, m_tR.tOp, m_tR.Right - m_tR.left + 1, m_tR.Bottom - m_tR.tOp + 1, m_cWorkDC.hDC, m_tR.left, m_tR.tOp, vbSrcCopy
   Else
      BitBlt lHDC, 0, 0, tR.Right - tR.left + 1, tR.Bottom - tR.tOp + 1, m_cWorkDC.hDC, 0, 0, vbSrcCopy
   End If
   ReleaseDC m_hWnd, lHDC
   
End Sub
Public Sub Attach(ByVal hWndA As Long)
   Detach
   m_hWnd = hWndA
   AttachMessage Me, m_hWnd, WM_PAINT
   m_hWndParent = GetParent(m_hWnd)
   AttachMessage Me, m_hWndParent, WM_CTLCOLOREDIT
   AttachMessage Me, m_hWndParent, WM_COMMAND
   AttachMessage Me, m_hWnd, WM_ERASEBKGND
   AttachMessage Me, m_hWnd, WM_DESTROY
End Sub
Public Sub Detach()
   If Not (m_hWnd = 0) Then
      DetachMessage Me, m_hWnd, WM_PAINT
      If Not m_hWndParent = 0 Then
         DetachMessage Me, m_hWndParent, WM_CTLCOLOREDIT
         DetachMessage Me, m_hWndParent, WM_COMMAND
      End If
      m_hWndParent = 0
      DetachMessage Me, m_hWnd, WM_ERASEBKGND
      DetachMessage Me, m_hWnd, WM_DESTROY
      m_hWnd = 0
   End If
End Sub

Private Sub Class_Initialize()
   m_bFixBackground = True
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
   Select Case CurrentMessage
   Case WM_DESTROY
      ISubclass_MsgResponse = emrPreprocess
   Case Else
      ISubclass_MsgResponse = emrConsume
   End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim wCode As Long
   '
   Select Case iMsg
   Case WM_DESTROY
      Detach
   Case WM_PAINT
      ' Debug.Print "Paint", Hex$(hwnd)
      GetUpdateRect hwnd, m_tR, 0
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      If m_iScroll = 0 Then
         PostPaint
      Else
         Select Case m_iScroll
         Case 1
            m_iScroll = 1
         Case 2
            m_iScroll = 0
         End Select
      End If
   Case WM_CTLCOLOREDIT
      If m_hWnd = lParam Then
         'Debug.Print "CtlColorEdit", Hex$(m_hWnd), wParam, dc
         SetBkMode wParam, TRANSPARENT
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         If m_iScroll = 0 Then
            If Not m_bCE Then
               m_bCE = True
               PostPaint
               m_bCE = False
            End If
         Else
            Select Case m_iScroll
            Case 1
               m_iScroll = 2
               If Not m_bCE Then
                  m_bCE = True
'                  Debug.Print "Scroll", m_tR.left, m_tR.top, m_tR.right, m_tR.bottom
                  PostPaint True
                  m_bCE = False
               End If
            Case 2
               'Debug.Print "CtlColorEdit;Scroll=2"
               m_iScroll = 0
            End Select
         End If
      Else
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
   Case WM_COMMAND
      If m_hWnd = lParam Then
         wCode = wParam \ &H10000
         If wCode = EN_HSCROLL Or wCode = EN_VSCROLL Then
            'Debug.Print "SCROLL"
            m_iScroll = 1
         End If
      End If
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
   Case WM_ERASEBKGND
      ISubclass_WindowProc = 1
   End Select
End Function


