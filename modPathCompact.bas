Attribute VB_Name = "modPathCompact"
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_BOTTOM = &H8&
Private Const DT_CENTER = &H1&
Private Const DT_LEFT = &H0&
Private Const DT_CALCRECT = &H400&
Private Const DT_WORDBREAK = &H10&
Private Const DT_VCENTER = &H4&
Private Const DT_TOP = &H0&
Private Const DT_TABSTOP = &H80&
Private Const DT_SINGLELINE = &H20&
Private Const DT_RIGHT = &H2&
Private Const DT_NOCLIP = &H100&
Private Const DT_INTERNAL = &H1000&
Private Const DT_EXTERNALLEADING = &H200&
Private Const DT_EXPANDTABS = &H40&
Private Const DT_CHARSTREAM = 4&
Private Const DT_NOPREFIX = &H800&
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_PATH_ELLIPSIS = &H4000&
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal lpszPath As String, ByVal dx As Long) As Long

Public Function CompactedPath(ByVal sPath As String, ByVal lMaxPixels As Long, ByVal hDC As Long) As String
    Dim tR As RECT
    tR.right = lMaxPixels
    DrawText hDC, sPath, -1, tR, DT_PATH_ELLIPSIS Or DT_SINGLELINE Or DT_MODIFYSTRING
    CompactedPath = sPath
End Function

Public Function CompactedPathSh(ByVal sPath As String, ByVal lMaxPixels As Long, ByVal hDC As Long) As String
    Dim lR As Long
    Dim iPos As Long
    lR = PathCompactPath(hDC, sPath, lMaxPixels)
    iPos = InStr(sPath, Chr$(0))
    If iPos <> 0 Then
        CompactedPathSh = left$(sPath, iPos - 1)
    Else
        CompactedPathSh = sPath
    End If
End Function
