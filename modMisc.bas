Attribute VB_Name = "modMisc"
Option Explicit

Public Function ZHex(lHex As Long, iZeros As Integer) As String
'returns a HEX string of specified length (pad zeros on left)
    ZHex = Right$(String$(iZeros - 1, "0") & Hex$(lHex), iZeros)
End Function

Public Function UtoH(U As String) As String
'takes the Unicode string of custom colors and converts it to hex codes
'that can be easily saved
Dim i As Integer, strHex As String, strH As String
    For i = 1 To Len(U)
        strH = ZHex(AscW(Mid$(U, i, 1)), 4)
        strHex = strHex & strH
    Next i
    UtoH = strHex
End Function

Public Sub HtoU(strHex As String)
'Takes the hex string and loads the custom colors
Dim i As Integer, strU As String
Dim customcolors() As Byte  ' dynamic (resizable) array
    If strHex = "" Then ColorDialog.lpCustColors = "": Exit Sub
    
    ReDim customcolors(0 To (Len(strHex) / 4)) As Byte  'resize the array
    
    For i = 3 To Len(strHex) - 1 Step 4
        customcolors((i - 3) / 4) = Val("&H" & Mid$(strHex, i, 2))
    Next i
    ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
End Sub


