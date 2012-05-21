Attribute VB_Name = "modColors"
Public Declare Sub ColorRGBToHLS Lib "SHLWAPI.DLL" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Public Declare Function ColorHLSToRGB Lib "SHLWAPI.DLL" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long

Public Sub ColorRGBToGamma(ByVal clrRGB As Long, ByRef intGammaRed As Integer, ByRef intGammaGreen As Integer, ByRef intGammaBlue As Integer)
    Dim intRed, intGreen, intBlue As Integer
    
    intRed = RGBRed(clrRGB)
    intGreen = RGBGreen(clrRGB)
    intBlue = RGBBlue(clrRGB)
    
    intGammaRed = (intRed - 128) * 32
    intGammaGreen = (intGreen - 128) * 32
    intGammaBlue = (intBlue - 128) * 32
    
End Sub

Public Function ColorGammaToRGB(ByRef intGammaRed As Integer, ByRef intGammaGreen As Integer, ByRef intGammaBlue As Integer) As Long
    Dim intRed, intGreen, intBlue As Integer
    
    intRed = Int((intGammaRed / 32) + 128)
    intGreen = Int((intGammaGreen / 32) + 128)
    intBlue = Int((intGammaBlue / 32) + 128)
    
    ColorGammaToRGB = RGB(intRed, intGreen, intBlue)
End Function

Public Function RGBRed(RGBCol As Long) As Integer
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
