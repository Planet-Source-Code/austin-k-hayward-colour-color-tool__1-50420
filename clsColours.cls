VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CurrentModule As String = "clsColours"


Public Function GetHexFromRGB(ByVal lRed As Long, ByVal lGreen As Long, ByVal lBlue As Long) As String

On Error GoTo Err_GetHexFromRGB

    'take the rgb values and convert to hex

    Dim lColour As Long
    Dim sHex As String

    lColour = RGB(lRed, lGreen, lBlue)

    sHex = Hex(lColour)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If

    GetHexFromRGB = "&H" & sHex & "&"

Exit Function

Err_GetHexFromRGB:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "GetHexFromRGB", Err.Number, Err.Description

End Function

Public Function GetHexFromDecimal(lColour As Long) As String

On Error GoTo Err_GetHexFromDecimal

    Dim sHex As String

    sHex = Hex(lColour)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If

    GetHexFromDecimal = "&H" & sHex & "&"

Exit Function

Err_GetHexFromDecimal:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "GetHexFromDecimal", Err.Number, Err.Description

End Function

Public Function GetDecimalFromRGB(ByVal lRed As Long, ByVal lGreen As Long, ByVal lBlue As Long) As Long

On Error GoTo Err_GetDecimalFromRGB

    GetDecimalFromRGB = RGB(lRed, lGreen, lBlue)

Exit Function

Err_GetDecimalFromRGB:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "GetDecimalFromRGB", Err.Number, Err.Description

End Function

Public Function GetHTMLColourCodeFromRGB(ByVal lRed As Long, ByVal lGreen As Long, ByVal lBlue As Long) As String

On Error GoTo Err_GetHTMLColourCodeFromRGB

    Dim lColour As Long
    Dim sHex As String

    lColour = RGB(lRed, lGreen, lBlue)

    sHex = Hex(lColour)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    
    GetHTMLColourCodeFromRGB = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)

Exit Function

Err_GetHTMLColourCodeFromRGB:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "GetHTMLColourCodeFromRGB", Err.Number, Err.Description

End Function

Public Function GetOppositeColour(ByVal lRed As Long, ByVal lGreen As Long, ByVal lBlue As Long) As Long

On Error GoTo Err_GetOppositeColour

    Dim tRed, tGreen, tBlue As Long

    tRed = Abs(lRed - 255)
    tGreen = Abs(lGreen - 255)
    tBlue = Abs(lBlue - 255)

    GetOppositeColour = RGB(tRed, tGreen, tBlue)

Exit Function

Err_GetOppositeColour:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "GetOppositeColour", Err.Number, Err.Description

End Function

















