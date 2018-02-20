Attribute VB_Name = "modAPI"
Option Explicit
DefLng A-Z

'Public Type POINTAPI
'    X          As Long
'    Y          As Long
'End Type

'fro CAIRO use
Public Type POINTAPI
    X       As Double
    Y       As Double
End Type



Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function Arc Lib "gdi32" (ByVal HDC As Long, _
                                         ByVal xInizioRettangolo As Long, _
                                         ByVal yInizioRettangolo As Long, _
                                         ByVal xFineRettangolo As Long, _
                                         ByVal yFineRettangolo As Long, _
                                         ByVal xInizioArco As Long, _
                                         ByVal yInizioArco As Long, _
                                         ByVal xFineArco As Long, _
                                         ByVal yFineArco As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal HDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

Private Declare Function Rectangle Lib "gdi32.dll" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


'Declare Function Arc Lib "gdi32.dll" (ByVal HDC As Long, ByVal X1 As Long, _
 ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
 ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Public PrevColor As Long
Public PrevWidth As Long


Public Sub SetBrush(ByVal HDC As Long, ByVal PenWidth As Long, ByVal PenColor As Long)


    DeleteObject (SelectObject(HDC, CreatePen(vbSolid, PenWidth, PenColor)))
    'kOBJ = SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor))
    'SetBrush = kOBJ


End Sub




Public Sub FastLine(ByRef HDC As Long, ByRef X1 As Long, ByRef Y1 As Long, _
                    ByRef X2 As Long, ByRef Y2 As Long, ByRef W As Long, ByRef color As Long)

    Dim poi As POINTAPI

    'SetBrush hdc, W, color
    'If Color <> PrevColor Or W <> PrevWidth Then

    'If Color <> PrevColor Then
    DeleteObject (SelectObject(HDC, CreatePen(vbSolid, W, color)))
    '    PrevColor = Color
    '    PrevWidth = W
    'End If

    MoveToEx HDC, X1, Y1, poi
    LineTo HDC, X2, Y2

End Sub

Sub MyCircle(ByRef HDC As Long, ByRef X As Long, ByRef Y As Long, ByRef R As Long, W As Long, color)
    Dim XpR As Long

    'If color <> PrevColor Or w <> PrevWidth Then
    DeleteObject (SelectObject(HDC, CreatePen(vbSolid, W, color)))
    '    PrevColor = color
    '    PrevWidth = w
    'End If

    XpR = X + R

    Arc HDC, X - R, Y - R, XpR, Y + R, XpR, Y, XpR, Y

End Sub


Public Sub bLOCK(ByRef HDC As Long, X As Long, Y As Long, W As Long, h As Long, color As Long)

    DeleteObject (SelectObject(HDC, CreatePen(vbSolid, 1, color)))
    DeleteObject (SelectObject(pHDC, CreateSolidBrush(color)))

    Rectangle HDC, X, Y, X + W, Y + h

End Sub


