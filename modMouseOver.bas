Attribute VB_Name = "modMouseOver"

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long


Public Function MouseOver(ByVal hWnd As Long) As Boolean
 Dim Mouse As POINTAPI
 Dim lpRect As RECT
 
 GetCursorPos Mouse
 GetWindowRect hWnd, lpRect
 
 If Mouse.X >= lpRect.Left And Mouse.X <= lpRect.Right And _
    Mouse.Y >= lpRect.Top And Mouse.Y <= lpRect.Bottom Then
        MouseOver = True
 Else
        MouseOver = False
 End If
 
End Function


