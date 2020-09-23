Attribute VB_Name = "mdlLight"

'APIs:
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Variables (these are needed for GetRGBs):
Dim Red As Long
Dim Green As Long
Dim Blue As Long

Public Sub DrawLight(Target As Long, X As Integer, Y As Integer, RedB As Byte, GreenB As Byte, BlueB As Byte, Radius As Long, NumberOfSteps As Long)
    Dim cX As Long 'X counter
    Dim cY As Long 'Y counter
    Dim TempColor As Long
    Dim TempRadius As Integer
    
    'This boolean array holds the state of
    'each pixel (if it was drawn or not).
    'It's useful to not draw a pixel twice.
    Dim Done() As Boolean
    ReDim Done(-Radius To Radius, -Radius To Radius)
    
    For i = 1 To NumberOfSteps
        
        'Update "TempRadius" (the radius of the
        'circle that is currently being drawn)
        TempRadius = Radius / NumberOfSteps * i
        
        For cX = -TempRadius To TempRadius
            For cY = -TempRadius To TempRadius
                If Not Done(cX, cY) Then 'The pixel hasn't been drawn yet
                    'Next is the formula for getting the circle
                    If (cX * cX) + (cY * cY) <= TempRadius * TempRadius Then
                        'Get the pixel and extract RGBs
                        TempColor = GetPixel(Target, cX + X, cY + Y)
                        GetRGBs TempColor
                        
                        'Increase RGB values with the given brightnesses
                        Red = Red + RedB * (NumberOfSteps - i)
                        If Red > 255 Then Red = 255
                        
                        Green = Green + GreenB * (NumberOfSteps - i)
                        If Green > 255 Then Green = 255
                        
                        Blue = Blue + BlueB * (NumberOfSteps - i)
                        If Blue > 255 Then Blue = 255
                        
                        'Draw pixel
                        SetPixel Target, cX + X, cY + Y, RGB(Red, Green, Blue)
                        
                        'This pixel has been drawn, so write it in the
                        'array (so it won't be drawn twice)
                        Done(cX, cY) = True
                    End If
                End If
            Next cY
        Next cX
    Next i
End Sub

    ' This function extracts the Red, Green and Blue
    'values from a color to 3 variables with their
    'names. You may keep it as well, but please give
    'me some credit for it too ;)
    'By the way, this function isn't very exact, so
    'you may end up with the green and blue values
    'not as they should (plus or minus 1 or 2).

Public Sub GetRGBs(RGBVal As Long)

    If RGBVal = 16777215 Then
        Red = 255
        Green = 255
        Blue = 255
        Exit Sub
    End If

    Red = RGBVal And 255
    Green = RGBVal / 256 And 255
    Blue = RGBVal / 65536 And 255
    
End Sub

