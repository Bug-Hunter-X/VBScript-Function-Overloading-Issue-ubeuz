Function CalculateArea(shape, param1, param2)
  Select Case shape
    Case "rectangle"
      CalculateArea = param1 * param2
    Case "circle"
      CalculateArea = 3.14159 * param1 ^ 2
    Case Else
      CalculateArea = -1 ' Indicate an error
  End Select
End Function

' Example usage:
Dim rectArea, circleArea
rectArea = CalculateArea("rectangle", 5, 10)  ' Correctly calculates rectangle area
circleArea = CalculateArea("circle", 5)      ' Correctly calculates circle area
MsgBox "Rectangle Area: " & rectArea & ", Circle Area: " & circleArea

' Incorrect attempt (using multiple functions with same name):
'Function CalculateArea(length, width)
'  CalculateArea = length * width
'End Function
'
'Function CalculateArea(radius)
'  CalculateArea = 3.14159 * radius ^ 2
'End Function