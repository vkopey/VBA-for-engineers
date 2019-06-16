Public Name As String
Public ModelViewName As Integer '0-*Front,1-*Front,2-*Back,3-*Left,4-*Right,5-*Top,6-*Bottom
Public PositionX As Double
Public PositionY As Double
Public ReferenceModelPath As String
Public Configuration As String

Private Sub Class_Initialize()
Name = ""
ModelViewName = 0
PositionX = 0
PositionY = 0
ReferenceModelPath = ""
Configuration = ""
End Sub
