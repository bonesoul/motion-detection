Attribute VB_Name = "modWebcam"
Public strFileName As String

Public mdTriger As Single
Public mdSample(50, 50, 250) As Single

Public Declare Function GetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long


