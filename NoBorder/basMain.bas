Attribute VB_Name = "basMain"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Form_DragX, Form_DragY As Single
Public MinMax As Boolean

Public Function PutOnTop(F As Form)
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Function

Public Function FormDesign(frmName As Form, Height As Integer, Width As Integer, WindowName As String)
    
    PutOnTop frmName

    frmName.Height = Height
    frmName.Width = Width
    
    frmName.TitleBar.Height = 280
    frmName.TitleBar.Width = frmName.Width
    
    frmName.TitleBarText.Caption = WindowName
    frmName.TitleBarText.Width = Width - frmName.CloseWindow.Width
    
    frmName.CloseWindow.Top = frmName.TitleBarText.Top - 20
    frmName.CloseWindow.Left = frmName.Width - 200
    
    frmName.MainWindow.Height = frmName.Height - 280
    frmName.MainWindow.Width = frmName.Width

End Function

Public Function MinMaxWindow(frmName As Form)
Dim i As Integer

    If MinMax = False Then
        Do Until frmName.Height <= frmName.TitleBar.Height
        frmName.Height = frmName.Height - 5
            For i = 1 To 100
                DoEvents
            Next i
        Loop
        MinMax = True
    Else
        Do Until frmName.Height >= frmName.MainWindow.Height + 280
        frmName.Height = frmName.Height + 5
            For i = 1 To 100
            DoEvents
            Next i
        Loop
        MinMax = False
    End If
End Function

Public Function TitleBarClick(frmName As Form, button As Integer, x As Single, y As Single)
    
    If button = 1 Then
        Form_DragX = x
        Form_DragY = y
    End If
    
    If button = 2 Then
        MinMaxWindow frmName
    End If
    
End Function

Public Function WindowMove(frmName As Form, x As Single, y As Single)
    
    If Form_DragX > 0 Then
        frmName.Move frmName.Left + x - Form_DragX + 15, frmName.Top + y - Form_DragY + 15:
    End If

End Function

Public Function WindowStopMove()
Form_DragX = 0
End Function
