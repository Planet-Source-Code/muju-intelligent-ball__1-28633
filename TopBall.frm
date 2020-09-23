VERSION 5.00
Begin VB.Form MyForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   ControlBox      =   0   'False
   Icon            =   "TopBall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TopBall.frx":164A
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '----Start Drag
    DragPosL.X = X
    DragPosL.Y = Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '----Drag
    If DragPosL.X <> 0 And DragPosL.Y <> 0 Then
        DragPos.X = X - DragPosL.X
        DragPos.Y = Y - DragPosL.Y
        GetWindowRect MyForm.hWnd, TempRect
        MoveWindow MyForm.hWnd, DragPos.X + TempRect.Left, DragPos.Y + TempRect.Top, (BallInfo.BSize * 2), (BallInfo.BSize * 2), True
        'DragPosL.X = X
        'DragPosL.Y = Y
        'Debug.Print DragPos.X; DragPos.Y; DragPosL.X; DragPosL.Y
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '----Stop Drag
    DragPos.X = X + DragPosL.X
    DragPos.Y = Y + DragPosL.Y
    DragPosL.X = 0
    DragPosL.Y = 0
    GetWindowRect MyForm.hWnd, TempRect
    BallInfo.BPosX = TempRect.Left + BallInfo.BSize
    BallInfo.BPosY = TempRect.Top + BallInfo.BSize
    BallInfo.BVelX = (X - DragPos.X) / 10
    BallInfo.BVelY = (Y - DragPos.Y) / 10
End Sub
