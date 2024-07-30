Dim obj As New ClsReadCard
Dim info As String
Private Sub Command1_Click()
    info = obj.ReadIDCardNo("")
    MsgBox info
End Sub
Private Sub Command2_Click()
    info = obj.ReadInusCardNo("")
    MsgBox info
End Sub
Private Sub Command3_Click()
    info = obj.ReadIDCardInfo("")
    MsgBox info
End Sub
Private Sub Command4_Click()
    info = obj.ReadSSPersonInfo()
    MsgBox info
End Sub

Private Sub Command5_Click()
    info = obj.ReadPersonInfo()
    MsgBox info
End Sub

Private Sub Command6_Click()
    info = obj.ReadSSMagCard()
    MsgBox info
End Sub

Private Sub Command7_Click()
    info = obj.ReadMagCard()
    MsgBox info
End Sub
