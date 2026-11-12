Private Sub Command1_Click()
    With Data1.Recordset
        .AddNew
        .Fields(0) = Text1.Text
        .Fields(1) = Text2.Text
        .Fields(2) = Text3.Text
        .Fields(3) = Text4.Text
        .Fields(4) = Text5.Text
        .Update
    End With
    MsgBox "DETAILS SUBMITTED"
End Sub

' Next Button
Private Sub Command2_Click()
    Form3.Show
    Me.Hide
End Sub
