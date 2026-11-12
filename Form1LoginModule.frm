Private Sub Command1_Click()
    With Data1.Recordset
        .AddNew
        .Fields(0) = Text1.Text
        .Fields(1) = Text2.Text
        .Update
    End With
    MsgBox "LOGIN SUCCESSFULLY"
    Form2.Show
    Me.Hide
End Sub
