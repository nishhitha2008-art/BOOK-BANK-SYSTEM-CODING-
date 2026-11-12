Private Sub Command1_Click()
    With Data1.Recordset
        .AddNew
        .Fields(0) = Text1.Text
        .Fields(1) = Text2.Text
        .Update
    End With
    MsgBox "BOOK DETAILS SAVED"
End Sub

' Exit / Bill Calculation Button
Private Sub Command2_Click()
    MsgBox "BILL AMOUNT Rs.1500/- only. CASH on DELIVERY. BOOKS WILL BE DELIVERED SOON!"
    End
End Sub
