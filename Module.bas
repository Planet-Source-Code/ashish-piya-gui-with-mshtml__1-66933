Attribute VB_Name = "Module1"
Public Enum GridType
    HomeList = 1001
    ListView = 1002
End Enum

Sub Main()
Dim frm As formMain
    Set frm = New formMain
        frm.MyWindow = HomeList
        frm.Show
End Sub

