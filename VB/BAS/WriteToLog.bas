Attribute VB_Name = "Module2"
Sub WriteToLog(What As String, FilePath As String)
If FilePath = "" Then Exit Sub
F% = FreeFile
Open FilePath For Binary Access Write As F%
P$ = What & Chr(10)
Put #1, LOF(1) + 1, P$
Close F%
End Sub
