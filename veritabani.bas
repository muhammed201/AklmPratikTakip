Attribute VB_Name = "veritabani"
Global veri As Database
Global tablo As Recordset
Global harf As String
Global secilisatir As Integer



Sub veri_ac(X1 As Boolean, X2 As Boolean)
Set veri = Workspaces(0).OpenDatabase(App.Path & "\veri_Backup.mdb", X1, X2)
End Sub


Sub tablo_ac(sql As String)
Set tablo = veri.OpenRecordset(sql)
End Sub



