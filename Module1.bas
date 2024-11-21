Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public sql As String

Public Sub koneksi()
Set con = New ADODB.Connection
con.ConnectionString = "PROVIDER=MSDASQL;SERVER=localhost;data source=digitalsignage;user=root;pwd=123456;port=3306"
con.Open
End Sub


