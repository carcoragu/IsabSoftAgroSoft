<%
Const DB_RUTA = "C:\presu\backup\bdganaderia\isabsoft.accdb"
Const DB_SISTEMA = "IsabSoft"
Const DB_VERSION = "2.0"

Function AbrirConexion()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_RUTA & ";Persist Security Info=False;"
    Set AbrirConexion = conn
End Function
%>
