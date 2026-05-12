<%
' ============================================================
' IsabSoft - Configuración de base de datos
' INSTRUCCIONES:
'   1. Copiá este archivo y renombralo a: config.asp
'   2. Cambiá DB_RUTA por la ruta real de tu base de datos
'   3. config.asp NO se sube a GitHub (está en .gitignore)
' ============================================================

Const DB_RUTA = "C:\presu\backup\bdganaderia\isabsoft.accdb"
Const DB_SISTEMA = "IsabSoft"
Const DB_VERSION = "2.0"

' Función de conexión reutilizable
Function AbrirConexion()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_RUTA & ";Persist Security Info=False;"
    Set AbrirConexion = conn
End Function
%>
