<%@ Language=VBScript CodePage=65001 %>
<%
' Hata yakalamayı aç
On Error Resume Next

Dim conn, dbPath, connectionString
Set conn = Server.CreateObject("ADODB.Connection")

' Veritabanı yolunu al
dbPath = Server.MapPath("db/veritabani.MDB")

' 1. Sürücü Denemesi: Standart JET (32-bit sistemler ve BabyWeb için genelde bu çalışır)
connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
conn.Open connectionString

' Eğer 1. yöntem hata verdiyse 2. yöntemi (ACE - 64-bit) dene
If Err.Number <> 0 Then
    Err.Clear
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    conn.Open connectionString
End If

' Hala hata varsa işlemi durdur ve hatayı göster
If Err.Number <> 0 Then
    Response.Write "<div style='background-color:red; color:white; padding:10px;'>"
    Response.Write "<h3>VERİTABANI BAĞLANTISI KURULAMADI!</h3>"
    Response.Write "<p><b>Hata Kodu:</b> " & Err.Number & "</p>"
    Response.Write "<p><b>Hata Mesajı:</b> " & Err.Description & "</p>"
    Response.Write "<p><b>Dosya Yolu:</b> " & dbPath & "</p>"
    Response.Write "</div>"
    Response.End
End If
%>