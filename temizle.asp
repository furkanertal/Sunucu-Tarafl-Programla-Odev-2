<%@ Language=VBScript CodePage=65001 %>
<%
' 1. Bağlantı Değişkenlerini Tanımla
Dim conn, dbPath
Set conn = Server.CreateObject("ADODB.Connection")
dbPath = Server.MapPath("db/veritabani.mdb")

' 2. Bağlantıyı Aç (Garantili Yöntem)
On Error Resume Next
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath

' Jet sürücüsü çalışmazsa ACE (64-bit) dene
If Err.Number <> 0 Then
    Err.Clear
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
End If

' Hala hata varsa işlemi durdur
If Err.Number <> 0 Then
    Response.Write "<h3>Hata! Veritabanına bağlanılamadı.</h3>"
    Response.Write "<p>" & Err.Description & "</p>"
    Response.End
End If
On Error Goto 0

' 3. Tabloyu Temizle (Tüm kayıtları sil)
conn.Execute "DELETE FROM Noktalar"

' 4. Bağlantıyı Kapat
conn.Close
Set conn = Nothing

' 5. Çizim sayfasına geri dön (Boş halini görmek için)
Response.Redirect "ciz.asp"
%>