<%@ Language=VBScript CodePage=65001 %>
<%
' 1. Değişkenleri Tanımla
Dim sira, x, y, sql, conn, dbPath

' 2. Formdan gelen verileri al
sira = Request.Form("SiraNo")
x = Request.Form("X")
y = Request.Form("Y")

' 3. Veritabanı Bağlantısını Burada Kur (Include kullanmadan)
Set conn = Server.CreateObject("ADODB.Connection")
dbPath = Server.MapPath("db/veritabani.mdb")

' Bağlantıyı açmayı dene
On Error Resume Next
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath

If Err.Number <> 0 Then
    ' Jet çalışmazsa ACE dene (64-bit için)
    Err.Clear
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
End If

' Eğer hala bağlantı yoksa hatayı bas
If Err.Number <> 0 Then
    Response.Write "<h3>Veritabanına bağlanılamadı!</h3>"
    Response.Write "Hata: " & Err.Description
    Response.End
End If
On Error Goto 0

' 4. Kaydetme İşlemi
If sira <> "" And x <> "" And y <> "" Then
    sql = "INSERT INTO Noktalar (SiraNo, X, Y) VALUES (" & sira & ", " & x & ", " & y & ")"
    
    ' Sorguyu çalıştır
    conn.Execute sql
    
    ' Temizlik yap ve yönlendir
    conn.Close
    Set conn = Nothing
    Response.Redirect "index.asp"
Else
    Response.Write "Lütfen tüm alanları doldurun."
End If
%>