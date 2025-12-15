<%@ Language=VBScript CodePage=65001 %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Çizim Ekranı</title>
    <style>
        body { text-align: center; font-family: sans-serif; background-color: #f0f0f0; }
        svg { background-color: white; border: 2px solid #333; box-shadow: 0 0 10px rgba(0,0,0,0.2); margin-top: 20px; }
        .nokta { fill: #2c3e50; }
        text { font-size: 10px; fill: #777; }
    </style>
</head>
<body>

    <h2>Veritabanından Okunan Grafiğin Çizimi</h2>
    <a href="index.asp">Yeni Nokta Ekle</a> | <a href="temizle.asp" onclick="return confirm('Tüm noktalar silinecek?');">Çizimi Sıfırla</a>
    <br>

    <%
    ' --- BAĞLANTI AYARLARI (Garantili Yöntem) ---
    Dim conn, rs, dbPath, sql, connectionString
    Set conn = Server.CreateObject("ADODB.Connection")
    dbPath = Server.MapPath("db/veritabani.mdb")

    ' Bağlantıyı açmayı dene
    On Error Resume Next
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath

    ' Eğer hata varsa 64-bit sürücüyü dene
    If Err.Number <> 0 Then
        Err.Clear
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    End If

    ' Hala hata varsa ekrana bas
    If Err.Number <> 0 Then
        Response.Write "<h3>Veritabanı Bağlantı Hatası!</h3>"
        Response.Write "<p>" & Err.Description & "</p>"
        Response.End
    End If
    On Error Goto 0
    ' -------------------------------------------

    Dim pointsString, svgContent, x, y, sira
    
    ' Verileri Sıra Numarasına göre çek
    sql = "SELECT * FROM Noktalar ORDER BY SiraNo ASC"
    Set rs = conn.Execute(sql)
    
    pointsString = ""
    svgContent = ""
    
    ' Kayıt varsa döngüye gir
    If Not rs.EOF Then
        Do While Not rs.EOF
            x = rs("X")
            y = rs("Y")
            sira = rs("SiraNo")
            
            ' Çizgi koordinatlarını birleştir
            pointsString = pointsString & x & "," & y & " "
            
            ' Görsel noktalar ve yazılar
            svgContent = svgContent & "<circle cx='" & x & "' cy='" & y & "' r='5' class='nokta' />"
            svgContent = svgContent & "<text x='" & x + 8 & "' y='" & y & "'>" & sira & ". (" & x & "," & y & ")</text>"
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    %>

    <svg width="800" height="600">
        <defs>
            <pattern id="grid" width="20" height="20" patternUnits="userSpaceOnUse">
                <path d="M 20 0 L 0 0 0 20" fill="none" stroke="#eee" stroke-width="1"/>
            </pattern>
        </defs>
        <rect width="100%" height="100%" fill="url(#grid)" />

        <polyline points="<%= pointsString %>" style="fill:none; stroke:#e74c3c; stroke-width:3;" />
        
        <%= svgContent %>
        
        <% If pointsString = "" Then %>
            <text x="50%" y="50%" text-anchor="middle" font-size="20" fill="red">Henüz çizim yok. Veri ekleyin.</text>
        <% End If %>
    </svg>

</body>
</html>