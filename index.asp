<%@ Language=VBScript CodePage=65001 %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Koordinat Ekle</title>
</head>
<body style="font-family:sans-serif; text-align:center; padding-top:50px;">
    <h2>Yeni Koordinat Girişi</h2>
    <form action="ekle.asp" method="post" style="border:1px solid #ccc; display:inline-block; padding:20px;">
        <div>
            <label>Sıra No:</label><br>
            <input type="number" name="SiraNo" required value="1">
        </div>
        <br>
        <div>
            <label>X Koordinatı (0-800):</label><br>
            <input type="number" name="X" required>
        </div>
        <br>
        <div>
            <label>Y Koordinatı (0-600):</label><br>
            <input type="number" name="Y" required>
        </div>
        <br>
        <button type="submit" style="padding:10px 20px; background:blue; color:white;">KAYDET</button>
    </form>
    <br><br>
    <a href="ciz.asp">>> ÇİZİMİ GÖR <<</a>
</body>
</html>