Attribute VB_Name = "TekMarkaAnaliz"
'=====================================================================
'  TEK MARKA NOZUL PERFORMANSI
'  Makro adý: TekMarkaPerformans
'  1 marka sorar -> o markanýn pivotunu oluþturur:
'  Ort. Aldýðý Süre, Ort. Aldýðý Döküm Sayýsý, Kullaným Adedi
'  Filtreler: Marka (tek seçim) ve Deðiþim Sebebi
'  NOZUL 2026 sayfasý açýkken çalýþtýrýn.
'=====================================================================
Option Explicit

' True : sadece "Deðiþim Sebebi" girilmiþ satýrlar alýnýr
Private Const SADECE_DEGISENLER As Boolean = True

Private mWb As Workbook
Private mN As Long
Private mMarka() As String
Private mSebep() As String
Private mSure() As Double
Private mSureOK() As Boolean
Private mDokum() As Double
Private mDokumOK() As Boolean

Public Sub TekMarkaPerformans()
    Dim dMarka As Object, i As Long, k As String, marka As String
    Dim wsV As Worksheet, wsP As Worksheet
    Dim pc As PivotCache, pt As PivotTable, df As PivotField
    Dim outArr() As Variant

    If Not VeriYukle() Then Exit Sub

    Set dMarka = CreateObject("Scripting.Dictionary")
    For i = 1 To mN
        k = Norm(mMarka(i))
        If Not dMarka.Exists(k) Then dMarka.Add k, mMarka(i)
    Next i
    marka = MarkaSec(dMarka, "Performansýna bakýlacak marka", "")
    If marka = "" Then Exit Sub

    On Error GoTo Hata
    Application.ScreenUpdating = False
    Set wsV = SayfaHazirla(mWb, "PerfVeri")
    Set wsP = SayfaHazirla(mWb, "MarkaPerformans")

    ReDim outArr(1 To mN + 1, 1 To 4)
    outArr(1, 1) = "Marka"
    outArr(1, 2) = "Deðiþim Sebebi"
    outArr(1, 3) = "Aldýðý Süre"
    outArr(1, 4) = "Aldýðý Döküm Sayýsý"
    For i = 1 To mN
        outArr(i + 1, 1) = dMarka.Item(Norm(mMarka(i)))
        outArr(i + 1, 2) = IIf(mSebep(i) = "", "(Boþ)", mSebep(i))
        If mSureOK(i) Then outArr(i + 1, 3) = mSure(i)
        If mDokumOK(i) Then outArr(i + 1, 4) = mDokum(i)
    Next i
    wsV.Range("A1").Resize(mN + 1, 4).Value = outArr
    wsV.Range("C2").Resize(mN, 1).NumberFormat = "[h]:mm:ss"

    Set pc = mWb.PivotCaches.Create(SourceType:=xlDatabase, _
                                    SourceData:=wsV.Range("A1").Resize(mN + 1, 4))
    Set pt = pc.CreatePivotTable(TableDestination:=wsP.Range("A8"), TableName:="MarkaPerformans")

    With pt
        .ManualUpdate = True
        .PivotFields("Marka").Orientation = xlPageField
        .PivotFields("Marka").EnableMultiplePageItems = False
        .PivotFields("Deðiþim Sebebi").Orientation = xlPageField
        .PivotFields("Deðiþim Sebebi").EnableMultiplePageItems = True

        Set df = .AddDataField(.PivotFields("Aldýðý Süre"), "Ort. Aldýðý Süre", xlAverage)
        df.NumberFormat = "[h]:mm:ss"
        Set df = .AddDataField(.PivotFields("Aldýðý Döküm Sayýsý"), "Ort. Aldýðý Döküm Sayýsý", xlAverage)
        df.NumberFormat = "0.00"
        Set df = .AddDataField(.PivotFields("Marka"), "Kullaným Adedi", xlCount)
        df.NumberFormat = "0"
        .DataPivotField.Orientation = xlColumnField
        .ManualUpdate = False
        .PivotFields("Marka").CurrentPage = marka
    End With

    With wsP
        .Range("A1").Value = "NOZUL PERFORMANSI"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Marka deðiþtirmek için 'Marka' filtresini, sebep seçmek için 'Deðiþim Sebebi' filtresini kullanýn."
        .Range("A2").Font.Italic = True
        .Columns("A:D").ColumnWidth = 26
    End With

    wsV.Visible = xlSheetHidden
    wsP.Activate
    Application.ScreenUpdating = True
    Exit Sub
Hata:
    Application.ScreenUpdating = True
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

'=====================================================================
'  YARDIMCI FONKSÝYONLAR
'=====================================================================
Private Function VeriYukle() As Boolean
    Dim lo As ListObject
    Dim hdr As Variant, dat As Variant
    Dim cM As Long, cS As Long, cSure As Long, cDok As Long
    Dim r As Long, nR As Long
    Dim mk As String, sb As String, eksik As String

    Set lo = TabloBul()
    If lo Is Nothing Then
        MsgBox "'Marka' ve 'Aldýðý Süre' sütunlarýný içeren tablo bulunamadý." & vbCrLf & _
               "NOZUL 2026 sayfasýný açýp makroyu tekrar çalýþtýrýn.", vbExclamation
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        MsgBox "Tablo boþ.", vbExclamation
        Exit Function
    End If

    Set mWb = lo.Parent.Parent
    hdr = lo.HeaderRowRange.Value2
    dat = lo.DataBodyRange.Value2

    cM = KolonBul(hdr, "MARKA")
    cS = KolonBul(hdr, "DEGISIMSEBEBI")
    cSure = KolonBul(hdr, "ALDIGISURE")
    cDok = KolonBul(hdr, "ALDIGIDOKUMSAYISI")
    If cM = 0 Then eksik = eksik & "Marka, "
    If cS = 0 Then eksik = eksik & "Deðiþim Sebebi, "
    If cSure = 0 Then eksik = eksik & "Aldýðý Süre, "
    If cDok = 0 Then eksik = eksik & "Aldýðý Döküm Sayýsý, "
    If eksik <> "" Then
        MsgBox "Þu sütun(lar) bulunamadý: " & Left$(eksik, Len(eksik) - 2), vbExclamation
        Exit Function
    End If

    nR = UBound(dat, 1)
    ReDim mMarka(1 To nR): ReDim mSebep(1 To nR)
    ReDim mSure(1 To nR): ReDim mSureOK(1 To nR)
    ReDim mDokum(1 To nR): ReDim mDokumOK(1 To nR)
    mN = 0
    For r = 1 To nR
        mk = Temiz(dat(r, cM))
        sb = Temiz(dat(r, cS))
        If mk <> "" Then
            If (Not SADECE_DEGISENLER) Or sb <> "" Then
                mN = mN + 1
                mMarka(mN) = mk
                mSebep(mN) = sb
                mSure(mN) = SureCevir(dat(r, cSure), mSureOK(mN))
                mDokum(mN) = SayiCevir(dat(r, cDok), mDokumOK(mN))
            End If
        End If
    Next r

    If mN = 0 Then
        MsgBox "Uygun veri satýrý bulunamadý.", vbExclamation
        Exit Function
    End If
    VeriYukle = True
End Function

Private Function TabloBul() As ListObject
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    For Each lo In ActiveSheet.ListObjects
        If TabloUygun(lo) Then Set TabloBul = lo: Exit Function
    Next lo
    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If TabloUygun(lo) Then Set TabloBul = lo: Exit Function
        Next lo
    Next ws
End Function

Private Function TabloUygun(lo As ListObject) As Boolean
    Dim hdr As Variant
    hdr = lo.HeaderRowRange.Value2
    TabloUygun = (KolonBul(hdr, "MARKA") > 0 And KolonBul(hdr, "ALDIGISURE") > 0)
End Function

Private Function KolonBul(hdr As Variant, anahtar As String) As Long
    Dim j As Long
    For j = LBound(hdr, 2) To UBound(hdr, 2)
        If Norm(hdr(1, j)) = anahtar Then KolonBul = j: Exit Function
    Next j
    For j = LBound(hdr, 2) To UBound(hdr, 2)
        If InStr(1, Norm(hdr(1, j)), anahtar, vbBinaryCompare) > 0 Then KolonBul = j: Exit Function
    Next j
End Function

' Türkçe karakter, boþluk ve satýr sonlarýný temizleyip büyük harfe çevirir
Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = UCase$(CStr(v))
    s = Replace(s, vbCr, ""): s = Replace(s, vbLf, "")
    s = Replace(s, " ", ""): s = Replace(s, Chr(160), "")
    s = Replace(s, ChrW(775), "")                                  ' birleþik nokta
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I") ' Ý ý
    s = Replace(s, ChrW(286), "G"): s = Replace(s, ChrW(287), "G") ' Ð ð
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U") ' Ü ü
    s = Replace(s, ChrW(350), "S"): s = Replace(s, ChrW(351), "S") ' Þ þ
    s = Replace(s, ChrW(214), "O"): s = Replace(s, ChrW(246), "O") ' Ö ö
    s = Replace(s, ChrW(199), "C"): s = Replace(s, ChrW(231), "C") ' Ç ç
    s = Replace(s, "i", "I")
    Norm = s
End Function

Private Function Temiz(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Temiz = Trim$(Replace(CStr(v), Chr(160), " "))
End Function

' Süre: sayý (Excel saat deðeri) veya "03:43:00" gibi metin olabilir
Private Function SureCevir(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim s As String, p As Variant, h As Double, m As Double, sn As Double, d As Double
    ok = False
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(v)
        If s = "" Then Exit Function
        p = Split(s, ":")
        If UBound(p) < 1 Then Exit Function
        h = Val(p(0)): m = Val(p(1))
        If UBound(p) >= 2 Then sn = Val(Replace(p(2), ",", "."))
        d = (h + m / 60 + sn / 3600) / 24
    ElseIf IsNumeric(v) Then
        d = CDbl(v)
    Else
        Exit Function
    End If
    ok = (d > 0)
    SureCevir = d
End Function

' Döküm sayýsý: sayý veya "1,5" gibi metin olabilir
Private Function SayiCevir(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim s As String, d As Double
    ok = False
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(Replace(v, ",", "."))
        If s = "" Then Exit Function
        If Not (s Like "*#*") Then Exit Function
        d = Val(s)
    ElseIf IsNumeric(v) Then
        d = CDbl(v)
    Else
        Exit Function
    End If
    ok = (d >= 0)
    SayiCevir = d
End Function

Private Function MarkaSec(d As Object, baslik As String, varsayilan As String) As String
    Dim keys As Variant, i As Long, liste As String, defIdx As Long
    Dim cevap As String, kv As String
    keys = d.keys
    kv = Norm(varsayilan)
    For i = 0 To d.Count - 1
        liste = liste & (i + 1) & " - " & d.Item(keys(i)) & vbCrLf
        If defIdx = 0 And keys(i) = kv Then defIdx = i + 1
    Next i
    If defIdx = 0 Then
        For i = 0 To d.Count - 1
            If InStr(keys(i), kv) > 0 Or InStr(kv, keys(i)) > 0 Then defIdx = i + 1: Exit For
        Next i
    End If
    cevap = InputBox(baslik & " numarasýný yazýn:" & vbCrLf & vbCrLf & liste, _
                     "Marka Seçimi", IIf(defIdx > 0, CStr(defIdx), ""))
    cevap = Trim$(cevap)
    If cevap = "" Then Exit Function
    If IsNumeric(cevap) Then
        i = CLng(cevap)
        If i >= 1 And i <= d.Count Then MarkaSec = d.Item(keys(i - 1)): Exit Function
    End If
    If d.Exists(Norm(cevap)) Then MarkaSec = d.Item(Norm(cevap)): Exit Function
    MsgBox "Geçersiz seçim: " & cevap, vbExclamation
End Function

Private Function SayfaHazirla(wb As Workbook, ad As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(ad)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Visible = xlSheetVisible
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = ad
    Set SayfaHazirla = ws
End Function
