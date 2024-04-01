Dim rowIndex As Integer ' Global olarak tanimlanacak
Sub ResimSay()
   Dim folderPath As String
   Dim objFSO As Object
   ' Ana klasör yolu
   folderPath = "C:\Resimler"
   ' File System Object olustur
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   ' Excel tablosuna baslik yaz
   Sheets("Sheet1").Range("B3").Value = "Klasör Adi"
   Sheets("Sheet1").Range("C3").Value = "JPG Sayisi"
   Sheets("Sheet1").Range("D3").Value = "XML Sayisi"
   ' Baslangiç satir indeksi
   rowIndex = 4
   ' Klasördeki dosya sayilarini al ve yazdir
   DosyaSayVeYazdir objFSO.GetFolder(folderPath)
   ' Obje bellegi temizle
   Set objFSO = Nothing
End Sub
Sub DosyaSayVeYazdir(objFolder As Object)
   Dim subFolder As Object
   Dim objFile As Object
   Dim countJPG As Integer
   Dim countXML As Integer
   ' Dosya sayilari sifirla
   countJPG = 0
   countHTML = 0
   ' Sub klasörlerdeki dosya sayilarini al ve yazdir
   For Each objFile In objFolder.Files
       If Right(objFile.Name, 4) = ".jpg" Then
           countJPG = countJPG + 1
       ElseIf Right(objFile.Name, 5) = ".html" Or Right(objFile.Name, 4) = ".xml" Then
           countHTML = countHTML + 1
       End If
   Next objFile
   ' Alt klasörleri gezerek islem yap
   For Each subFolder In objFolder.SubFolders
       DosyaSayVeYazdir subFolder ' Alt klasörlerin içindeki dosyalari say
   Next subFolder
   ' Klasör adini yaz sadece alt klasörleri gezdikten sonra yazilmali
   Sheets("Sheet1").Range("B" & rowIndex).Value = objFolder.Path
   ' JPG sayisini yaz
   Sheets("Sheet1").Range("C" & rowIndex).Value = countJPG
   ' HTML sayisini yaz
   Sheets("Sheet1").Range("D" & rowIndex).Value = countHTML
   ' Satir indeksini bir artir
   rowIndex = rowIndex + 1
End Sub
