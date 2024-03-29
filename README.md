Dim rowIndex As Integer ' Global olarak tanımlanacak
Sub ResimSay()
   Dim folderPath As String
   Dim objFSO As Object
   ' Ana klasör yolu
   folderPath = "C:\Resimler"
   ' File System Object oluştur
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   ' Excel tablosuna başlık yaz
   Sheets("Sheet1").Range("B3").Value = "Klasör Adi"
   Sheets("Sheet1").Range("C3").Value = "JPG Sayisi"
   Sheets("Sheet1").Range("D3").Value = "XML Sayisi"
   ' Başlangıç satır indeksi
   rowIndex = 4
   ' Klasördeki dosya sayılarını al ve yazdır
   DosyaSayVeYazdir objFSO.GetFolder(folderPath)
   ' Obje belleği temizle
   Set objFSO = Nothing
End Sub
Sub DosyaSayVeYazdir(objFolder As Object)
   Dim subFolder As Object
   Dim objFile As Object
   Dim countJPG As Integer
   Dim countHTML As Integer
   ' Dosya sayıları sıfırla
   countJPG = 0
   countHTML = 0
   ' Sub klasörlerdeki dosya sayılarını al ve yazdır
   For Each objFile In objFolder.Files
       If Right(objFile.Name, 4) = ".jpg" Then
           countJPG = countJPG + 1
       ElseIf Right(objFile.Name, 5) = ".html" Or Right(objFile.Name, 4) = ".xml" Then
           countHTML = countHTML + 1
       End If
   Next objFile
   ' Alt klasörleri gezerek işlem yap
   For Each subFolder In objFolder.SubFolders
       DosyaSayVeYazdir subFolder ' Alt klasörlerin içindeki dosyaları say
   Next subFolder
   ' Klasör adını yaz sadece alt klasörleri gezdikten sonra yazılmalı
   Sheets("Sheet1").Range("B" & rowIndex).Value = objFolder.Path
   ' JPG sayısını yaz
   Sheets("Sheet1").Range("C" & rowIndex).Value = countJPG
   ' HTML sayısını yaz
   Sheets("Sheet1").Range("D" & rowIndex).Value = countHTML
   ' Satır indeksini bir artır
   rowIndex = rowIndex + 1
End Sub
