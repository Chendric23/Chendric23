Sub ResimSay()
   Dim folderPath As String
   Dim AWDMARK_VAR_Path As String
   Dim AWDMARK_YOK_Path As String
   Dim BACKDOORGNKAMERA_1AD_Path As String
   Dim BACKDOORGNKAMERA_2AD_Path As String
   ' Ana klasör yolu
   folderPath = "C:\resimler\"
   ' AWDMARK klasörlerinin yolları
   AWDMARK_VAR_Path = folderPath & "AWDMARK\VAR\"
   AWDMARK_YOK_Path = folderPath & "AWDMARK\YOK\"
   ' BACKDOORGNKAMERA klasörlerinin yolları
   BACKDOORGNKAMERA_1AD_Path = folderPath & "BACKDOORGNKAMERA\1AD+KAPALI\"
   BACKDOORGNKAMERA_2AD_Path = folderPath & "BACKDOORGNKAMERA\2_Adet\"
   ' AWDMARK VAR klasöründeki jpg ve html sayısını say
   Range("C4").Value = DosyaSay(AWDMARK_VAR_Path, "*.jpg")
   Range("F4").Value = DosyaSay(AWDMARK_VAR_Path, "*.html")
   ' AWDMARK YOK klasöründeki jpg ve html sayısını say
   Range("D4").Value = DosyaSay(AWDMARK_YOK_Path, "*.jpg")
   Range("G4").Value = DosyaSay(AWDMARK_YOK_Path, "*.html")
   ' BACKDOORGNKAMERA 1AD+KAPALI klasöründeki jpg ve html sayısını say
   Range("C6").Value = DosyaSay(BACKDOORGNKAMERA_1AD_Path, "*.jpg")
   Range("F6").Value = DosyaSay(BACKDOORGNKAMERA_1AD_Path, "*.html")
   ' BACKDOORGNKAMERA 2_Adet klasöründeki jpg ve html sayısını say
   Range("D6").Value = DosyaSay(BACKDOORGNKAMERA_2AD_Path, "*.jpg")
   Range("G6").Value = DosyaSay(BACKDOORGNKAMERA_2AD_Path, "*.html")
End Sub
Function DosyaSay(folderPath As String, filePattern As String) As Integer
   
        Dim objFSO As Object
        Dim objFolder As Object
        Dim objFile As Object
        Dim count As Integer
        ' File System Object oluştur
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        ' Klasörü al
        On Error Resume Next
        Set objFolder = objFSO.GetFolder(folderPath)
        On Error GoTo 0
        If objFolder Is Nothing Then
            MsgBox "Klasör bulunamadi: " & folderPath, vbExclamation
            Exit Function
   End If
   ' Dosyaları say
   count = objFolder.Files.count
   ' Obje belleği temizle
   Set objFolder = Nothing
   Set objFile = Nothing
   Set objFSO = Nothing
   ' Sayıyı döndür
   DosyaSay = count
End Function

