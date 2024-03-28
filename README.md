Sub DosyaSaydir()

Dim ftpServer As String
Dim ftpUser As String
Dim ftpPassword As String
Dim ftpFolder As String
Dim dosyaListesi As String
Dim dosyaListesiArray() As String
Dim jpegSayacAWDMARK As Integer
Dim xmlSayacAWDMARK As Integer
Dim jpegSayacVAR As Integer
Dim xmlSayacVAR As Integer
Dim jpegSayacYOK As Integer
Dim xmlSayacYOK As Integer


'ftp sunucu bilgileri

ftpServer = "ftp://10.116.184.104/"
ftpUser = "qqww001"
ftpPassword = "3518"
ftpFolder = "/"
 
 'ftp sunucusundan dosya listesini alma
 dosyaListesi = GetFTPDirectoryListing(ftpServer, ftpUser, ftpPassword, ftpFolder)

'Dosya listesini diziye ayirma
dosyaListesiArray = Split(dosyaListesi, vbCrLf)

'AWDMARK(RHLHRRDOOR)klasörü içerisindeki JPG ve XML dosyalarini sayma

For Each dosya In dosyaListesiArray

If InStr(dosya, "/AWDMARK/") > 0 Then
If InStr(dosya, ".jpg") > 0 Then

   jpegSayacAWDMARK = jpegSayacAWDMARK + 1
   
 ElseIf InStr(dosya, ".xml") > 0 Then
 
 xmlSayacAWDMARK = xmlSayacAWDMARK + 1
 
   
  End If
 
 End If
 
 Next dosya
 
 'Var Klasörü Içerisindeki resimleri sayma
 
 For Each dosya In dosyaListesiArray
 
  If InStr(dosya, "/AWDMARK/VAR/") > 0 Then
    
    If InStr(dosya, ".jpg") > 0 Then
      jpgSayacVAR = jpgSayacVAR + 1
      
      ElseIf InStr(dosya, ".xml") > 0 Then
   xmlSayacVAR = xmlSayacVAR + 1
    
    End If
  End If
  
  Next dosya
  
  'YOK klasörü içerisindeki jpg ve xml resimleri
  
  
    
       For Each dosya In dosyaListesiArray
 
  If InStr(dosya, "/AWDMARK/YOK/") > 0 Then
    
    If InStr(dosya, ".jpg") > 0 Then
      jpgSayacYOK = jpgSayacYOK + 1
      
      ElseIf InStr(dosya, ".xml") > 0 Then
   xmlSayacVAR = xmlSayacVAR + 1
    
    End If
  End If
 
 Next dosya
 
   'sonuclari belirtilen hücrelere yazdir
   
   Range("O3").Value = jpgSayacAWDMARK + xmlSayacAWDMARK
   Range("O5").Value = jpgSayacVAR
   Range("O6").Value = xmlSayacVAR
   Range("O7").Value = jpgSayacYOK
   Range("O8").Value = xmlSayacYOK
   
   End Sub
   
   
 Function GetFTPDirectoryListing(ByVal server As String, ByVal user As String, ByVal password As String, ByVal folder As String) As String
 
Dim ftp As Object

'FTP nesnesi olusturma

Set ftp = CreateObject("WinHttp.WinHttpRequest.5.1")

'ftp sunucusuna baglan

ftp.Open "GET", "ftp://" & qww001 & ":" & 3518 & "Q" & server & folder, False
ftp.send

'sonuclari döndürün

GetFTPDirectoryLisng = ftp.responseText



End Function



