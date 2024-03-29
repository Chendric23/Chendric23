Sub JPG_Saydir()
 
Dim anaKlasor As String
Dim anaKlasorler As Variant
Dim i As Integer
 
anaKlasorler = Array("AWDMARK", "BACKDOORGNKAMERA")
 
For i = LBound(anaKlasorler) To UBound(anaKlasorler)
 
anaKlasor = "C:\Resimler\" & anaKlasorler(i) & "\"
 
 
Cells(3 + i, 4).Value = DosyaSay(anaKlasor, "*.JPG")
 
If anaKlasorler(i) = "AWDMARK" Then
 
Cells(6, 4).Value = DosyaSay(anaKlasor, "*.HTML")
 
  ElseIf anaKlasorler(i) = "BACKDOORGNKAMERA" Then
Cells(7, 4).Value = DosyaSay(anaKlasor, ".HTML")
 
Cells(7, 6).Value = DosyaSay(anaKlasor & "1AD+KAPALI\", "*.JPG")
Cells(7, 9).Value = DosyaSay(anaKlasor & "2_Adet\", "*.JPG")
 
   End If
Next i
 
 
End Sub
 
Function DosyaSay(ByVal klasor As String, ByVal filtre As String) As Integer
 
Dim fs As Object
Dim dosyalar As Object
Dim dosya As Object
Dim altKlasor As Object
Dim sayac As Integer
 
 
Set fs = CreateObject("Scripting.FileSystemObject")
 
On Error Resume Next
 
On Error GoTo 0
 
If Not dosyalar Is Nothing Then
 
For Each dosya In dosyalar
  If LCase(Right(dosya.Name, 4)) = Right(filtre, 4) Then
  sayac = sayac + i
   End If
  Next dosya
  End If
On Error Resume Next
 
For Each altKlasor In fs.GetFolder(klasor).SubFolders
  sayac = sayac + DosyaSay(altKlasor.Path, filtre)
Next altKlasor
On Error GoTo 0
DosyaSay = sayac


 
End Function

