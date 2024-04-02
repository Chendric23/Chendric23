

Private Sub BoxAciklama1_Change()

End Sub

Private Sub ButtonAra_Click()

 On Error GoTo bitir
 
 
aranan = InputBox("Parça No Giriniz", "Arama Islemi")

Range("B:B").Find(aranan).Select


sil_satir = ActiveCell.Row

  TextBoxPRTNO.Value = Worksheets("Anasayfa").Cells(sil_satir, 2)
  TextBoxPRTNM.Value = Worksheets("Anasayfa").Cells(sil_satir, 3)
  TextBoxNSCTN.Value = Worksheets("Anasayfa").Cells(sil_satir, 4)
  TextBoxSHKBT.Value = Worksheets("Anasayfa").Cells(sil_satir, 5)
  ComboBoxMDL.Value = Worksheets("Anasayfa").Cells(sil_satir, 6)
  TextBoxÖRBRK.Value = Worksheets("Anasayfa").Cells(sil_satir, 7)
  TextBoxKYW.Value = Worksheets("Anasayfa").Cells(sil_satir, 8)
  TextBoxFRMT.Value = Worksheets("Anasayfa").Cells(sil_satir, 9)
  TextBoxTP.Value = Worksheets("Anasayfa").Cells(sil_satir, 10)
  TextBoxPCT.Value = Worksheets("Anasayfa").Cells(sil_satir, 11)
  BoxLen1.Value = Worksheets("Anasayfa").Cells(sil_satir, 12)
  BoxLen2.Value = Worksheets("Anasayfa").Cells(sil_satir, 13)
  BoxLen3.Value = Worksheets("Anasayfa").Cells(sil_satir, 14)
  BoxAciklama1.Value = Worksheets("Anasayfa").Cells(sil_satir, 15)
  
  Exit Sub

bitir:  MsgBox "Aranan Kayit Bulunamadi", , "HATA"
  

End Sub

Private Sub ButtonGüncelle_Click()

On Error GoTo bitir

     aranan = TextBoxPRTNO.Value
     Range("B:B").Find(aranan).Select
     
     
     
     guncelle = ActiveCell.Row
     
  Worksheets("Anasayfa").Cells(guncelle, 2) = TextBoxPRTNO.Value
  Worksheets("Anasayfa").Cells(guncelle, 3) = TextBoxPRTNM.Value
  Worksheets("Anasayfa").Cells(guncelle, 4) = TextBoxNSCTN.Value
  Worksheets("Anasayfa").Cells(guncelle, 5) = TextBoxSHKBT.Value
  Worksheets("Anasayfa").Cells(guncelle, 6) = ComboBoxMDL.Value
  Worksheets("Anasayfa").Cells(guncelle, 7) = TextBoxÖRBRK.Value
  Worksheets("Anasayfa").Cells(guncelle, 8) = TextBoxKYW.Value
  Worksheets("Anasayfa").Cells(guncelle, 9) = TextBoxFRMT.Value
  Worksheets("Anasayfa").Cells(guncelle, 10) = TextBoxTP.Value
  Worksheets("Anasayfa").Cells(guncelle, 11) = TextBoxPCT.Value
  Worksheets("Anasayfa").Cells(guncelle, 12) = BoxLen1.Value
  Worksheets("Anasayfa").Cells(guncelle, 13) = BoxLen2.Value
  Worksheets("Anasayfa").Cells(guncelle, 14) = BoxLen3.Value
  Worksheets("Anasayfa").Cells(guncelle, 15) = BoxAciklama1.Value
  
  
bitir:






End Sub

Private Sub ButtonKaydet_Click()

If TextBoxPRTNO <> "" And TextBoxPRTNM <> "" And TextBoxNSCTN <> "" And TextBoxSHKBT <> "" And ComboBoxMDL <> "" And TextBoxKYW <> "" And TextBoxFRMT <> "" And TextBoxTP <> "" And TextBoxPCT <> "" Then




Sonsatir = WorksheetFunction.CountA(Worksheets("Anasayfa").Range("B:B")) + 1

If Sonsatir = 2 Then

  Worksheets("Anasayfa").Cells(Sonsatir, 1) = 1

  Worksheets("Anasayfa").Cells(Sonsatir, 2) = TextBoxPRTNO.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 3) = TextBoxPRTNM.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 4) = TextBoxNSCTN.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 5) = TextBoxSHKBT.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 6) = ComboBoxMDL.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 7) = TextBoxÖRBRK.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 8) = TextBoxKYW.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 9) = TextBoxFRMT.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 10) = TextBoxTP.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 11) = TextBoxPCT.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 12) = BoxLen1.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 13) = BoxLen2.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 14) = BoxLen3.Value
  Worksheets("Anasayfa").Cells(Sonsatir, 15) = BoxAciklama1.Value


Else

Worksheets("Anasayfa").Cells(Sonsatir, 1) = Worksheets("Anasayfa").Cells(Sonsatir - 1, 1) + 1

    Worksheets("Anasayfa").Cells(Sonsatir, 2) = TextBoxPRTNO.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 3) = TextBoxPRTNM.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 4) = TextBoxNSCTN.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 5) = TextBoxSHKBT.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 6) = ComboBoxMDL.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 7) = TextBoxÖRBRK.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 8) = TextBoxKYW.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 9) = TextBoxFRMT.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 10) = TextBoxTP.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 11) = TextBoxPCT.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 12) = BoxLen1.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 13) = BoxLen2.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 14) = BoxLen3.Value
    Worksheets("Anasayfa").Cells(Sonsatir, 15) = BoxAciklama1.Value

End If
Else

MsgBox " Giris Alanlarinin Tümünü Doldurunuz...", , "HATA !"

End If


End Sub

Private Sub ButtonSil_Click()

   If TextBoxPRTNM.Value <> "" Then
 
  Rows(ActiveCell.Row).Delete
  
  TextBoxPRTNO.Value = ""
  TextBoxPRTNM.Value = ""
  TextBoxNSCTN.Value = ""
  TextBoxSHKBT.Value = ""
  ComboBoxMDL.Value = ""
  TextBoxÖRBRK.Value = ""
  TextBoxKYW.Value = ""
  TextBoxFRMT.Value = ""
  TextBoxTP.Value = ""
  TextBoxPCT.Value = ""
  BoxLen1.Value = ""
  BoxLen2.Value = ""
  BoxLen3.Value = ""
  BoxAciklama1.Value = ""
 
Else
  MsgBox "Öncelikle Arama Islemi Yapmaniz Gerekmektedir", , "HATA"
  
  End If
  


End Sub

Private Sub CommandButtonKapat_Click()

Unload Me


End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

ListBox1.ColumnCount = 11
ListBox1.RowSource = "Anasayfa!kayiit"
ListBox1.ColumnWidths = "68;62;160;50;50;40;150;90;170;30;40;20"


End Sub
