Attribute VB_Name = "SaveasPDF"
Sub SaveSelectedSheetsToPDF()
Dim str As String, myfolder As String, myfile As String
str = "Do you want to save these sheets to a single pdf file?" & Chr(10)
For Each sht In ActiveWindow.SelectedSheets
str = str & sht.Name & Chr(10)
Next sht
answer = MsgBox(str, vbYesNo, "Continue with save?")
If answer = vbNo Then Exit Sub
With Application.FileDialog(msoFileDialogFolderPicker)
.Show
myfolder = .SelectedItems(1) & "\"
End With
myfile = InputBox("Enter filename", "Save as..")
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
myfolder & myfile _
, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=True
End Sub



