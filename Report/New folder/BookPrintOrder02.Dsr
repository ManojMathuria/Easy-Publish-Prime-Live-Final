VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptBookPrintOrder02 
   ClientHeight    =   10350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14565
   OleObjectBlob   =   "BookPrintOrder02.dsx":0000
End
Attribute VB_Name = "rptBookPrintOrder02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public picFile As String
Private Sub Section2_Format(ByVal pFormattingInfo As Object)
    On Error Resume Next
    With Section2.ReportObjects
        Set .Item("Picture1").FormattedPicture = LoadPicture(IIf(FileExist(App.Path & "\Icon\Logo" & CompCode & ".jpg"), App.Path & "\Icon\Logo" & CompCode & ".jpg", ""))
    End With
End Sub

Private Sub Subreport2_Section19_Format(ByVal pFormattingInfo As Object)
    On Error Resume Next
    If FileExist(picFile) Then Set Subreport2_Section19.ReportObjects.Item("Picture1").FormattedPicture = LoadPicture(IIf(FileExist(picFile), picFile, "")) Else Subreport2_Section19.Suppress = True
End Sub



