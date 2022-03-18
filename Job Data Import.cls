VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim imgPath As String
Sub dataFromExcel()

Dim excelApp As Excel.Application
Dim eBook As Excel.Workbook
Dim eSheet As Excel.Worksheet
Dim ctl As Word.ContentControl
Dim objPic As Object
Dim exportData As Excel.Range
Dim ttlField As Excel.Range
Dim attPath As String
Dim i As Integer
Dim pctCompl As Integer
Dim imgPath, wbkPath As String

'progress counter
i = 0

'get the path to the job export folder
MsgBox Prompt:="Choose the excel file exported with your job. (zipped export files must be extracted first).", Buttons:=vbOKOnly

wbkPath = SelectFile
    
If wbkPath = "" Then
    Exit Sub
End If
    
Set excelApp = CreateObject("excel.application")

'open the excel file
Set eBook = excelApp.Workbooks.Open(wbkPath)

excelApp.Visible = True
Set eSheet = eBook.Worksheets(1)
Set exportData = eSheet.Range("A5:I65")

'Open progress control form
UserForm.Show

For Each ctl In Word.ActiveDocument.ContentControls
    'loop through all content controls and populate them
    If ctl.Type = wdContentControlRichText Then
        Set ttlField = exportData.Find(ctl.Title)
        ctl.Range.Text = ttlField.Offset(0, 1).Value 'looks to the right of the matching step name for the value, content control titles must match the corresponding step name in the spreadsheet
    ElseIf ctl.Type = wdContentControlPicture And ctl.Title <> "" Then
        
        imgPath = GetImagePath(eBook.Path & "\attachments", ctl)
        
        ' delete existing image in the control if there is one
        If ctl.Range.InlineShapes.Count > 0 Then
            ctl.Range.InlineShapes(1).Delete
        End If
        
            If imgPath <> "Image Not Found" Then
                Set objPic = ActiveDocument.InlineShapes.AddPicture(FileName:=imgPath, linktofile:=True, Range:=ctl.Range)
                With objPic
                    .ScaleHeight = 40
                    .ScaleWidth = 40
                End With
            End If
            
    End If
    'Progress bar
    i = i + 1
    pctCompl = (i / Word.ActiveDocument.ContentControls.Count) * 100
    Progress pctCompl
Next
eBook.Close
UserForm.Hide
MsgBox Prompt:="Data import complete", Buttons:=vbOKOnly
End Sub
Sub Progress(pctCompl As Integer)

    UserForm.Text.Caption = pctCompl & "% Completed"
    UserForm.Bar.Width = pctCompl * 2
    
    DoEvents
    
End Sub
Function CheckDirectory(attPath) As Boolean

Dim CheckDir As String

CheckDir = Dir(attPath, vbDirectory)

If CheckDir <> "" Then
    CheckDirectory = True
Else
    CheckDirectory = False
End If

End Function

Function GetImagePath(attPath As String, ctl As Word.ContentControl) As String
    
Dim FileName As String

FileName = Dir(attPath & "\*.jp*g")

Do While FileName <> ""
    If ctl.Title <> "" And FileName Like "*" & ctl.Title & "*" Then
        GetImagePath = attPath & "\" & FileName
        Exit Function
    End If
    
        FileName = Dir()
Loop
            
    GetImagePath = "Image Not Found"
    
End Function

Function SelectFile() As String
Dim sFile As String

    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then
            sFile = .SelectedItems(1)
        End If
    End With
    
    If sFile <> "" Then
        SelectFile = sFile
    End If
    
End Function
Sub PrintControlTitles()

Dim ctl As Word.ContentControl
Dim objPic As Object
Dim i As Integer
Dim pctCompl As Integer

For Each ctl In Word.ActiveDocument.ContentControls
    'loop through all content controls and populate them
    Debug.Print ctl.ID & " - " & ctl.Title & " - " & ctl.Type
Next
End Sub
