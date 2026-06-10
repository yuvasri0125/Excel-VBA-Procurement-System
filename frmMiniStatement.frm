VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMiniStatement 
   Caption         =   "Mini Statement"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmMiniStatement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMiniStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnShow_Click()

Dim ws As Worksheet
Set ws = Sheets("Transactions")

Dim accNo As String
Dim i As Long
Dim lastRow As Long

accNo = Trim(txtAccountNo.Value)

' Clear old statement
lstStatement.Clear

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Load Transactions
For i = 2 To lastRow

    If Trim(ws.Cells(i, 2).Value) = accNo Then

        lstStatement.AddItem ws.Cells(i, 3).Value

        lstStatement.List(lstStatement.ListCount - 1, 1) = ws.Cells(i, 4).Value

        lstStatement.List(lstStatement.ListCount - 1, 2) = ws.Cells(i, 5).Value

    End If

Next i

If lstStatement.ListCount = 0 Then

    MsgBox "No Transactions Found"

Else

    MsgBox "Mini Statement Loaded"

End If

End Sub



Private Sub btnExportPDF_Click()

Dim wsTemp As Worksheet
Dim i As Long
Dim filePath As String

' Check if statement loaded
If lstStatement.ListCount = 0 Then

    MsgBox "First Load Mini Statement"

    Exit Sub

End If

' Create Temporary Sheet
Set wsTemp = Worksheets.Add

' Heading
wsTemp.Cells(1, 1).Value = "Transaction Type"
wsTemp.Cells(1, 2).Value = "Amount"
wsTemp.Cells(1, 3).Value = "Date"

' Copy ListBox Data To Sheet
For i = 0 To lstStatement.ListCount - 1

    wsTemp.Cells(i + 2, 1).Value = lstStatement.List(i, 0)
    wsTemp.Cells(i + 2, 2).Value = lstStatement.List(i, 1)
    wsTemp.Cells(i + 2, 3).Value = lstStatement.List(i, 2)

Next i

' Autofit Columns
wsTemp.Columns("A:C").AutoFit

' PDF Path
filePath = ThisWorkbook.Path & "\MiniStatement.pdf"

' Export PDF
wsTemp.ExportAsFixedFormat _
Type:=xlTypePDF, _
Filename:=filePath, _
Quality:=xlQualityStandard

' Delete Temp Sheet
Application.DisplayAlerts = False
wsTemp.Delete
Application.DisplayAlerts = True

MsgBox "PDF Exported Successfully!"

' Open PDF
Shell "explorer.exe " & filePath, vbNormalFocus

End Sub

