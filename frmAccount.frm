VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAccount 
   Caption         =   "Create Account"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmAccount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCreate_Click()

Dim ws As Worksheet
Set ws = Sheets("Accounts")

Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

Dim accNo As String

accNo = "ACC" & lastRow

ws.Cells(lastRow, 1).Value = accNo
ws.Cells(lastRow, 2).Value = txtName.Value
ws.Cells(lastRow, 3).Value = txtPhone.Value
ws.Cells(lastRow, 4).Value = txtAddress.Value
ws.Cells(lastRow, 5).Value = txtDeposit.Value
ws.Cells(lastRow, 6).Value = txtPIN.Value

MsgBox "Account Created Successfully!"
frmTransaction.Show

    Unload Me
End Sub

