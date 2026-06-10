VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransaction 
   Caption         =   "Deposit & Withdraw System"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmTransaction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnATM_Click()

frmATM.Show

Unload Me

End Sub
Private Sub btnInterest_Click()

frmInterest.Show

End Sub

Private Sub btnSearch_Click()

Dim ws As Worksheet
Dim accNo As String
Dim i As Long

Set ws = Sheets("Accounts")

accNo = Trim(txtAccountNo.Text)

For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If Trim(ws.Cells(i, 1).Value) = accNo Then

        MsgBox "Account Found: " & ws.Cells(i, 2).Value & vbCrLf & _
               "Balance: Rs. " & ws.Cells(i, 5).Value

        Exit Sub

    End If

Next i

MsgBox "Account Not Found"

End Sub
Private Sub btnDeposit_Click()

Dim ws As Worksheet
Dim transWs As Worksheet
Dim accNo As String
Dim amount As Double
Dim i As Long
Dim tRow As Long

Set ws = Sheets("Accounts")
Set transWs = Sheets("Transactions")

accNo = Trim(txtAccountNo.Text)
amount = Val(txtAmount.Text)

For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If Trim(ws.Cells(i, 1).Value) = accNo Then

        ws.Cells(i, 5).Value = Val(ws.Cells(i, 5).Value) + amount

        tRow = transWs.Cells(transWs.Rows.Count, 1).End(xlUp).Row + 1

        transWs.Cells(tRow, 1).Value = "T" & tRow
        transWs.Cells(tRow, 2).Value = accNo
        transWs.Cells(tRow, 3).Value = "Deposit"
        transWs.Cells(tRow, 4).Value = amount
        transWs.Cells(tRow, 5).Value = Now
        transWs.Cells(tRow, 6).Value = ws.Cells(i, 5).Value

        MsgBox "Deposit Successful!" & vbCrLf & _
               "New Balance: Rs. " & ws.Cells(i, 5).Value

        Exit Sub

    End If

Next i

MsgBox "Account Not Found"

End Sub
Private Sub btnWithdraw_Click()

Dim ws As Worksheet
Dim transWs As Worksheet
Dim accNo As String
Dim amount As Double
Dim i As Long
Dim tRow As Long

Set ws = Sheets("Accounts")
Set transWs = Sheets("Transactions")

accNo = Trim(txtAccountNo.Text)
amount = Val(txtAmount.Text)

For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If Trim(ws.Cells(i, 1).Value) = accNo Then

        If Val(ws.Cells(i, 5).Value) < amount Then
            MsgBox "Insufficient Balance"
            Exit Sub
        End If

        ws.Cells(i, 5).Value = Val(ws.Cells(i, 5).Value) - amount

        tRow = transWs.Cells(transWs.Rows.Count, 1).End(xlUp).Row + 1

        transWs.Cells(tRow, 1).Value = "T" & tRow
        transWs.Cells(tRow, 2).Value = accNo
        transWs.Cells(tRow, 3).Value = "Withdraw"
        transWs.Cells(tRow, 4).Value = amount
        transWs.Cells(tRow, 5).Value = Now
        transWs.Cells(tRow, 6).Value = ws.Cells(i, 5).Value

        MsgBox "Withdraw Successful!" & vbCrLf & _
               "Remaining Balance: Rs. " & ws.Cells(i, 5).Value

        Exit Sub

    End If

Next i

MsgBox "Account Not Found"

End Sub

