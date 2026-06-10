VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmATM 
   Caption         =   "ATM Machine"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmATM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmATM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBalance_Click()

Dim ws As Worksheet
Set ws = Sheets("Accounts")

Dim i As Long

For i = 2 To ws.Cells(ws.Rows.Count, 6).End(xlUp).Row

    If Trim(txtPIN.Value) = Trim(CStr(ws.Cells(i, 6).Value)) Then

        MsgBox "Current Balance = Rs. " & ws.Cells(i, 5).Value, vbInformation, "Balance Enquiry"

        Exit Sub

    End If

Next i

MsgBox "Invalid PIN", vbExclamation

End Sub

Private Sub btnExit_Click()

Unload Me

End Sub

Private Sub btnLoginATM_Click()

Dim ws As Worksheet
Set ws = Sheets("Accounts")

Dim i As Long
Dim pinFound As Boolean

pinFound = False

For i = 2 To ws.Cells(ws.Rows.Count, 6).End(xlUp).Row

    If Trim(txtPIN.Value) = Trim(CStr(ws.Cells(i, 6).Value)) Then

        MsgBox "Welcome " & ws.Cells(i, 2).Value

        pinFound = True

        Exit For

    End If

Next i

If pinFound = False Then

    MsgBox "Wrong PIN"

End If

End Sub
Private Sub btnFastCash_Click()

Dim ws As Worksheet
Dim transWs As Worksheet
Dim i As Long
Dim amount As Double
Dim tRow As Long

Set ws = Sheets("Accounts")
Set transWs = Sheets("Transactions")

amount = 500

For i = 2 To ws.Cells(ws.Rows.Count, 6).End(xlUp).Row

    If Trim(txtPIN.Value) = Trim(CStr(ws.Cells(i, 6).Value)) Then

        If Val(ws.Cells(i, 5).Value) < amount Then

            MsgBox "Insufficient Balance", vbExclamation

            Exit Sub

        End If

        ' Deduct Amount
        ws.Cells(i, 5).Value = Val(ws.Cells(i, 5).Value) - amount

        ' Save Transaction
        tRow = transWs.Cells(transWs.Rows.Count, 1).End(xlUp).Row + 1

        transWs.Cells(tRow, 1).Value = "T" & tRow
        transWs.Cells(tRow, 2).Value = ws.Cells(i, 1).Value
        transWs.Cells(tRow, 3).Value = "ATM Withdraw"
        transWs.Cells(tRow, 4).Value = amount
        transWs.Cells(tRow, 5).Value = Now
        transWs.Cells(tRow, 6).Value = ws.Cells(i, 5).Value

        MsgBox "Rs.500 Withdraw Successful!" & vbCrLf & _
               "Remaining Balance = Rs. " & ws.Cells(i, 5).Value, _
               vbInformation, "Fast Cash"

        Exit Sub

    End If

Next i

MsgBox "Invalid PIN", vbExclamation

End Sub
Private Sub btnMini_Click()

Dim ws As Worksheet
Dim accWs As Worksheet
Dim i As Long
Dim accNo As String
Dim pinValue As String

Set ws = Sheets("Transactions")
Set accWs = Sheets("Accounts")

pinValue = Trim(txtPIN.Text)

accNo = ""

' Find Account using PIN
For i = 2 To accWs.Cells(accWs.Rows.Count, 1).End(xlUp).Row

    If Trim(CStr(accWs.Cells(i, 6).Value)) = pinValue Then

        accNo = accWs.Cells(i, 1).Value

        Exit For

    End If

Next i

If accNo = "" Then

    MsgBox "Invalid PIN"

    Exit Sub

End If

' Open Mini Statement Form
frmMiniStatement.Show

' Clear ListBox
frmMiniStatement.lstStatement.Clear

' Load Transactions
For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If Trim(CStr(ws.Cells(i, 2).Value)) = accNo Then

        frmMiniStatement.lstStatement.AddItem _
        ws.Cells(i, 3).Value & " | " & _
        ws.Cells(i, 4).Value & " | " & _
        ws.Cells(i, 5).Value

    End If

Next i

MsgBox "Mini Statement Loaded"
frmMiniStatement.Show
Unload Me
End Sub

