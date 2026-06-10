VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInterest 
   Caption         =   "Interest Calculator"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmInterest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalculate_Click()

Dim P As Double
Dim R As Double
Dim T As Double
Dim Interest As Double
Dim Total As Double

P = Val(txtPrincipal.Text)
R = Val(txtRate.Text)
T = Val(txtTime.Text)

Interest = (P * R * T) / 100
Total = P + Interest

MsgBox "Interest = Rs. " & Interest & vbCrLf & _
       "Total Amount = Rs. " & Total, vbInformation, "Result"

End Sub

Private Sub btnClear_Click()

txtPrincipal.Text = ""
txtRate.Text = ""
txtTime.Text = ""

txtPrincipal.SetFocus

End Sub

Private Sub UserForm_Click()

End Sub
