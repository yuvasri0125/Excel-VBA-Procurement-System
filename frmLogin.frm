VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Bank Login System"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogin_Click()

If txtUsername.Value = "admin" And txtPassword.Value = "1234" Then

    MsgBox "Login Successful!"

    frmAccount.Show

    Unload Me

Else

    MsgBox "Invalid Username or Password"

End If

End Sub
