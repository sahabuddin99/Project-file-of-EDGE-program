VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   6324
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5160
   OleObjectBlob   =   "Calculator12.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim c As Double
Dim operators As String


Private Sub CommandButton1_Click()
TextBox1.Text = TextBox1.Text + "1"
End Sub

Private Sub CommandButton2_Click()
TextBox1.Text = TextBox1.Text + "2"
End Sub

Private Sub CommandButton22_Click()

End Sub

Private Sub CommandButton3_Click()
TextBox1.Text = TextBox1.Text + "3"
End Sub

Private Sub CommandButton4_Click()
TextBox1.Text = TextBox1.Text + "4"
End Sub

Private Sub CommandButton5_Click()
TextBox1.Text = TextBox1.Text + "5"
End Sub

Private Sub CommandButton6_Click()
TextBox1.Text = TextBox1.Text + "6"
End Sub

Private Sub CommandButton7_Click()
TextBox1.Text = TextBox1.Text + "7"
End Sub

Private Sub CommandButton8_Click()
TextBox1.Text = TextBox1.Text + "8"
End Sub

Private Sub CommandButton9_Click()
TextBox1.Text = TextBox1.Text + "9"
End Sub

Private Sub CommandButtonAc_Click()
TextBox1.Text = ""
End Sub

Private Sub CommandButtonadd_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "+"
End Sub

Private Sub CommandButtonCos_Click()
TextBox1.Text = Math.Cos(a)
End Sub

Private Sub CommandButtondivid_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "/"
End Sub

Private Sub CommandButtonDot_Click()
TextBox1.Text = TextBox1.Text + "."
End Sub

Private Sub CommandButtonDZero_Click()
TextBox1.Text = TextBox1.Text + "00"
End Sub

Private Sub CommandButtonequal_Click()
b = TextBox1.Text
If operators = "+" Then
c = a + b
TextBox1.Text = c
End If
b = TextBox1.Text
If operators = "-" Then
c = a - b
TextBox1.Text = c
End If
b = TextBox1.Text
If operators = "*" Then
c = a * b
TextBox1.Text = c
End If
b = TextBox1.Text
If operators = "/" Then
c = a / b
TextBox1.Text = c
End If
b = TextBox1.Text
If operators = "^" Then
c = a ^ b
TextBox1.Text = c
End If
End Sub

Private Sub CommandButtonLeft_Click()
TextBox1.Text = Left(TextBox1.Value, TextBox1.TextLength - 1)
End Sub

Private Sub CommandButtonminus_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "-"
End Sub

Private Sub CommandButtonmulti_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "*"
End Sub

Private Sub CommandButtonpercetage_Click()
b = TextBox1.Text
If operators = "*" Then
c = (a * b) / 100
TextBox1.Text = c
End If
End Sub

Private Sub CommandButtonpower_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "^"
End Sub

Private Sub CommandButtonSin_Click()
TextBox1.Text = Math.Sin(a)
End Sub

Private Sub CommandButtonTan_Click()
TextBox1.Text = Math.Tan(a)
End Sub

Private Sub CommandButtonZero_Click()
TextBox1.Text = TextBox1.Text + "0"
End Sub

Private Sub UserForm_Click()

End Sub
