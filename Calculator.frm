VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculator 
   Caption         =   "UserForm1"
   ClientHeight    =   6912
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5052
   OleObjectBlob   =   "Calculator.frx":0000
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

Private Sub CommandButton_0_Click()
TextBox1.Text = TextBox1.Text + "0"
End Sub

Private Sub CommandButton_1_Click()
TextBox1.Text = TextBox1.Text + "1"
End Sub

Private Sub CommandButton_7_Click()
TextBox1.Text = TextBox1.Text + "7"
End Sub

Private Sub CommandButton_8_Click()
TextBox1.Text = TextBox1.Text + "8"
End Sub

Private Sub CommandButton_9_Click()
TextBox1.Text = TextBox1.Text + "9"
End Sub

Private Sub CommandButton_AC_Click()
TextBox1.Text = ""
End Sub

Private Sub CommandButton_add_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "+"
End Sub

Private Sub CommandButton_Cos_Click()
TextBox1.Text = Math.Cos(a)
End Sub

Private Sub CommandButton_divid_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "/"
End Sub

Private Sub CommandButton_equal_Click()
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

Private Sub CommandButton_log_Click()
TextBox1.Text = Math.Log(a)
End Sub

Private Sub CommandButton_multi_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "*"
End Sub

Private Sub CommandButton_percentis_Click()
b = TextBox1.Text
If operators = "*" Then
c = (a * b) / 100
TextBox1.Text = c
End If
End Sub

Private Sub CommandButton_point_Click()
TextBox1.Text = TextBox1.Text + "."
End Sub

Private Sub CommandButton_power_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "^"
End Sub


Private Sub CommandButton_sin_Click()
TextBox1.Text = Math.Sin(a)
End Sub

Private Sub CommandButton_subt_Click()
a = TextBox1.Text
TextBox1.Text = ""
operators = "-"
End Sub

Private Sub CommandButton_tan_Click()
TextBox1.Text = Math.Tan(a)
End Sub

Private Sub CommandButton00_Click()
TextBox1.Text = TextBox1.Text + "00"
End Sub

Private Sub CommandButton2_Click()
TextBox1.Text = TextBox1.Text + "2"
End Sub

Private Sub CommandButton24_Click()
a = TextBox1.Text
TextBox1.Text = Sqr(a)
End Sub

Private Sub CommandButton3_Click()
TextBox1.Text = TextBox1.Text + "3"
End Sub

Private Sub CommandButton32_Click()
a = Val(TextBox1.Text)
TextBox1.Text = a * WorksheetFunction.Pi()
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

Private Sub CommandButtonLeft_Click()
TextBox1.Text = Left(TextBox1.Value, TextBox1.TextLength - 1)
End Sub
