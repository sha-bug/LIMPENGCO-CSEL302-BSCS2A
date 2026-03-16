VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CSEL302 
   Caption         =   "Lab Activity"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   OleObjectBlob   =   "CSEL302.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CSEL302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gender As String
Dim Subject As String

Private Sub button1_Click()
Dim x, y As Integer

x = Int(TextBox3.Value)
y = Int(TextBox4.Value)

sum = x + y

Label1.Caption = ("Sum: " & sum)

End Sub


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
Subject = Subject & "CSEL"
End If

End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
Subject = Subject & "CSMC"
End If

End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
Subject = Subject & "ITEC"
End If

End Sub

Private Sub CommandButton1_Click()
Dim ws As Worksheet
Dim lastRow As Long

Set ws = ThisWorkbook.Sheets("Sheet1")

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

ws.Cells(lastRow, 1).Value = TextBox1
ws.Cells(lastRow, 2).Value = TextBox2
ws.Cells(lastRow, 3).Value = ComboBox1
ws.Cells(lastRow, 4).Value = Gender
ws.Cells(lastRow, 5).Value = Subject
End Sub

Private Sub CommandButton2_Click()
Dim x, y As Integer

x = Int(TextBox6.Value)
y = Int(TextBox7.Value)

sum = x + y

MsgBox ("Sum: " & sum)

End Sub

Private Sub CommandButton3_Click()
Dim ws As Worksheet
Dim r As Range

Set ws = ThisWorkbook.Sheets("Sheet1")

For Each r In ws.Range("A2:A100")

If r.Value = TextBox1.Value Then
MsgBox "Record Found"
Exit Sub
End If
Next r

MsgBox "Record Not Found"

End Sub

Private Sub OptionButton1_Click()
If OptionButton1.Value = True Then
Gender = "Male"
End If

End Sub

Private Sub OptionButton2_Click()
If OptionButton2.Value = True Then
Gender = "Female"
End If

End Sub


Private Sub UserForm_Initialize()
ComboBox1.AddItem ("BSCS")
ComboBox1.AddItem ("BSIT")

End Sub
