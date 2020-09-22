VERSION 5.00
Begin VB.Form frmCountCharAppearance 
   Caption         =   "Character count"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match Case"
      Height          =   195
      Left            =   1980
      TabIndex        =   7
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton btnCount 
      Caption         =   "Count them!"
      Height          =   315
      Left            =   3180
      TabIndex        =   6
      Top             =   1020
      Width           =   1275
   End
   Begin VB.TextBox txtNumber 
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txtChar 
      Height          =   315
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   4
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   300
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Character:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Text:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmCountCharAppearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCount_Click()
  ' Let's check it out!!!
  txtNumber.Text = CountCharAppearance(txtText.Text, txtChar.Text, chkMatchCase.Value)
End Sub

Private Function CountCharAppearance(strText As String, strChar As String, Optional blnMatchCase As Boolean = False) As Integer
  ' strText is the text to search through
  ' strChar is the character to search for
  ' blnMatchCase indicates matching of upper/lowercase
  Dim Parts

  If Not blnMatchCase Then
    ' No Case matching: convert everything to uppercase
    strText = UCase$(strText)
    strChar = UCase$(txtChar)
  End If

  If Len(strText) > 0 And Len(strChar) = 1 Then
    Parts = Split(strText, strChar)
    CountCharAppearance = UBound(Parts)
  End If
End Function

Private Sub txtChar_GotFocus()
  ' Focus on txtChar => select the text
  SelectText txtChar
End Sub

Private Sub txtText_GotFocus()
  ' Focus on txtText => select the text
  SelectText txtText
End Sub

Public Sub SelectText(ControlName As Control)
  ' Select the text in a textbox
  If TypeName(ControlName) = "TextBox" Then
    If Len(ControlName) > 0 Then
      ControlName.SelStart = 0
      ControlName.SelLength = Len(ControlName.Text)
      Exit Sub
    End If
  End If
End Sub
