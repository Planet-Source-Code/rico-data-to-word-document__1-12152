VERSION 5.00
Begin VB.Form frmData 
   Caption         =   "Personal Data"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Word"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtTelephone 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Address"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Telephone"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "City"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdClose_Click()
Word.Application.Quit wdDoNotSaveChanges
cmdClose.Visible = False
End Sub

Private Sub cmdPrint_Click()
Dim Name As String
Dim Address As String
Dim City As String
Dim Telephone As String

'data for variables
Name = txtName.Text
Address = txtAddress.Text
City = txtCity.Text
Telephone = txtTelephone.Text

'call module
Call Printen(Name, Address, City, Telephone)
cmdClose.Visible = True
End Sub

Private Sub cmdReset_Click()
txtName.Text = ""
txtAddress.Text = ""
txtCity.Text = ""
txtTelephone.Text = ""
End Sub

Private Sub Form_Load()
cmdClose.Visible = False
End Sub
