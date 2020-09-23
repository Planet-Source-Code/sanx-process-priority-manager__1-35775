VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Process"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the name of the process you wish to prioritise. Don't include the path, and remember include the .exe extension."
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAdd.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

frmAdd.Hide

End Sub

Private Sub cmdCancel_Click()

txtName.Text = ""
frmAdd.Hide

End Sub

Private Sub Form_Activate()

txtName.SetFocus

End Sub

