VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "123-45-6789"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ssn_checker As SSNChecker
Attribute ssn_checker.VB_VarHelpID = -1

Private Sub cmdTest_Click()
    ssn_checker.CheckSsn txtValue.Text
End Sub


Private Sub Form_Load()
    Set ssn_checker = New SSNChecker
End Sub


Private Sub ssn_checker_BadSsn(ByVal bad_ssn As String)
    MsgBox bad_ssn & " has an incorrect format"
End Sub


