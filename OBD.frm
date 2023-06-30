VERSION 5.00
Begin VB.Form OBD 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Oyun Barada"
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3945
   Icon            =   "OBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   120
      Picture         =   "OBD.frx":0CCA
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "OBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub
