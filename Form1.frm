VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xo oyny"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8340
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6960
      TabIndex        =   9
      Text            =   "x"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "&Cyk"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF00FF&
      Caption         =   "&Oyun barada"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "&Gaytadan basla"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "&Skory arassala"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "0"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "o:"
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "0"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "x:"
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   240
      Y2              =   5760
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   3600
      X2              =   3600
      Y1              =   240
      Y2              =   5760
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   240
      Y2              =   5760
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   720
      X2              =   6600
      Y1              =   240
      Y2              =   5760
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   6600
      X2              =   720
      Y1              =   240
      Y2              =   5760
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   720
      X2              =   6600
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   720
      X2              =   6600
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   720
      X2              =   6600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image Image30 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":0CCA
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Image Image29 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":1075
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Image Image28 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":1420
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Image Image27 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":17CB
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Image Image26 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":1B76
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Image Image25 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":1F21
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Image Image24 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":22CC
      Top             =   360
      Width           =   1635
   End
   Begin VB.Image Image23 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":2677
      Top             =   360
      Width           =   1635
   End
   Begin VB.Image Image22 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":2A22
      Top             =   360
      Width           =   1635
   End
   Begin VB.Image Image21 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":2DCD
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image20 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":4009
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image19 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":5245
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image18 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":6481
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image17 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":76BD
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image16 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":88F9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image15 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":9B35
      Top             =   360
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image14 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":AD71
      Top             =   360
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image13 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":BFAD
      Top             =   360
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image12 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":D1E9
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image11 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":E3DB
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image10 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":F5CD
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image9 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":107BF
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image8 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":119B1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image7 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":12BA3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image6 
      Height          =   1575
      Left            =   4920
      Picture         =   "Form1.frx":13D95
      Top             =   360
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image5 
      Height          =   1575
      Left            =   2880
      Picture         =   "Form1.frx":14F87
      Top             =   360
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   840
      Picture         =   "Form1.frx":16179
      Top             =   360
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   4680
      Y1              =   360
      Y2              =   5760
   End
   Begin VB.Line Line3 
      X1              =   720
      X2              =   6600
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   6600
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   360
      Y2              =   5760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function h1()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image13.Visible = True
Image22.Visible = False
Image4.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image13.Visible = False
Image22.Visible = False
Image4.Visible = True
Label1.Caption = 1
End If
End Function

Private Sub Command1_Click()
Label3.Caption = 0
Label5.Caption = 0
End Sub

Private Sub Command2_Click()
Dim xo As String
xo = Combo1.Text
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
Image9.Visible = False
Image10.Visible = False
Image11.Visible = False
Image12.Visible = False
Image13.Visible = False
Image14.Visible = False
Image15.Visible = False
Image16.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
Line12.Visible = False
Image22.Visible = True
Image23.Visible = True
Image24.Visible = True
Image25.Visible = True
Image26.Visible = True
Image27.Visible = True
Image28.Visible = True
Image29.Visible = True
Image30.Visible = True
If xo = "x" Then
Label1.Caption = 1
End If
If xo = "o" Then
Label1.Caption = 2
End If
End Sub

Private Sub Command3_Click()
OBD.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Combo1.AddItem "x"
Combo1.AddItem "o"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image22_Click()
h1
c1
c4
c7
End Sub

Private Function h2()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image14.Visible = True
Image23.Visible = False
Image5.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image14.Visible = False
Image23.Visible = False
Image5.Visible = True
Label1.Caption = 1
End If
End Function
Private Function h3()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image15.Visible = True
Image24.Visible = False
Image6.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image15.Visible = False
Image24.Visible = False
Image6.Visible = True
Label1.Caption = 1
End If
End Function

Private Sub Image23_Click()
h2
c1
c5
End Sub

Private Sub Image24_Click()
h3
c1
c6
c8
End Sub

Private Function h4()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image16.Visible = True
Image25.Visible = False
Image7.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image16.Visible = False
Image25.Visible = False
Image7.Visible = True
Label1.Caption = 1
End If
End Function
Private Function h5()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image17.Visible = True
Image26.Visible = False
Image8.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image17.Visible = False
Image26.Visible = False
Image8.Visible = True
Label1.Caption = 1
End If
End Function
Private Function h6()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image18.Visible = True
Image27.Visible = False
Image9.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image18.Visible = False
Image27.Visible = False
Image9.Visible = True
Label1.Caption = 1
End If
End Function

Private Sub Image25_Click()
h4
c2
c4
End Sub

Private Sub Image26_Click()
h5
c2
c5
c7
c8
End Sub

Private Sub Image27_Click()
h6
c2
c6
End Sub

Private Function h7()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image19.Visible = True
Image28.Visible = False
Image10.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image19.Visible = False
Image28.Visible = False
Image10.Visible = True
Label1.Caption = 1
End If
End Function
Private Function h8()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image20.Visible = True
Image29.Visible = False
Image11.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image20.Visible = False
Image29.Visible = False
Image11.Visible = True
Label1.Caption = 1
End If
End Function
Private Function h9()
Dim xo As Byte
xo = Label1.Caption
If xo = 1 Then
Image21.Visible = True
Image30.Visible = False
Image12.Visible = False
Label1.Caption = 2
End If
If xo = 2 Then
Image21.Visible = False
Image30.Visible = False
Image12.Visible = True
Label1.Caption = 1
End If
End Function

Private Sub Image28_Click()
h7
c3
c4
c8
End Sub

Private Sub Image29_Click()
h8
c3
c5
End Sub

Private Sub Image30_Click()
h9
c3
c6
c7
End Sub

Private Function c1()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image13.Visible
x2 = Image14.Visible
x3 = Image15.Visible
o1 = Image4.Visible
o2 = Image5.Visible
o3 = Image6.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line5.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line5.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c2()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image16.Visible
x2 = Image17.Visible
x3 = Image18.Visible
o1 = Image7.Visible
o2 = Image8.Visible
o3 = Image9.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line6.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line6.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c3()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image19.Visible
x2 = Image20.Visible
x3 = Image21.Visible
o1 = Image10.Visible
o2 = Image11.Visible
o3 = Image12.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line7.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line7.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c4()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image13.Visible
x2 = Image16.Visible
x3 = Image19.Visible
o1 = Image4.Visible
o2 = Image7.Visible
o3 = Image10.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line10.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line10.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c5()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image14.Visible
x2 = Image17.Visible
x3 = Image20.Visible
o1 = Image5.Visible
o2 = Image8.Visible
o3 = Image11.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line11.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line11.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c6()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image15.Visible
x2 = Image18.Visible
x3 = Image21.Visible
o1 = Image6.Visible
o2 = Image9.Visible
o3 = Image12.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line12.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line12.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c7()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image13.Visible
x2 = Image17.Visible
x3 = Image21.Visible
o1 = Image4.Visible
o2 = Image8.Visible
o3 = Image12.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line9.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line9.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function
Private Function c8()
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim xh As Byte
Dim oh As Byte
xh = Label3.Caption
oh = Label5.Caption
x1 = Image15.Visible
x2 = Image17.Visible
x3 = Image19.Visible
o1 = Image6.Visible
o2 = Image8.Visible
o3 = Image10.Visible
If x1 = "True" Then
If x2 = "True" Then
If x3 = "True" Then
Line8.Visible = True
Label1.Caption = 0
xh = xh + 1
Label3.Caption = xh
End If
End If
End If
If o1 = "True" Then
If o2 = "True" Then
If o3 = "True" Then
Line8.Visible = True
Label1.Caption = 0
oh = oh + 1
Label5.Caption = oh
End If
End If
End If
End Function

Private Sub Timer1_Timer()
Dim sk1 As Byte
Dim sk2 As Byte
sk1 = Label3.Caption
sk2 = Label5.Caption
If sk1 >= 50 Then
Label3.Caption = 50
End If
If sk2 >= 50 Then
Label5.Caption = 50
End If
End Sub
