VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   971
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLoadPrev 
      Caption         =   "Load previous Population"
      Height          =   255
      Left            =   9960
      TabIndex        =   5
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox tCode 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "fMain.frx":0000
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "go"
      Height          =   615
      Left            =   9960
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label Label2 
      Caption         =   "Task: Stay at a given distance from each other."
      Height          =   615
      Left            =   11880
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Best Code:"
      Height          =   255
      Left            =   9960
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()


    If fMain.chkLoadPrev.Value = vbChecked Then GE.LoadPopulation "POP.txt"


    MainLOOP

End Sub

Private Sub Form_Load()
    Randomize Timer



    pHDC = PIC.HDC
    MaxX = PIC.Width
    MaxY = PIC.Height


    INITW

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

