VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Adding Pictures to Text Boxes in Visual Basic"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   DrawWidth       =   5
   LinkTopic       =   "Form1"
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   480
      Left            =   1530
      TabIndex        =   3
      Top             =   3420
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   60
      Width           =   8655
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   105
      Picture         =   "Form1.frx":00EF
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   4035
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2340
      Top             =   4605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Picture"
      Height          =   480
      Left            =   165
      TabIndex        =   0
      Top             =   3420
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Command1_Click()
Dim I, J As Long
Dim Col As Long
Dim DC As Long

    DC = GetDC(Text1.hwnd)
    For I = 1 To p1.Width - 1
        For J = 1 To p1.Height - 1
            Col = GetPixel(p1.hdc, I, J)
            SetPixel DC, 10 + I * 2, 10 + J * 2, Col
        Next
    Next
    
End Sub

Private Sub Command2_Click()
    End
    
End Sub

Private Sub Form_Load()
    Command1_Click
    
End Sub

Private Sub Text1_Change()
    Command1_Click
    
End Sub
