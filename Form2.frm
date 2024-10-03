VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "PASSWORD :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim username As String
Dim password As String
    
username = "user"
password = "password"
    
If Text1.Text = username And Text2.Text = password Then
MsgBox "Login successful!"
Form1.Show
Form2.Hide
        
Else
MsgBox "Invalid username or password. Please try again."
Text2.Text = ""
Text2.SetFocus
End If
End Sub


