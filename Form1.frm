VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hotel Management system"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   15510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   13560
      TabIndex        =   30
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "unbook / clear"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5400
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   8400
      TabIndex        =   28
      Top             =   1200
      Width           =   6855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2520
      TabIndex        =   27
      Top             =   2280
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1800
      TabIndex        =   26
      Top             =   2280
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1200
      TabIndex        =   25
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Submit/Booking"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1560
      TabIndex        =   23
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   22
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1200
      TabIndex        =   21
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   20
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Room 15"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Room 14"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Room 13"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Room 12"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Room 11"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Room 10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Room 9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Room 8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Room 7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Room 6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Room 5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Room 4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Room 3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Room 2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Room 1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      Height          =   2295
      Left            =   11520
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Unbook or clear the details"
      Height          =   210
      Left            =   12000
      TabIndex        =   34
      Top             =   4080
      Width           =   1920
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Enter custumber details for booking"
      Height          =   210
      Left            =   360
      TabIndex        =   33
      Top             =   840
      Width           =   2550
   End
   Begin VB.Shape Shape4 
      Height          =   6135
      Left            =   8280
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7095
   End
   Begin VB.Shape Shape3 
      Height          =   6375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Hotel Management system"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   32
      Top             =   0
      Width           =   4860
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Room no."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   31
      Top             =   4560
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Room No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mobile no."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Addharno."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   585
   End
   Begin VB.Shape Shape2 
      Height          =   6135
      Left            =   5160
      Top             =   720
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   6135
      Left            =   3720
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim roomButtons(1 To 15) As CommandButton
Dim bookings(1 To 15) As String
Const AvailableColor As Long = vbGreen
Const BookedColor As Long = vbRed

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 15
Set roomButtons(i) = Controls("Command" & i)
bookings(i) = ""
roomButtons(i).BackColor = AvailableColor
Next i
End Sub

Private Sub Command16_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "All fields must be filled out.", vbExclamation
Exit Sub
End If

Dim gender As String
If Option1.Value = True Then
gender = "F"
ElseIf Option2.Value = True Then
gender = "M"
ElseIf Option3.Value = True Then
gender = "O"
Else
MsgBox "Please select a gender.", vbExclamation
Exit Sub
End If

Dim roomNumber As Integer
roomNumber = CInt(Text4.Text)
If roomNumber < 1 Or roomNumber > 15 Then
MsgBox "Invalid room number.", vbExclamation
Exit Sub
End If

If bookings(roomNumber) <> "" Then
MsgBox "This room is already booked.", vbExclamation
Exit Sub
End If

Dim bookingInfo As String
bookingInfo = "Name: " & Text1.Text & vbCrLf & _
"Gender: " & gender & vbCrLf & _
"Address: " & Text2.Text & vbCrLf & _
"Addhar: " & Text3.Text & vbCrLf & _
"Room No: " & roomNumber & vbCrLf & _
"Date: " & Date
bookings(roomNumber) = bookingInfo
List1.AddItem bookingInfo

roomButtons(roomNumber).Enabled = False
roomButtons(roomNumber).BackColor = BookedColor

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False

MsgBox "Room booked successfully!", vbInformation
End Sub

Private Sub ShowRoomDetails(roomNumber As Integer)
If bookings(roomNumber) = "" Then
MsgBox "Room " & roomNumber & " is available.", vbInformation
Else
MsgBox "Room " & roomNumber & " is booked. Details:" & vbCrLf & bookings(roomNumber), vbInformation
End If
End Sub

Private Sub Command1_Click()
Text4.Text = "1"
ShowRoomDetails 1
End Sub

Private Sub Command2_Click()
Text4.Text = "2"
ShowRoomDetails 2
End Sub

Private Sub Command3_Click()
Text4.Text = "3"
ShowRoomDetails 3
End Sub

Private Sub Command4_Click()
Text4.Text = "4"
ShowRoomDetails 4
End Sub

Private Sub Command5_Click()
Text4.Text = "5"
ShowRoomDetails 5
End Sub

Private Sub Command6_Click()
Text4.Text = "6"
ShowRoomDetails 6
End Sub

Private Sub Command7_Click()
Text4.Text = "7"
ShowRoomDetails 7
End Sub

Private Sub Command8_Click()
Text4.Text = "8"
ShowRoomDetails 8
End Sub

Private Sub Command9_Click()
Text4.Text = "9"
ShowRoomDetails 9
End Sub

Private Sub Command10_Click()
Text4.Text = "10"
ShowRoomDetails 10
End Sub

Private Sub Command11_Click()
Text4.Text = "11"
ShowRoomDetails 11
End Sub

Private Sub Command12_Click()
Text4.Text = "12"
ShowRoomDetails 12
End Sub

Private Sub Command13_Click()
Text4.Text = "13"
ShowRoomDetails 13
End Sub

Private Sub Command14_Click()
Text4.Text = "14"
ShowRoomDetails 14
End Sub

Private Sub Command15_Click()
Text4.Text = "15"
ShowRoomDetails 15
End Sub

Private Sub Command17_Click()
If Text5.Text = "" Then
MsgBox "Please enter a room number to unbook.", vbExclamation
Exit Sub
End If

Dim roomNumber As Integer
roomNumber = CInt(Text5.Text)
If roomNumber < 1 Or roomNumber > 15 Then
MsgBox "Invalid room number.", vbExclamation
Exit Sub
End If

If bookings(roomNumber) = "" Then
MsgBox "This room is already available.", vbExclamation
Exit Sub
End If

bookings(roomNumber) = ""
roomButtons(roomNumber).Enabled = True
roomButtons(roomNumber).BackColor = AvailableColor

Dim i As Integer
For i = 0 To List1.ListCount - 1
If InStr(List1.List(i), "Room No: " & roomNumber) > 0 Then
List1.RemoveItem i
Exit For
End If
Next i

MsgBox "Room unbooked successfully!", vbInformation
End Sub


