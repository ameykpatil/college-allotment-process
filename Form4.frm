VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Create Account"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      MaxLength       =   15
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      MouseIcon       =   "Form4.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      MaxLength       =   3
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      MaxLength       =   15
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Create My Account"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      MouseIcon       =   "Form4.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " In Block Capital Letters "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form4.frx":02A4
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "College Allotment Process"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seat No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   4440
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strSQL As String
Public sn As Integer
Public pw As String

Private Sub Command1_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox ("No field can be left blank")
End If
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
sn = Text1.Text
Form2.sn2 = sn
pw = Gen_Rand_Password(8)
strSQL = "select * from STUDENT"
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
Do While rs.EOF <> True
If Text1.Text = rs!seatno Then
MsgBox " Your Account Already Exists "
rs.Close
Exit Sub
End If
rs.MoveNext
Loop
rs.AddNew
rs!seatno = Val(Text1.Text)
rs!Lname = Text2.Text
rs!Fname = Text3.Text
rs!Password = pw
rs.Update
Form6.Show
Unload Me
rs.Close
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_LostFocus()
If Val(Text1.Text) > 99 Then
MsgBox ("invalid seat number")
Text1.Text = ""
Text1.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "A" And ch <= "Z" Or KeyAscii = 8) Then
        MsgBox " Block Letters Only"
        KeyAscii = 0
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii = 8) Then
        MsgBox " Block Letters Only"
        KeyAscii = 0
    End If
End Sub

Public Function Gen_Rand_Password(PassLength As Integer) As String
Dim RetVal As String
Dim Max As Integer
Dim Min As Integer
Max = 122
Min = 97
Randomize Timer
For i = 1 To PassLength
    RetVal = RetVal & Chr(Int((Max - Min + 1) * Rnd + Min))
Next i
Gen_Rand_Password = RetVal
End Function
 
