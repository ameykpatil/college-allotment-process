VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Login"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
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
      MouseIcon       =   "Form5.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   2055
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
      MouseIcon       =   "Form5.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "OK"
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
      Left            =   6720
      MouseIcon       =   "Form5.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   2535
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
      IMEMode         =   3  'DISABLE
      Left            =   8160
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5160
      Width           =   2655
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
      Left            =   8160
      MaxLength       =   15
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invalid UserName or Password. Login denied."
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
      Left            =   4560
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   5520
      TabIndex        =   4
      Top             =   5160
      Width           =   2175
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
      Left            =   5520
      TabIndex        =   3
      Top             =   4080
      Width           =   2175
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
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form5.frx":03F6
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim flag As Boolean
Public sn1 As Integer

Private Sub Command1_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("No field can be left blank")
End If
If Text1.Text <> "" And Text2.Text <> "" Then
strSQL = "select * from STUDENT"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenStatic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
flag = False
Do While rs.EOF <> True
If rs!seatno = Text1.Text And rs!Password = Text2.Text Then
flag = True
Exit Do
Else
rs.MoveNext
End If
Loop
If flag = False Then
Label4.Visible = True
Label4.Caption = "Invalid UserName/Password. Login denied"
Text2.Text = ""
Text2.SetFocus
End If
If flag = True Then
sn1 = Text1.Text
Form2.sn2 = sn1
Form7.Show
Unload Me
End If
End If
rs.Close
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ch = Chr$(KeyAscii)
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

