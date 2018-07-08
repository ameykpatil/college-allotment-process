VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Result"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Information Module"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      MouseIcon       =   "Form11.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Go"
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
      Left            =   6600
      MouseIcon       =   "Form11.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
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
      MouseIcon       =   "Form11.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
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
      MouseIcon       =   "Form11.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
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
      TabIndex        =   1
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit The College And Confirm Your Admission"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   8160
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Congratulations !!!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Try Again In The Next Round"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Seat No."
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
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
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
      Picture         =   "Form11.frx":0548
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim strSQL As String
Dim flag As Boolean
Public out As Integer

Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
strSQL = "select * from Student where Student.submit = 1"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
flag = False
Do While rs.EOF <> True
If Val(Text1.Text) = rs!seatno Then
flag = True
Exit Do
End If
rs.MoveNext
Loop
If flag = False Then
MsgBox "Invalid Entry"
Text1.Text = ""
Exit Sub
End If

If rs!colid <> "" Or rs!strid <> "" Then
strSQL = "select * from Student,College,Stream where Student.seatno = '" & Val(Text1.Text) & "' and Student.colid=College.colid and Student.strid=Stream.strid"
Set rs1 = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs1.CursorType = adOpenDynamic
rs1.CursorLocation = adUseClient
rs1.LockType = adLockOptimistic
rs1.Open strSQL, oconn, , , adCmdText
Label4.Caption = "College Allotted Is"
Label5.Caption = rs1!colname & " " & rs1!strname
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True

Else
Label4.Caption = "No College Is Allotted"
Label3.Visible = True
Label4.Visible = True

End If
End Sub

Private Sub Command3_Click()
Unload Me
End
End Sub

Private Sub Command4_Click()
Form10.Show
Unload Me
End Sub
