VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
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
      Left            =   6600
      MouseIcon       =   "Form1.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Administrator"
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
      Left            =   6480
      TabIndex        =   1
      Top             =   5760
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Student"
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
      Left            =   6480
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
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
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form1.frx":02A4
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Which Type of User are you?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   2
      Top             =   3360
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strSQL As String

Private Sub Command2_Click()
If Option1.Value = True Then
strSQL = "select * from result where name = 'admin' "
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
If rs!flag = 1 Then
Form11.Show
Unload Me
Else
Form2.Show
Unload Me
End If
rs.Close
Else
Form3.Show
Unload Me
End If
End Sub

Private Sub Command1_Click()
Unload Me
End
End Sub

