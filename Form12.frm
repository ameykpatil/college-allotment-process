VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Admin Operations"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form12"
   MouseIcon       =   "Form12.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080C0FF&
      Caption         =   "Execute"
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
      Left            =   6480
      MouseIcon       =   "Form12.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   8040
      Width           =   10095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "Stream Info."
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
      Left            =   6120
      MouseIcon       =   "Form12.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "College Info."
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
      Left            =   6120
      MouseIcon       =   "Form12.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Student Info."
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
      Left            =   6120
      MouseIcon       =   "Form12.frx":0548
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Results"
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
      Left            =   6120
      MouseIcon       =   "Form12.frx":069A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Declare Results"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      MouseIcon       =   "Form12.frx":07EC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset CAP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      MouseIcon       =   "Form12.frx":093E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Log Out"
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
      MouseIcon       =   "Form12.frx":0A90
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form12.frx":0BE2
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
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim strSQL As String
Public f As Integer
Public query As String

Private Sub Command1_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command2_Click()
strSQL = "select * from result where name = 'admin' "
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
rs!flag = 0
rs.Update

strSQL = "Delete from student"
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
oconn.Execute strSQL
oconn.Close

Command3.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
MsgBox "CAP Is Reset For Next Round"
End Sub

Private Sub Command3_Click()

strSQL = "select * from student where submit=1 order by ( pmark + cmark + mmark) desc"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText

Do While rs.EOF <> True
    strSQL = "select * from optfor,preferences,offers where optfor.coursecode = preferences.coursecode and preferences.colid=offers.colid and preferences.strid=offers.strid and optfor.seatno = '" & rs!seatno & "' order by prefno "
    Set rs1 = New ADODB.Recordset
    Set oconn = New ADODB.Connection
    oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
    rs1.CursorType = adOpenDynamic
    rs1.CursorLocation = adUseClient
    rs1.LockType = adLockOptimistic
    rs1.Open strSQL, oconn, , , adCmdText
    
    Do While rs1.EOF <> True
        If rs1!vacant > 0 Then
        rs!colid = rs1!colid
        rs!strid = rs1!strid
        rs1!vacant = rs1!vacant - 1
        rs.Update
        rs1.Update
        Exit Do
        Else
        rs1.MoveNext
        End If
    Loop
    
rs.MoveNext
Loop

strSQL = "select * from result where name = 'admin' "
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
rs!flag = 1
rs.Update
Command2.Enabled = True
Command4.Enabled = True
Command3.Enabled = False
MsgBox "Results Are Declared"
End Sub

Private Sub Command4_Click()
f = 0
Form13.Show
Unload Me
End Sub

Private Sub Command5_Click()
f = 1
Form13.Show
Unload Me
End Sub

Private Sub Command6_Click()
f = 2
Form13.Show
Unload Me
End Sub

Private Sub Command7_Click()
f = 3
Form13.Show
Unload Me
End Sub

Private Sub Command8_Click()
Dim splt() As String
splt = Split(Text1.Text)
If splt(0) <> "select" Then
MsgBox ("Invalid Query" & vbCrLf & "Only Select Query Is Allowed")
Exit Sub
End If
query = Text1.Text
f = 4
Form13.Show
Unload Me
End Sub

Private Sub Form_Load()
strSQL = "select * from result where name = 'admin' "
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
If rs!flag = 0 Then
Command2.Enabled = False
Command4.Enabled = False
Else
Command3.Enabled = False
End If
End Sub
