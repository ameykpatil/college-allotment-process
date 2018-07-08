VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Application Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "Save"
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
      Left            =   9120
      MouseIcon       =   "Form8.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "Submit"
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
      Left            =   11640
      MouseIcon       =   "Form8.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MouseIcon       =   "Form8.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "v"
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
      MouseIcon       =   "Form8.frx":03F6
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8880
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "^"
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
      MouseIcon       =   "Form8.frx":0548
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8160
      Width           =   615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Columns         =   3
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      ItemData        =   "Form8.frx":069A
      Left            =   1440
      List            =   "Form8.frx":069C
      TabIndex        =   34
      Top             =   7560
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MouseIcon       =   "Form8.frx":069E
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6840
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      ItemData        =   "Form8.frx":07F0
      Left            =   3960
      List            =   "Form8.frx":07F2
      TabIndex        =   32
      Text            =   "Select Stream"
      Top             =   6120
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      ItemData        =   "Form8.frx":07F4
      Left            =   1440
      List            =   "Form8.frx":07F6
      TabIndex        =   31
      Text            =   "Select College"
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   27
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   25
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   22
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      MaxLength       =   20
      TabIndex        =   19
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      MaxLength       =   10
      TabIndex        =   18
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
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
      MouseIcon       =   "Form8.frx":07F8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preferences"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   38
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   360
      X2              =   15000
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   30
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   29
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   28
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maths"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   26
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chemistry"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   24
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Physics"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   23
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Out Of"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   21
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   20
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   15000
      X2              =   15000
      Y1              =   10320
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   360
      X2              =   360
      Y1              =   2280
      Y2              =   10320
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   17
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email Id"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   16
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " Surname"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   11880
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seat No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   9600
      X2              =   15000
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   360
      X2              =   5760
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Form"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form8.frx":094A
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
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim strSQL As String
Dim TestArray() As String
Dim CourseArray(5) As String
Dim m As Integer

Private Sub Combo1_Click()
Combo2.Clear
Combo2.Refresh
Combo2.Text = "Select Stream"
End Sub

Private Sub Combo2_GotFocus()
If Combo1.Text = "Select College" Then
MsgBox "select college first"
Else
Combo2.Clear
strSQL = "select * from Stream,College,Offers where Stream.Strid = Offers.Strid and Offers.Colid = College.Colid and College.Colname  = '" & Combo1.Text & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
While rs.EOF <> True
    Combo2.AddItem rs!strname
    rs.MoveNext
Wend
End If
End Sub

Private Sub Combo2_LostFocus()
If Combo2.Text = "" Then
Combo2.Text = "Select Stream"
End If
End Sub

Private Sub Command1_Click()
If Combo1.Text = "Select College" Then
MsgBox "Select College and Stream "
Else
If Combo2.Text = "Select Stream" Then
MsgBox "Select Stream "
Else
If List1.ListCount = 5 Then
MsgBox "5 preferences only"
Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.List(i) = (Combo1.Text & " " & Combo2.Text) Then
MsgBox "preference already exist"
Exit Sub
End If
Next i
List1.AddItem (Combo1.Text & " " & Combo2.Text)
End If
End If
End Sub

Private Sub Command2_Click()
If List1.SelCount = 0 Then Exit Sub
If List1.ListIndex = 0 Then Exit Sub
Dim z As String
Dim w As Integer
Dim u As Integer
Dim v(5) As String
z = List1.Text
w = List1.ListIndex
u = List1.ListCount - 1
For i = 0 To List1.ListCount - 1
v(i) = List1.List(i)
Next i
v(w) = v(w - 1)
v(w - 1) = z
List1.Clear
For i = 0 To u
List1.AddItem (v(i))
Next i
List1.ListIndex = w - 1
End Sub

Private Sub Command3_Click()
Form7.Show
Unload Me
End Sub

Private Sub Command4_Click()
If List1.SelCount = 0 Then Exit Sub
If List1.ListIndex = List1.ListCount - 1 Then Exit Sub
Dim z As String
Dim w As Integer
Dim u As Integer
Dim v(5) As String
z = List1.Text
w = List1.ListIndex
u = List1.ListCount - 1
For i = 0 To List1.ListCount - 1
v(i) = List1.List(i)
Next i
v(w) = v(w + 1)
v(w + 1) = z
List1.Clear
For i = 0 To u
List1.AddItem (v(i))
Next i
List1.ListIndex = w + 1
End Sub

Private Sub Command5_Click()
Call ListBoxRemSel(List1)
End Sub

Private Sub Command6_Click()

strSQL = "select * from Student where Seatno = '" & Form2.sn2 & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
rs!Fname = Text1.Text
rs!Lname = Text2.Text
rs!addr = Text3.Text

If Text4.Text <> "" Then
rs!Pin = Val(Text4.Text)
Else
rs!Pin = Null
End If

If Text5.Text <> "" Then
Dim dt As String
Dim splt() As String
splt = Split(Text5.Text, "/")
dt = splt(1) & "/" & splt(0) & "/" & splt(2)
rs!Dob = CDate(dt)
End If

If Text6.Text <> "" Then
rs!Phone = Val(Text6.Text)
Else
rs!Phone = Null
End If

rs!Email = Text7.Text

If Text8.Text <> "" Then
rs!Pmark = Val(Text8.Text)
Else
rs!Pmark = Null
End If

If Text9.Text <> "" Then
rs!Cmark = Val(Text9.Text)
Else
rs!Cmark = Null
End If

If Text10.Text <> "" Then
rs!Mmark = Val(Text10.Text)
Else
rs!Mmark = Null
End If

rs.Update

strSQL = "Delete from Optfor where Seatno = '" & Form2.sn2 & "'"
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
oconn.Execute strSQL
oconn.Close

For i = 0 To List1.ListCount - 1
TestArray() = Split(List1.List(i))
strSQL = "select * from Preferences where Preferences.colname = '" & TestArray(0) & "' and Preferences.strname = '" & TestArray(1) & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
CourseArray(i) = rs!Coursecode
Next i

strSQL = "select * from Optfor"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
For i = 0 To List1.ListCount
rs.AddNew
rs!seatno = Label4.Caption
rs!Coursecode = CourseArray(i)
rs!Prefno = i + 1
Next i


If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "Some fields are left blank"
Exit Sub
End If

m = MsgBox("Are you sure you want to submit the form?" & vbCrLf & "You can not edit the form after submission", vbOKCancel)
If m = vbCancel Then
Exit Sub
End If

strSQL = "select * from Student where Seatno = '" & Form2.sn2 & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
rs!Submit = 1
rs.Update
Form9.Show
Unload Me
End Sub

Private Sub Command7_Click()
strSQL = "select * from Student where Seatno = '" & Form2.sn2 & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
rs!Fname = Text1.Text
rs!Lname = Text2.Text
rs!addr = Text3.Text

If Text4.Text <> "" Then
rs!Pin = Val(Text4.Text)
Else
rs!Pin = Null
End If

If Text5.Text <> "" Then
Dim dt As String
Dim splt() As String
splt = Split(Text5.Text, "/")
dt = splt(1) & "/" & splt(0) & "/" & splt(2)
rs!Dob = CDate(dt)
End If

If Text6.Text <> "" Then
rs!Phone = Val(Text6.Text)
Else
rs!Phone = Null
End If

rs!Email = Text7.Text

If Text8.Text <> "" Then
rs!Pmark = Val(Text8.Text)
Else
rs!Pmark = Null
End If

If Text9.Text <> "" Then
rs!Cmark = Val(Text9.Text)
Else
rs!Cmark = Null
End If

If Text10.Text <> "" Then
rs!Mmark = Val(Text10.Text)
Else
rs!Mmark = Null
End If

rs.Update

strSQL = "Delete from Optfor where Seatno = '" & Form2.sn2 & "'"
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
oconn.Execute strSQL
oconn.Close

For i = 0 To List1.ListCount - 1
TestArray() = Split(List1.List(i))
strSQL = "select * from Preferences where Preferences.colname = '" & TestArray(0) & "' and Preferences.strname = '" & TestArray(1) & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
CourseArray(i) = rs!Coursecode
Next i

strSQL = "select * from Optfor"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
For i = 0 To List1.ListCount
rs.AddNew
rs!seatno = Label4.Caption
rs!Coursecode = CourseArray(i)
rs!Prefno = i + 1
Next i

m = MsgBox("Information is saved successfully" & vbCrLf & "You can edit it before submission", vbOKCancel)
If m = vbOK Then
Form7.Show
Unload Me
End If
End Sub

Private Sub Form_Click()
Form8.SetFocus
If Combo1.Text = "" Then
Combo1.Text = "Select College"
End If
If Combo2.Text = "" Then
Combo2.Text = "Select Stream"
End If
End Sub

Private Sub Form_Load()
Label4.Caption = Form2.sn2
strSQL = "select * from STUDENT where Seatno = '" & Form2.sn2 & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText

Text1.Text = rs!Fname

Text2.Text = rs!Lname

If rs!addr = "" Then
Text3.Text = ""
Else
If rs!addr <> "" Then
Text3.Text = rs!addr
End If
End If

If rs!Pin = "" Then
Text4.Text = ""
Else
If rs!Pin <> "" Then
Text4.Text = rs!Pin
End If
End If

If rs!Dob = "" Then
Text5.Text = ""
Else
If rs!Dob <> "" Then
Text5.Text = rs!Dob
Dim dt As String
Dim splt() As String
splt = Split(Text5.Text, "/")
dt = splt(1) & "/" & splt(0) & "/" & splt(2)
Text5.Text = dt
End If
End If

If rs!Phone = "" Then
Text6.Text = ""
Else
If rs!Phone <> "" Then
Text6.Text = rs!Phone
End If
End If

If rs!Email = "" Then
Text7.Text = ""
Else
If rs!Email <> "" Then
Text7.Text = rs!Email
End If
End If

If rs!Pmark = "" Then
Text8.Text = ""
Else
If rs!Pmark <> "" Then
Text8.Text = rs!Pmark
End If
End If

If rs!Cmark = "" Then
Text9.Text = ""
Else
If rs!Cmark <> "" Then
Text9.Text = rs!Cmark
End If
End If

If rs!Mmark = "" Then
Text10.Text = ""
Else
If rs!Mmark <> "" Then
Text10.Text = rs!Mmark
End If
End If

strSQL = "select * from College"
Set rs1 = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs1.CursorType = adOpenDynamic
rs1.CursorLocation = adUseClient
rs1.LockType = adLockOptimistic
rs1.Open strSQL, oconn, , , adCmdText
While rs1.EOF <> True
    Combo1.AddItem rs1!colname
    rs1.MoveNext
Wend
rs1.Close

strSQL = "select * from Optfor,Preferences where Optfor.Seatno = '" & Form2.sn2 & "' and Optfor.Coursecode = Preferences.CourseCode order by Prefno"
Set rs2 = New ADODB.Recordset
rs2.CursorType = adOpenDynamic
rs2.CursorLocation = adUseClient
rs2.LockType = adLockOptimistic
rs2.Open strSQL, oconn, , , adCmdText
While rs2.EOF <> True
    List1.AddItem (rs2!colname & " " & rs2!strname)
    rs2.MoveNext
Wend
rs2.Close
End Sub
Private Sub Combo1_After_Update()
If Combo1.Text = "" Then
Combo1.Text = "Select College"
Combo2.Clear
Combo2.Text = "Select Stream"
Else
Combo2.Clear
strSQL = "select * from Stream,College,Offers where Stream.Strid = Offers.Strid and Offers.Colid = College.Colid and College.Colname  = '" & Combo1.Text & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
While rs.EOF <> True
    Combo2.AddItem rs!strname
    rs.MoveNext
Wend
End If
End Sub
Sub ListBoxRemSel(lst As ListBox)
    Do Until lst.SelCount = 0
        If lst.Selected(a) Then lst.RemoveItem a: a = a - 1
        a = a + 1
    Loop
End Sub
'Private Sub Text1_KeyPress(KeyAscii As Integer)
    'ch = Chr$(KeyAscii)
   ' If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
  '      KeyAscii = 0
 '   End If
'End Sub

'Private Sub Text1_LostFocus()
'If Val(Text1.Text) > 99 Then
'MsgBox ("invalid seat number")
'Text1.Text = ""
'Text1.SetFocus
'End If
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "A" And ch <= "Z" Or KeyAscii = 8) Then
        KeyAscii = 0
        MsgBox "Block Letters Only"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "A" And ch <= "Z" Or KeyAscii = 8) Then
        KeyAscii = 0
        MsgBox "Block Letters Only"
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text4_LostFocus()
If Text4.Text = "" Then
Exit Sub
End If
If Val(Text4.Text) < 100000 Or Val(Text4.Text) > 999999 Then
MsgBox ("Invalid Pin Code")
Text4.Text = ""
Text4.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8 Or KeyAscii = 47) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_LostFocus()
If Text5.Text = "" Then
Exit Sub
End If
Dim splt() As String
splt = Split(Text5.Text, "/")
If Val(splt(1)) > 12 Or Val(splt(0)) > 31 Or Val(splt(2)) < 1987 Or Val(splt(2)) > 1991 Then
MsgBox ("Invalid Date")
Text5.SetFocus
Exit Sub
End If
If Not IsDate(Text5.Text) Then
MsgBox ("Invalid Date")
Text5.SetFocus
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text7_LostFocus()

If Text7.Text = "" Then
Exit Sub
End If

Dim Valid_Email As Boolean

Valid_Email = IsValidEmail(Text7.Text)

If Valid_Email = False Then
MsgBox "Enter A Valid Email Address"
Text7.SetFocus
End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text8_LostFocus()
If Text8.Text = "" Then
Exit Sub
End If
If Val(Text8.Text) < 0 Or Val(Text8.Text) > 50 Then
MsgBox ("Invalid Marks")
Text8.Text = ""
Text8.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text9_LostFocus()
If Text9.Text = "" Then
Exit Sub
End If
If Val(Text9.Text) < 0 Or Val(Text9.Text) > 50 Then
MsgBox ("Invalid Marks")
Text9.Text = ""
Text9.SetFocus
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    ch = Chr(KeyAscii)
    If Not (ch >= "0" And ch <= "9" Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text10_LostFocus()
If Text10.Text = "" Then
Exit Sub
End If
If Val(Text10.Text) < 0 Or Val(Text10.Text) > 100 Then
MsgBox ("Invalid Marks")
Text10.Text = ""
Text10.SetFocus
End If
End Sub

Public Function IsValidEmail(strEmail As String) As Boolean
Dim names, name, i, c
IsValidEmail = True

names = Split(strEmail, "@")

If UBound(names) <> 1 Then
IsValidEmail = False
Exit Function
End If

If names(1) = "yahoo.co.in" Then
Exit Function
End If

For Each name In names

If Len(name) <= 0 Then
IsValidEmail = False
Exit Function
End If

For i = 1 To Len(name)
c = LCase(Mid(name, i, 1))

If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
IsValidEmail = False
Exit Function
End If
Next

If Left(name, 1) = "." Or Right(name, 1) = "." Then
IsValidEmail = False
Exit Function
End If

Next

If InStr(names(1), ".") <= 0 Then
IsValidEmail = False
Exit Function
End If

i = Len(names(1)) - InStrRev(names(1), ".")

If i <> 2 And i <> 3 Then
IsValidEmail = False
Exit Function
End If

If InStr(strEmail, "..") > 0 Then
IsValidEmail = False
Exit Function
End If

End Function
