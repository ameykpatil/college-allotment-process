VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Information Module"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   2
      RowHeight       =   33
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      ItemData        =   "Form10.frx":0000
      Left            =   6720
      List            =   "Form10.frx":0002
      TabIndex        =   2
      Text            =   "Select College"
      Top             =   6240
      Width           =   2295
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
      MouseIcon       =   "Form10.frx":0004
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2655
      Left            =   3960
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   2
      RowHeight       =   33
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   3375
      Left            =   11880
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   3375
      Left            =   360
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form10.frx":0156
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
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim strSQL As String

Private Sub Combo1_Click()
If Combo1.Text = "" Then
Combo1.Text = "Select College"
Else
DataGrid2.Visible = True
Image2.Visible = True
Image3.Visible = True
strSQL = "select offers.strid, strname, intake, vacant, cutoff from Stream,College,Offers where Stream.Strid = Offers.Strid and Offers.Colid = College.Colid and College.Colname  = '" & Combo1.Text & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
Set DataGrid2.DataSource = rs
DataGrid2.Columns("strid").Width = 800
DataGrid2.Columns("strname").Width = 2000
DataGrid2.Columns("intake").Width = 1500
DataGrid2.Columns("vacant").Width = 1550
DataGrid2.Columns("cutoff").Width = 1550

strSQL = "select image1,image2 from college where College.Colname  = '" & Combo1.Text & "'"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenDynamic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
Image2 = LoadPicture(rs!Image1)
Image3 = LoadPicture(rs!Image2)
End If
End Sub

Private Sub Command3_Click()
strSQL = "select * from result where name = 'admin' "
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs1.CursorType = adOpenDynamic
rs1.CursorLocation = adUseClient
rs1.LockType = adLockOptimistic
rs1.Open strSQL, oconn, , , adCmdText
If rs1!flag = 1 Then
rs1.Close
Form11.Show
Unload Me
Exit Sub
End If
rs1.Close
Form7.Show
Unload Me
End Sub

Private Sub Form_Click()
Command3.SetFocus
End Sub

Private Sub Form_Load()
strSQL = "select colid,colname,addr,fees,grade,type,mintype,quota from COLLEGE"
Set rs = New ADODB.Recordset
Set oconn = New ADODB.Connection
oconn.Open "Provider=OraOLEDB.Oracle;Data Source=Soham_Laptop;User Id=sneha;Password=jaas;"
rs.CursorType = adOpenStatic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
Set DataGrid1.DataSource = rs
DataGrid1.Columns("colid").Width = 800
DataGrid1.Columns("colname").Width = 1500
DataGrid1.Columns("addr").Width = 1500
DataGrid1.Columns("fees").Width = 1000
DataGrid1.Columns("grade").Width = 1000
DataGrid1.Columns("type").Width = 2000
DataGrid1.Columns("mintype").Width = 2000
DataGrid1.Columns("quota").Width = 1050
While rs.EOF <> True
    Combo1.AddItem rs!colname
    rs.MoveNext
Wend
End Sub


