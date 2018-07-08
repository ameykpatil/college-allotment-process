VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New User / Old Account  "
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Login"
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
      MouseIcon       =   "Form2.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Register"
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
      MouseIcon       =   "Form2.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Already have an account?"
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
      TabIndex        =   3
      Top             =   6000
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New user? "
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
      TabIndex        =   1
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label Label1 
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
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form2.frx":03F6
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sn2 As Integer

Private Sub Command1_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form5.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form1.Show
Unload Me
End Sub
