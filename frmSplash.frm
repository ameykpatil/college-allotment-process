VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form0 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5160
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7800
      Top             =   240
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSplash.frx":000C
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   4
      X1              =   8400
      X2              =   8400
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   4
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H008080FF&
      BorderWidth     =   4
      X1              =   120
      X2              =   8400
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      BorderWidth     =   4
      X1              =   120
      X2              =   8400
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Top             =   2400
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Top             =   2160
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "College Allotment Process"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   4440
      Picture         =   "frmSplash.frx":00B4
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Value > 1 And ProgressBar1.Value < 20 Then
Label2.Caption = "Loading Program...."
End If

If ProgressBar1.Value > 20 And ProgressBar1.Value < 35 Then
Label2.Caption = "Loading Settings...."
End If

If ProgressBar1.Value > 35 And ProgressBar1.Value < 45 Then
Label2.Caption = "Searching For Components ...."
End If

If ProgressBar1.Value = 70 Then
Label2.Caption = "Reading Database...."
Timer1.Interval = 150
End If
            
If ProgressBar1.Value = 80 Then
Timer1.Interval = 100
Label2.Caption = "Start Program...."
End If
            
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Interval = 0
Form1.Show
Unload Me
End If

End Sub
