VERSION 5.00
Begin VB.Form frmProgressBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple ProgressBar Demo"
   ClientHeight    =   5172
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   5736
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   5736
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBarDemo.ProgressBar prbVertical 
      Height          =   4932
      Left            =   5160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   8700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8454143
      Orientation     =   1
   End
   Begin ProgressBarDemo.ProgressBar prbHorizontal 
      Height          =   372
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   4932
      _ExtentX        =   8700
      _ExtentY        =   656
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8454143
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Stop"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Tag             =   "&Start"
      Top             =   4680
      Width           =   1452
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00C000C0&
      Height          =   612
      Index           =   6
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Magenta"
      Top             =   2280
      Width           =   612
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00C00000&
      Height          =   612
      Index           =   5
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Blue"
      Top             =   2280
      Width           =   612
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00C0C000&
      Height          =   612
      Index           =   4
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cyan"
      Top             =   2280
      Width           =   612
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H0000C000&
      Height          =   612
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Green"
      Top             =   2280
      Width           =   612
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H0000C0C0&
      Height          =   612
      Index           =   2
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Yellow"
      Top             =   2280
      Width           =   612
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H000000C0&
      Height          =   612
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Red"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00BC8854&
      Height          =   612
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Default"
      Top             =   2280
      Width           =   612
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   3600
      TabIndex        =   0
      Top             =   4680
      Width           =   1452
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "&Disable"
      Height          =   372
      Left            =   1680
      TabIndex        =   9
      Tag             =   "&Enable"
      Top             =   4680
      Width           =   1452
   End
   Begin VB.Timer tmrProgress 
      Interval        =   150
      Left            =   4680
      Top             =   720
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdColor_Click(Index As Integer)

   With prbHorizontal
      Select Case Index
         Case 0
            .BarColor = Default
         Case 1
            .BarColor = Red
            
         Case 2
            .BarColor = Yellow
            
         Case 3
            .BarColor = Green
            
         Case 4
            .BarColor = Cyan
            
         Case 5
            .BarColor = Blue
            
         Case 6
            .BarColor = Magenta
      End Select
   End With
   
   With prbVertical
      Select Case Index
         Case 0
            .BarColor = Default
         Case 1
            .BarColor = Red
            
         Case 2
            .BarColor = Yellow
            
         Case 3
            .BarColor = Green
            
         Case 4
            .BarColor = Cyan
            
         Case 5
            .BarColor = Blue
            
         Case 6
            .BarColor = Magenta
      End Select
   End With

End Sub

Private Sub cmdEnable_Click()

Dim strTemp As String

   With cmdEnable
      strTemp = .Tag
      .Tag = .Caption
      .Caption = strTemp
   End With
   
   prbHorizontal.Enabled = Not prbHorizontal.Enabled
   prbVertical.Enabled = Not prbVertical.Enabled

End Sub

Private Sub cmdExit_Click()

   Unload Me

End Sub

Private Sub cmdStartStop_Click()

Dim strTemp As String

   With cmdStartStop
      strTemp = .Tag
      .Tag = .Caption
      .Caption = strTemp
   End With
   
   If tmrProgress.Enabled Then
      prbHorizontal.Clear
      prbVertical.Clear
   End If
   
   tmrProgress.Enabled = Not tmrProgress.Enabled

End Sub

Private Sub Form_Load()

Dim intWidth As Integer

   prbHorizontal.Height = prbHorizontal.Height / 2
   prbVertical.Width = prbVertical.Width / 2
   DoEvents
   intWidth = Width - ScaleWidth
   Width = prbVertical.Left + prbVertical.Width + prbHorizontal.Left + intWidth

End Sub

Private Sub tmrProgress_Timer()

   With prbHorizontal
      If .Value = 100 Then
         Call .Clear
         
      Else
         .Value = .Value + 1
         .Caption = .Value & "%"
      End If
   End With
   
   With prbVertical
      If .Value = 100 Then
         Call .Clear
         
      Else
         .Value = .Value + 1
         .Caption = .Value & "%"
      End If
   End With

End Sub
