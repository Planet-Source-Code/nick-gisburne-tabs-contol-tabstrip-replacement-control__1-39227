VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabs Demonstration"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "TabsDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ctlTabs ctlTabs1 
      Height          =   2385
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   4207
      TABWIDE         =   70
      TABHIGH         =   20
      TABCOUNT        =   2
      TABSELECTED     =   1
      TABSTYLE        =   2
      CAPTIONSTYLE    =   4
      FOCUSRECT       =   0   'False
      TABCAPTION1     =   "First"
      TABCAPTION2     =   "Second"
      BeginProperty TABFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TABFONTACTIVE {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TABCOLOR        =   -2147483632
      TABCOLORACTIVE  =   -2147483633
      TEXTCOLOR       =   -2147483628
      TEXTCOLORACTIVE =   -2147483630
      Begin VB.Frame Framer 
         Caption         =   "Second Tab Contents"
         Height          =   1875
         Index           =   2
         Left            =   105
         TabIndex        =   3
         Top             =   75
         Visible         =   0   'False
         Width           =   4275
         Begin Project1.ctlTabs ctlTabs2 
            Height          =   825
            Left            =   165
            TabIndex        =   6
            Top             =   855
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   1455
            TABWIDE         =   20
            TABHIGH         =   15
            TABCOUNT        =   3
            TABSELECTED     =   1
            TABSTYLE        =   5
            CAPTIONSTYLE    =   4
            FOCUSRECT       =   0   'False
            TABCAPTION1     =   "1"
            TABCAPTION2     =   "2"
            TABCAPTION3     =   "3"
            BeginProperty TABFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TABFONTACTIVE {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TABCOLOR        =   -2147483632
            TABCOLORACTIVE  =   -2147483633
            TEXTCOLOR       =   -2147483628
            TEXTCOLORACTIVE =   -2147483630
            Begin VB.Label Label2 
               Caption         =   "You can even store one tab control inside another!"
               Height          =   450
               Left            =   210
               TabIndex        =   7
               Top             =   165
               Width           =   2865
            End
         End
         Begin VB.Label Label1 
            Caption         =   "This is the second tab - contents of the first one are now hidden."
            Height          =   1275
            Index           =   1
            Left            =   195
            TabIndex        =   4
            Top             =   330
            Width           =   3960
         End
      End
      Begin VB.Frame Framer 
         Caption         =   "First Tab Contents"
         Height          =   1875
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   75
         Width           =   4275
         Begin VB.CommandButton cmd 
            Caption         =   "Change Style"
            Height          =   360
            Left            =   180
            TabIndex        =   5
            Top             =   1305
            Width           =   1230
         End
         Begin VB.Label Label1 
            Caption         =   $"TabsDemo.frx":000C
            Height          =   1275
            Index           =   0
            Left            =   195
            TabIndex        =   2
            Top             =   330
            Width           =   3960
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    ctlTabs1.Style = IIf(ctlTabs1.Style = [Bottom Right], [Bottom Left], [Bottom Right])
End Sub

Private Sub ctlTabs1_TabClick(OldTab As Integer, NewTab As Integer)
    If OldTab <> NewTab Then
        Framer(OldTab).Visible = False
        Framer(NewTab).Visible = True
    End If
End Sub

Private Sub ctlTabs2_TabClick(OldTab As Integer, NewTab As Integer)
    Me.Caption = "Tab Changed from " & OldTab & " to " & NewTab
End Sub
