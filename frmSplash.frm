VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "版权所有:github上iGhx-me"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "网站:www.ighx.me"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "警告:免费软件,禁止盈利"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6765
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6240
         TabIndex        =   5
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5805
         TabIndex        =   6
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "万能计算器"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   32.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   3300
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "授权"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "https://www.iGhx.me/"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   7
         Top             =   705
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

