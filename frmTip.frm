VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "日积月累"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5685
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "在启动时显示提示(&S)"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2940
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "您知道吗..."
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmTip.frx":030A
         Height          =   1635
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DoNextTip()

    
End Sub


Private Sub chkLoadTipsAtStartup_Click()
    ' 保存在下次启动时是否显示此窗体
    SaveSetting App.EXEName, "Options", "在启动时显示提示", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
