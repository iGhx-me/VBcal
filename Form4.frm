VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "计算器"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   5430
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   4095
      End
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      TabIndex        =   29
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command29 
      Caption         =   "("
      Height          =   615
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command28 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   27
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command27 
      Caption         =   ")"
      Height          =   615
      Left            =   1440
      TabIndex        =   26
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      Caption         =   "/"
      Height          =   615
      Left            =   2280
      TabIndex        =   25
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   24
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "根号"
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "0"
      Height          =   615
      Left            =   840
      TabIndex        =   22
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "%"
      Height          =   615
      Left            =   1440
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "*"
      Height          =   615
      Left            =   2280
      TabIndex        =   20
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "-"
      Height          =   615
      Left            =   2280
      TabIndex        =   19
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "+"
      Height          =   615
      Left            =   2280
      TabIndex        =   18
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "小数点"
      Height          =   615
      Left            =   2280
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "MR"
      Height          =   615
      Left            =   3120
      TabIndex        =   16
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "MC"
      Height          =   615
      Left            =   3120
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "MS"
      Height          =   615
      Left            =   3120
      TabIndex        =   14
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "M-"
      Height          =   615
      Left            =   3840
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "M+"
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "传值"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<--"
      Height          =   615
      Left            =   3840
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   9
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   615
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
