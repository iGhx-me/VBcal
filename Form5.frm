VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "���ؼ�����"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   1590
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "��Ϣ"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4335
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "��"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ǧ��(����)"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function AddNumbers(ByVal num1 As Double) As Double
    If num1 <= 100 Then
        Label1.Caption = "��ϲ��,���ز���80KG,���ȫ��80%��"
    Else
        Label1.Caption = "��ϲ��,���ز�100KG,���ȫ��60%��"
    End If
End Function

Private Sub Command1_Click()

wi = Text1.Text

option1Selected = Option1.Value
option2Selected = Option2.Value

If option1Selected Then
    AddNumbers (wi)
ElseIf option2Selected Then
    AddNumbers (wi / 2)
Else
    MsgBox "��ѡ��λ"
End If

End Sub

Private Sub Form_Load()
Dim wi As Double
Dim option1Selected As Boolean
Dim option2Selected As Boolean
End Sub
