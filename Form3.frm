VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "���ɶ��������"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form3"
   ScaleHeight     =   4095
   ScaleWidth      =   4095
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "����������"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ֵ"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ֵ"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ֵ"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "���ֻ��һ��ֵ,�򲻳����"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "���ɽ�������"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ֻҪ������������ֵ"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   1200
      Y1              =   1440
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   2400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   1440
      Y2              =   2880
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private a As Double, b As Double, c As Double

Public Sub run()
    If IsNumeric(Text1.Text) Then
        a = CDbl(Text1.Text)
    Else
        a = 0 ' �����������һЩ������
    End If
    
    If IsNumeric(Text2.Text) Then
        b = CDbl(Text2.Text)
    Else
        b = 0 ' �����������һЩ������
    End If
    
    If a > 0 And b > 0 Then
        ' ��֪ a �� b���� c
        c = Sqr(a ^ 2 + b ^ 2)
        Text3.Text = c
    ElseIf a > 0 And c > 0 Then
        ' ��֪ a �� c���� b
        b = Sqr(c ^ 2 - a ^ 2)
        Text2.Text = b
    ElseIf b > 0 And c > 0 Then
        ' ��֪ b �� c���� a
        a = Sqr(c ^ 2 - b ^ 2)
        Text1.Text = a
    Else
        ' �����޷�����������������ʾ������Ϣ������ս����
        Text3.Text = ""
    End If
End Sub

Private Sub Command4_Click()
    Form2.Show vbModeless
End Sub

Private Sub Form_Load()
    ' ��ʼ������
    'Dim a As Double, b As Double, c As Double
End Sub

Private Sub Text1_Change()
    run ' ÿ�� Text1 ���ı��ı�ʱ���¼���
End Sub

Private Sub Text2_Change()
    run ' ÿ�� Text2 ���ı��ı�ʱ���¼���
End Sub

Private Sub Text3_Change()
    run ' Text3 �ǽ������ʾ�򣬲���Ҫ���¼���
End Sub

