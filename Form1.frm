VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��Ȩƽ����&���������"
   ClientHeight    =   5985
   ClientLeft      =   8700
   ClientTop       =   3525
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   4635
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton SumMain 
      Caption         =   "����������"
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton ChatO 
      Caption         =   "��ֵ"
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton ValuetO 
      Caption         =   "��ֵ"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "[�޸�Ȩ]"
      Height          =   495
      Left            =   2400
      TabIndex        =   22
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3480
      TabIndex        =   21
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1200
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "[�޸�ֵ]"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "�洢ֵ"
      Height          =   2895
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2055
      Begin VB.ListBox List1 
         Height          =   2580
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton reF 
      Caption         =   "���"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "�洢Ȩ"
      Height          =   2895
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
      Begin VB.ListBox List2 
         Height          =   2580
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton delv 
      Caption         =   "ɾ��βֵ"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   855
      Begin VB.TextBox fcnum 
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ƽ����"
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   855
      Begin VB.TextBox pjnum 
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox my 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton run 
      Caption         =   "����"
      Height          =   1215
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton addNum 
      Caption         =   "���"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox vl 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "����:2010tcsy/www.ighx.me"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "[Ȩ]"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "[ֵ]"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "δ����ֵ,�򱨴�"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "δ����Ȩ,����1��"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddItemTot1(itemText As Double)
    If Not Form1.Visible Then
        Form1.Show
    End If
        Form1.Text1.Text = itemText
End Sub
Public Sub AddItemTot2(itemText As Double)
    If Not Form1.Visible Then
        Form1.Show
    End If
        Form1.Text2.Text = itemText
End Sub
Public Sub AddItemTovl(itemText As Double)
    If Not Form1.Visible Then
        Form1.Show
    End If
        Form1.vl.Text = itemText
End Sub
Public Sub AddItemTomy(itemText As Double)
    If Not Form1.Visible Then
        Form1.Show
    End If
        Form1.my.Text = itemText
End Sub
Private Sub addNum_Click()
    inputValue = vl.Text
    inputValue1 = my.Text
    If IsNumeric(inputValue) Then
        If IsNumeric(inputValue1) Then
            List1.AddItem inputValue
            List2.AddItem inputValue1
            vl = ""
            my = ""
        ElseIf inputValue1 = "" Then
            inputValue1 = 1
            List1.AddItem inputValue
            List2.AddItem inputValue1
            vl = ""
            my = ""
        Else
        MsgBox "����������"
        vl = ""
        my = ""
        End If
    Else
        MsgBox "������[ֵ] �� ȷ�������[ֵ]Ϊ����"
        vl = ""
        my = ""
    End If

End Sub

Private Sub ChatO_Click()
    Form2.AddItemToList1 fcnum.Text
    fcnum.Text = ""
End Sub

Private Sub Command1_Click()
    Dim selectedIndex As String
    
    ' ��ȡ��ǰѡ���������
    selectedIndex = List1.ListIndex
    
    ' ����Ƿ����ѡ��
    If selectedIndex >= 0 Then
        inputV = Text1.Text
        If IsNumeric(inputV) Then
            ' �޸�ѡ�����ֵ
            List1.List(selectedIndex) = inputV
            Text1.Text = ""
        ElseIf inputV = "" Then
            MsgBox "������ѡ��ֵ"
        Else
        MsgBox "����������"
        End If
    Else
        MsgBox "����ѡ��"
    End If
End Sub

Private Sub Command2_Click()
    Dim selectedIndex As String
    
    ' ��ȡ��ǰѡ���������
    selectedIndex = List2.ListIndex
    
    ' ����Ƿ����ѡ��
    If selectedIndex >= 0 Then
        inputC = Text2.Text
        If IsNumeric(inputC) Then
            ' �޸�ѡ�����ֵ
            List2.List(selectedIndex) = inputC
            Text2.Text = ""
        ElseIf inputC = "" Then
            MsgBox "������ѡ��ֵ"
        Else
        MsgBox "����������"
        End If
    Else
        MsgBox "����ѡ��"
    End If
End Sub

Private Sub delv_Click()
    Dim lastIndex As Integer
    ' ��ȡ���һ�������
    lastIndex = List1.ListCount - 1
    ' ��� ListBox ������
    If lastIndex >= 0 Then
        ' ɾ�����һ��
        List1.RemoveItem lastIndex
        List2.RemoveItem lastIndex
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim inputC As String
    Dim inputV As String
    ' ��ȡ�ı����ֵ
    Dim inputValue As String
    ' ��ȡ�ı����ֵ
    Dim inputValue1 As String
    Dim textLength As Integer
End Sub

Private Sub Label5_Click()
frmTip.Show vbModeless
End Sub

Private Sub reF_Click()
returef = InputBox("ȷ��ɾ������ �洢ֵ �� �洢Ȩ��?" & vbCrLf & " Yes(y) Or No(n) ����y��n" & vbCrLf & "(����y,�����հ׼�ȡ��ɾ��)", "����ȷ��")
If returef = "y" Then
    List1.Clear
    List2.Clear
End If
End Sub
Private Sub run_Click()
    Dim total As Double
    Dim weightsTotal As Integer
    Dim weightedAverage As Double
    Dim variance As Double ' �������

    ' ��ʼ���ܺ͡�Ȩ���ܺͺͷ���
    total = 0
    weightsTotal = 0
    variance = 0
    
    ' ȷ�� List1 �� List2 ����Ŀ����ͬ
    If List1.ListCount <> List2.ListCount Then
        MsgBox "����List1 �� List2 ����Ŀ����һ�¡�"
        Exit Sub
    End If
    
    ' ѭ������ ListBox �е�ÿ����Ŀ
    For i = 0 To List1.ListCount - 1
        ' ��ȡ��ǰ��Ŀ��ֵ��������ת��Ϊ Double ����
        Dim itemValue As Double
        itemValue = CDbl(List1.List(i))
        
        ' ��ȡ��Ӧλ���ϵ�Ȩ��ֵ������Ȩ�ش洢�� List2 �У�
        Dim weight As Integer
        weight = CInt(List2.List(i))
        
        ' �����Ȩ�ܺ�
        total = total + itemValue * weight
        weightsTotal = weightsTotal + weight
        
        ' ���㷽��
        variance = variance + weight * (itemValue - weightedAverage) ^ 2
    Next i
    
    ' �����Ȩƽ����
    If weightsTotal > 0 Then
        weightedAverage = total / weightsTotal
    Else
        weightedAverage = 0 ' ������������
    End If
    
    ' ���Ȩ���ܺʹ���0����������յķ���ֵ
    If weightsTotal > 0 Then
        variance = variance / weightsTotal
    Else
        variance = 0 ' ������������
    End If
    
    ' ��ʾ��Ȩƽ�����ͷ�����
    'MsgBox "��Ȩƽ����Ϊ: " & weightedAverage & vbCrLf & "����Ϊ: " & variance
    pjnum.Text = weightedAverage
    fcnum.Text = variance
End Sub

Private Sub SumMain_Click()
    Form2.Show vbModeless  ' ���ز���ʾ Form1 ����
End Sub

Private Sub ValuetO_Click()
    Form2.AddItemToList1 pjnum.Text
    pjnum.Text = ""
End Sub

