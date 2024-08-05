VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "加权平均数&方差计算器"
   ClientHeight    =   5985
   ClientLeft      =   8700
   ClientTop       =   3525
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   4635
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton SumMain 
      Caption         =   "呼叫主程序"
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton ChatO 
      Caption         =   "传值"
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton ValuetO 
      Caption         =   "传值"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "[修改权]"
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
      Caption         =   "[修改值]"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "存储值"
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
      Caption         =   "清空"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "存储权"
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
      Caption         =   "删除尾值"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "方差"
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
      Caption         =   "平均数"
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
      Caption         =   "计算"
      Height          =   1215
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton addNum 
      Caption         =   "添加"
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
      Caption         =   "作者:2010tcsy/www.ighx.me"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "[权]"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "[值]"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "未输入值,则报错"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "未输入权,则作1算"
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
        MsgBox "请输入数字"
        vl = ""
        my = ""
        End If
    Else
        MsgBox "请输入[值] 或 确认输入的[值]为数字"
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
    
    ' 获取当前选中项的索引
    selectedIndex = List1.ListIndex
    
    ' 检查是否有项被选中
    If selectedIndex >= 0 Then
        inputV = Text1.Text
        If IsNumeric(inputV) Then
            ' 修改选中项的值
            List1.List(selectedIndex) = inputV
            Text1.Text = ""
        ElseIf inputV = "" Then
            MsgBox "请输入选项值"
        Else
        MsgBox "请输入数字"
        End If
    Else
        MsgBox "请先选择"
    End If
End Sub

Private Sub Command2_Click()
    Dim selectedIndex As String
    
    ' 获取当前选中项的索引
    selectedIndex = List2.ListIndex
    
    ' 检查是否有项被选中
    If selectedIndex >= 0 Then
        inputC = Text2.Text
        If IsNumeric(inputC) Then
            ' 修改选中项的值
            List2.List(selectedIndex) = inputC
            Text2.Text = ""
        ElseIf inputC = "" Then
            MsgBox "请输入选项值"
        Else
        MsgBox "请输入数字"
        End If
    Else
        MsgBox "请先选择"
    End If
End Sub

Private Sub delv_Click()
    Dim lastIndex As Integer
    ' 获取最后一项的索引
    lastIndex = List1.ListCount - 1
    ' 如果 ListBox 中有项
    If lastIndex >= 0 Then
        ' 删除最后一项
        List1.RemoveItem lastIndex
        List2.RemoveItem lastIndex
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim inputC As String
    Dim inputV As String
    ' 获取文本框的值
    Dim inputValue As String
    ' 获取文本框的值
    Dim inputValue1 As String
    Dim textLength As Integer
End Sub

Private Sub Label5_Click()
frmTip.Show vbModeless
End Sub

Private Sub reF_Click()
returef = InputBox("确认删除所有 存储值 和 存储权吗?" & vbCrLf & " Yes(y) Or No(n) 输入y或n" & vbCrLf & "(不是y,包括空白即取消删除)", "二次确认")
If returef = "y" Then
    List1.Clear
    List2.Clear
End If
End Sub
Private Sub run_Click()
    Dim total As Double
    Dim weightsTotal As Integer
    Dim weightedAverage As Double
    Dim variance As Double ' 方差变量

    ' 初始化总和、权重总和和方差
    total = 0
    weightsTotal = 0
    variance = 0
    
    ' 确保 List1 和 List2 的项目数相同
    If List1.ListCount <> List2.ListCount Then
        MsgBox "错误：List1 和 List2 的项目数不一致。"
        Exit Sub
    End If
    
    ' 循环遍历 ListBox 中的每个项目
    For i = 0 To List1.ListCount - 1
        ' 获取当前项目的值，并将其转换为 Double 类型
        Dim itemValue As Double
        itemValue = CDbl(List1.List(i))
        
        ' 获取对应位置上的权重值（假设权重存储在 List2 中）
        Dim weight As Integer
        weight = CInt(List2.List(i))
        
        ' 计算加权总和
        total = total + itemValue * weight
        weightsTotal = weightsTotal + weight
        
        ' 计算方差
        variance = variance + weight * (itemValue - weightedAverage) ^ 2
    Next i
    
    ' 计算加权平均数
    If weightsTotal > 0 Then
        weightedAverage = total / weightsTotal
    Else
        weightedAverage = 0 ' 避免除以零错误
    End If
    
    ' 如果权重总和大于0，则计算最终的方差值
    If weightsTotal > 0 Then
        variance = variance / weightsTotal
    Else
        variance = 0 ' 避免除以零错误
    End If
    
    ' 显示加权平均数和方差结果
    'MsgBox "加权平均数为: " & weightedAverage & vbCrLf & "方差为: " & variance
    pjnum.Text = weightedAverage
    fcnum.Text = variance
End Sub

Private Sub SumMain_Click()
    Form2.Show vbModeless  ' 加载并显示 Form1 窗口
End Sub

Private Sub ValuetO_Click()
    Form2.AddItemToList1 pjnum.Text
    pjnum.Text = ""
End Sub

