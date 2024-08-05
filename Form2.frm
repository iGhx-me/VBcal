VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "主程序"
   ClientHeight    =   6285
   ClientLeft      =   5280
   ClientTop       =   2145
   ClientWidth     =   11640
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6285
   ScaleWidth      =   11640
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame6 
      Caption         =   "说明"
      Height          =   3015
      Left            =   8160
      TabIndex        =   23
      Top             =   3240
      Width           =   3375
      Begin VB.Label Label9 
         Caption         =   "暂不支持传值传入原生计算器"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "注意,关闭本窗口,重置中转站"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "公告"
      Height          =   3135
      Left            =   8160
      TabIndex        =   20
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command13 
         Caption         =   "说明"
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "/4:第一次大更新"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "2024/8/3:软件诞生"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "工具选择"
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command12 
         Caption         =   "启动"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   735
      End
      Begin VB.Frame Frame7 
         Caption         =   "其他计算器"
         Height          =   2415
         Left            =   120
         TabIndex        =   26
         Top             =   3600
         Width           =   2055
         Begin VB.CommandButton Command11 
            Caption         =   "体重计算器"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "启动"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "启动"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "原生计算器:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "勾股定理计算器"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "方差/加权平均数计算器"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "传值(中转站)"
      Height          =   6135
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame4 
         Caption         =   "被传值"
         Height          =   5775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton Command9 
            Caption         =   "+"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   4680
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   15
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "/"
            Height          =   375
            Left            =   600
            TabIndex        =   14
            Top             =   4680
            Width           =   375
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H000000FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            MaskColor       =   &H000000FF&
            TabIndex        =   13
            Top             =   5280
            Width           =   375
         End
         Begin VB.CommandButton Command6 
            Caption         =   "-"
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   5280
            Width           =   375
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Height          =   4350
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "传值至"
         Height          =   5775
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            Caption         =   "[修改权]"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            Caption         =   "[修改值]"
            Height          =   255
            Left            =   1200
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            Caption         =   "[权]"
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "[值]"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "勾股定理计算器:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "方差/加权平均数计算器:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 在 Module1 模块中编写函数
Public Sub AddItemToList1(itemText As String)
    ' 确保 Form2 是打开的或者可以在需要时打开
    If Not Form2.Visible Then
        Form2.Show
    End If
    
    ' 添加一个判断，只向 List1 控件添加数字类型的值
    If IsNumeric(itemText) Then
        Form2.List1.AddItem itemText
    Else
        MsgBox "只能传数字类型的值", vbExclamation
    End If
End Sub

Private Sub Command1_Click()
Form1.Show vbModeless
End Sub

Private Sub Command10_Click()
MsgBox "注意:该程序有重大,Bug未来修复"
Form3.Show vbModeless
End Sub

Private Sub Command11_Click()
Form5.Show vbModeless
End Sub

Private Sub Command12_Click()
MsgBox "不好意思,还没做好"
MsgBox "我还是帮你调用系统计算器吧!"
Shell "calc.exe", vbNormalFocus
Form4.Show vbModeless
End Sub

Private Sub Command13_Click()
frmSplash.Show
frmAbout.Show
End Sub

Private Sub Command2_Click()
    Dim selectedIndex As String
    
    ' 获取当前选中项的索引
    selectedIndex = List1.ListIndex
    
    ' 检查是否有项被选中
    If selectedIndex >= 0 Then
        Form1.AddItemTovl List1.List(selectedIndex)
    Else
        MsgBox "请先选择"
    End If
End Sub

Private Sub Command3_Click()
    Dim selectedIndex As String
    
    ' 获取当前选中项的索引
    selectedIndex = List1.ListIndex
    
    ' 检查是否有项被选中
    If selectedIndex >= 0 Then
        Form1.AddItemTomy List1.List(selectedIndex)
    Else
        MsgBox "请先选择"
    End If
End Sub

Private Sub Command4_Click()
    Dim selectedIndex As String
    
    ' 获取当前选中项的索引
    selectedIndex = List1.ListIndex
    
    ' 检查是否有项被选中
    If selectedIndex >= 0 Then
        Form1.AddItemTot1 List1.List(selectedIndex)
    Else
        MsgBox "请先选择"
    End If
End Sub

Private Sub Command5_Click()
    Dim selectedIndex As String
    
    ' 获取当前选中项的索引
    selectedIndex = List1.ListIndex
    
    ' 检查是否有项被选中
    If selectedIndex >= 0 Then
        Form1.AddItemTot2 List1.List(selectedIndex)
    Else
        MsgBox "请先选择"
    End If
End Sub

Private Sub Command6_Click()
    ' 检查用户是否选择了某一项
    If List1.ListIndex <> -1 Then
        ' 删除用户选择的项
        List1.RemoveItem List1.ListIndex
    Else
        MsgBox "请选择一个项目来删除"
    End If
End Sub

Private Sub Command7_Click()
returef = InputBox("确认删除所有中转站的值吗?" & vbCrLf & " Yes(y) Or No(n) 输入y或n" & vbCrLf & "(不是y,包括空白即取消删除)", "二次确认")
If returef = "y" Then
    List1.Clear
End If
End Sub

Private Sub Command8_Click()
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

Private Sub Command9_Click()
inputValue1 = Text1.Text
If IsNumeric(inputValue1) Then
            List1.AddItem inputValue1
            Text1.Text = ""
        Else
            MsgBox "请输入数字"
        End If
End Sub

