VERSION 5.00
Begin VB.Form form1 
   Caption         =   "������"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "����"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox rst 
      Height          =   420
      Left            =   3240
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox bst 
      Height          =   420
      Left            =   3240
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Menu ��Ϸ 
      Caption         =   "��Ϸ"
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
      Begin VB.Menu ͳ����Ϣ 
         Caption         =   "ͳ����Ϣ"
      End
      Begin VB.Menu fgx 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(3, 3) As Integer    'ͨ��������������
Dim bs%, rs%    '�����������

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        a(1, 1) = a(1, 1) + 1
            If a(1, 1) Mod 2 = 0 Then
                Command1(0).Caption = "��"
                Command1(0).BackColor = RGB(3, 169, 244)
            Else
                Command1(0).Caption = "��"
                Command1(0).BackColor = RGB(244, 67, 54)
            End If
    Case 1
        a(1, 2) = a(1, 2) + 1
            If a(1, 2) Mod 2 = 0 Then
                Command1(1).Caption = "��"
                Command1(1).BackColor = RGB(3, 169, 244)
            Else
                Command1(1).Caption = "��"
                Command1(1).BackColor = RGB(244, 67, 54)
            End If
    Case 2
        a(1, 3) = a(1, 3) + 1
            If a(1, 3) Mod 2 = 0 Then
                Command1(2).Caption = "��"
                Command1(2).BackColor = RGB(3, 169, 244)
            Else
                Command1(2).Caption = "��"
                Command1(2).BackColor = RGB(244, 67, 54)
            End If
    Case 3
        a(2, 1) = a(2, 1) + 1
            If a(2, 1) Mod 2 = 0 Then
                Command1(3).Caption = "��"
                Command1(3).BackColor = RGB(3, 169, 244)
            Else
                Command1(3).Caption = "��"
                Command1(3).BackColor = RGB(244, 67, 54)
            End If
    Case 4
        a(2, 2) = a(2, 2) + 1
            If a(2, 2) Mod 2 = 0 Then
                Command1(4).Caption = "��"
                Command1(4).BackColor = RGB(3, 169, 244)
            Else
                Command1(4).Caption = "��"
                Command1(4).BackColor = RGB(244, 67, 54)
            End If
    Case 5
        a(2, 3) = a(2, 3) + 1
            If a(2, 3) Mod 2 = 0 Then
                Command1(5).Caption = "��"
                Command1(5).BackColor = RGB(3, 169, 244)
            Else
                Command1(5).Caption = "��"
                Command1(5).BackColor = RGB(244, 67, 54)
            End If
    Case 6
        a(3, 1) = a(3, 1) + 1
            If a(3, 1) Mod 2 = 0 Then
                Command1(6).Caption = "��"
                Command1(6).BackColor = RGB(3, 169, 244)
            Else
                Command1(6).Caption = "��"
                Command1(6).BackColor = RGB(244, 67, 54)
            End If
    Case 7
        a(3, 2) = a(3, 2) + 1
            If a(3, 2) Mod 2 = 0 Then
                Command1(7).Caption = "��"
                Command1(7).BackColor = RGB(3, 169, 244)
            Else
                Command1(7).Caption = "��"
                Command1(7).BackColor = RGB(244, 67, 54)
            End If
    Case 8
        a(3, 3) = a(3, 3) + 1
            If a(3, 3) Mod 2 = 0 Then
                Command1(8).Caption = "��"
                Command1(8).BackColor = RGB(3, 169, 244)
            Else
                Command1(8).Caption = "��"
                Command1(8).BackColor = RGB(244, 67, 54)
            End If
End Select
s = 0
Do
    'If s Mod 3 = 0 Then
        If a(1, 1) Mod 2 = 1 And a(1, 2) Mod 2 = 1 And a(1, 3) Mod 2 = 1 Then
            bs = bs + 1
            bst.Text = bs
        'ElseIf Command1(s).BackColor = Command1(s + 1).BackColor = Command1(s + 2).BackColor Then
            rs = rs + 1
            rst.Text = rs
        'Else
            '
        'End If
    End If
        s = s + 1
        If s = 6 Then
        Exit Do
        End If
Loop
'Print Str(a(1, 1))
End Sub

Private Sub Command2_Click()    '���(����һ��)
    For x = 1 To 3
        For y = 1 To 3
            a(x, y) = -1
        Next y
    Next x
Dim i As Integer
    For i = 0 To 8
        Command1(i).Caption = ""
        Command1(i).FontSize = 15
        Command1(i).BackColor = &H8000000F  '����ɫΪ����ť���桱
    Next i
End Sub

Private Sub Form_Load()
    For x = 1 To 3  '��ʼ����ť����
        For y = 1 To 3
            a(x, y) = -1
        Next y
    Next x
    For i = 0 To 8  '��ʼ������
        Command1(i).Caption = ""
        Command1(i).FontSize = 15
        'Command1(i).Style = 1
        'Command1(i).BackColor = RGB(3, 169, 244)��ɫ
        'Command1(i).BackColor = RGB(244��67��54)��ɫ
    Next i
bs = 0: rs = 0  '��ʼ������
End Sub
