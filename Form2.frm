VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form oyesyes��½�� 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "oyesyes��½��"
   ClientHeight    =   8250
   ClientLeft      =   2265
   ClientTop       =   1500
   ClientWidth     =   12000
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":1232A
   ScaleHeight     =   8250
   ScaleWidth      =   12000
   Begin SHDocVwCtl.WebBrowser ��ҳ 
      Height          =   4575
      Left            =   600
      TabIndex        =   0
      Top             =   2880
      Width           =   7815
      ExtentX         =   13785
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label ��ҳ 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label ��ǰ�汾 
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�汾"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label ��θ��� 
      BackStyle       =   0  'Transparent
      Caption         =   "(��θ���?)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Image �뿪 
      Height          =   1095
      Left            =   9120
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Image ��Ϸ���� 
      Height          =   1215
      Left            =   9120
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Image ��ʼ��Ϸ 
      Height          =   1605
      Left            =   9120
      Top             =   2640
      Width           =   2970
   End
   Begin VB.Image ������ 
      Height          =   8250
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "oyesyes��½��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'�����ڴ�����û�����еĳ���
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'����API����
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HTCAPTION = 2     '��������
Const WM_NCLBUTTONDOWN = &HA1

Private Const WS_EX_LAYERED = &H80000
'Const GWL_EXSTYLE = (0)
Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const WEB = "http://www.baidu.com"


Private Function fun_FindProcess(ByVal ProcessName As String) As Long

    Dim strdata As String
    Dim my As PROCESSENTRY32
    Dim l As Long
    Dim l1 As Long
    Dim mName As String
    Dim i As Integer, pid As Long
    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
    my.dwSize = 1060
    If (Process32First(l, my)) Then
    Do
    i = InStr(1, my.szExeFile, Chr(0))
    mName = LCase(Left(my.szExeFile, i - 1))
    If mName = LCase(ProcessName) Then
    pid = my.th32ProcessID
    fun_FindProcess = pid
    Exit Function
    End If
    Loop Until (Process32Next(l, my) < 1)
    End If
    l1 = CloseHandle(l)
    End If
    fun_FindProcess = 0
    
End Function

Private Function GetTxt(TxtPath As String) As String

    Dim i As Integer: i = FreeFile
    Open TxtPath For Input As #i
        GetTxt = StrConv(InputB(LOF(i), i), vbUnicode)
    Close #i

End Function

Sub Form_Load()
    
    oyesyes��½��.Width = 12000
    oyesyes��½��.Height = 8250
    
    oyesyes��½��.Left = (Screen.Width - Me.Width) / 2
    oyesyes��½��.Top = (Screen.Height - Me.Height) / 2
         
        
    ������.Width = 12000
    ������.Height = 8250
    ������.Left = 0
    ������.Top = 0
    
    ������.Picture = LoadPicture(App.Path & "\pic\������.bmp")
    
    ��ʼ��Ϸ.Width = 2970
    ��ʼ��Ϸ.Height = 1605
    ��ʼ��Ϸ.Top = 2700
    ��ʼ��Ϸ.Left = 8685
    ��ʼ��Ϸ.Picture = LoadPicture(App.Path & "\pic\start.jpg")
    
    ��Ϸ����.Width = 2970
    ��Ϸ����.Height = 1125
    ��Ϸ����.Top = 4650
    ��Ϸ����.Left = 8685
    ��Ϸ����.Picture = LoadPicture(App.Path & "\pic\setting.jpg")
    
    �뿪.Width = 2970
    �뿪.Height = 1065
    �뿪.Top = 6165
    �뿪.Left = 8685
    �뿪.Picture = LoadPicture(App.Path & "\pic\exit.jpg")

    ��ҳ.Width = 7930
    ��ҳ.Height = 4650
    ��ҳ.Top = 2840
    ��ҳ.Left = 500
    ��ҳ.Navigate "http://oyesyes.com/ro/Ragnarok.html"
    
    Dim rtn As Long
    BorderStyler = 0
    
    ��θ���.BackStyle = 0
    ��θ���.Top = 7630
    ��ǰ�汾.BackStyle = 0
    ��ǰ�汾.Top = 7630
    
    ��ҳ.BackStyle = 0
    ��ҳ.Top = 7630
    ��ҳ.Caption = "www.oyesyes.com"
    
    
    If Dir(App.Path & "\oyesyes.config") = "" Then
        '
    Else
        ��ǰ�汾.Caption = "��ǰ�汾: " & Split(GetTxt(App.Path & "\oyesyes.config"), vbCrLf)(1)
    End If

    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HFF00FF, 0, LWA_COLORKEY
    
End Sub


Private Sub ������_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
                                        '���������¼�
    Dim i, Xx As Long
    If Button = 1 Then                  '�������������
        i = ReleaseCapture()            '�ͷ���겶��
        Xx = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0) '����Ϣ��������
    End If
End Sub

Private Sub ��ʼ��Ϸ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ret As Long
    
    If Button = 1 Then
        ��ʼ��Ϸ.Picture = LoadPicture(App.Path & "\pic\start_click.jpg")
        If Dir(App.Path & "\oyesyes.exe") = "" Then
            ret = MsgBox("�ڵ�ǰĿ¼�£��Ҳ���oyesyes.exe����ѱ������Ƶ���Ϸ��Ŀ¼�£�����")
           
           If ret = 1 Then
            ��ʼ��Ϸ.Picture = LoadPicture(App.Path & "\pic\start.jpg")
           End If
           
        Else
            ret = Shell(App.Path & "\oyesyes.exe", 1)
        End If
    End If
End Sub

Private Sub ��ʼ��Ϸ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ��ʼ��Ϸ.Picture = LoadPicture(App.Path & "\pic\start_focus.jpg")

End Sub

Private Sub ��ʼ��Ϸ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ��ʼ��Ϸ.Picture = LoadPicture(App.Path & "\pic\start_focus.jpg")
    
End Sub

Private Sub ��Ϸ����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        ��Ϸ����.Picture = LoadPicture(App.Path & "\pic\setting_click.jpg")
        If fun_FindProcess("oyesyes.exe") <> 0 Then
            MsgBox "��Ϸ�������У������˳���Ϸ"
        Else
            If Dir(App.Path & "\setup.exe") = "" Then
                ret = MsgBox("�ڵ�ǰĿ¼�£��Ҳ���oyesyes.exe����ѱ������Ƶ���Ϸ��Ŀ¼�£�����")
            Else
                ret = Shell(App.Path & "\setup.exe", 1)
            End If
        End If
    End If
    
End Sub

Private Sub ��Ϸ����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ��Ϸ����.Picture = LoadPicture(App.Path & "\pic\setting_focus.jpg")
    
End Sub

Private Sub ��Ϸ����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ��Ϸ����.Picture = LoadPicture(App.Path & "\pic\setting_focus.jpg")
    
End Sub

Private Sub �뿪_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        �뿪.Picture = LoadPicture(App.Path & "\pic\exit_click.jpg")
    End If
    
End Sub

Private Sub �뿪_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    �뿪.Picture = LoadPicture(App.Path & "\pic\exit_focus.jpg")
    Unload oyesyes��½��
    End
    
End Sub

Private Sub �뿪_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
        �뿪.Picture = LoadPicture(App.Path & "\pic\exit_focus.jpg")

End Sub

Private Sub ������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ��ʼ��Ϸ.Picture = LoadPicture(App.Path & "\pic\start.jpg")
    ��Ϸ����.Picture = LoadPicture(App.Path & "\pic\setting.jpg")
    �뿪.Picture = LoadPicture(App.Path & "\pic\exit.jpg")
    ��θ���.Font.Underline = False
    ��ҳ.Font.Underline = False
    ��ҳ.ForeColor = &H80000012

End Sub


Private Sub ��θ���_click()

    Dim HyperJump As Long
    HyperJump = ShellExecute(0&, vbNullString, WEB, vbNullString, vbNullString, vbNormalFocus)
    
End Sub

Private Sub ��θ���_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ��θ���.Font.Underline = True

End Sub


Private Sub ��ҳ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ��ҳ.Font.Underline = True
    ��ҳ.ForeColor = &H80000002
    
End Sub
