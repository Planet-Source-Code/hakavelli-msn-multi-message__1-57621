VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Msn Multi-Message"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":058A
   ScaleHeight     =   6135
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D2D9D3&
      Caption         =   "Send"
      Height          =   375
      Left            =   4440
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D2D9D3&
      Caption         =   "Clear"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3510
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A61E8
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A66AE
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A6B74
            Key             =   "main"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "Form1.frx":A703A
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   7560
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   0
      Picture         =   "Form1.frx":A75C4
      Top             =   0
      Width           =   8295
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   3960
      Y1              =   6120
      Y2              =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   4920
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------- Declarations ---------------------
Private Const SW_SHOW = 5
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim msn As MessengerAPI.Messenger
Dim con As IMessengerContact
'-------- Declarations ---------------------

Private Sub Form_Load()
TreeView1.Nodes.Add , , "Onlinepeople", "Available Contacts", "up"
Set msn = New MessengerAPI.Messenger
RefreshContacts
End Sub

Private Sub RefreshContacts()
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , "Onlinepeople", "Online/Available Contacts", "up"
TreeView1.Nodes.Item(1).Expanded = True
For Each con In msn.MyContacts
If con.Status = MISTATUS_OFFLINE Then
'exclude offline contacts
Else
TreeView1.Nodes.Add "Onlinepeople", tvwChild, , con.SigninName, "main"
End If
Next
TreeView1.Nodes.Item(1).Bold = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Label11_Click()
End
End Sub

Private Sub Label12_Click()
Me.WindowState = 1
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Key = "Onlinepeople" Then
If Node.Expanded = True Then
Node.Expanded = False
Node.Image = "down"
Node.SelectedImage = "down"
Else
Node.Expanded = True
Node.Image = "up"
Node.SelectedImage = "up"
End If
End If
If Node.Key = "Onlinepeople" Then
Text11.Text = ""
Else
Text11.Text = TreeView1.SelectedItem
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Command1_Click()
If Text11.Text = "" Then
MsgBox "Please select a contact", , "Invalid Contact"
Else
msn.InstantMessage Text11.Text
SendKeys Text1.Text
SendKeys "{ENTER}"
SendKeys Text2.Text
SendKeys "{ENTER}"
SendKeys Text3.Text
SendKeys "{ENTER}"
SendKeys Text4.Text
SendKeys "{ENTER}"
SendKeys Text5.Text
SendKeys "{ENTER}"
SendKeys Text6.Text
SendKeys "{ENTER}"
SendKeys Text7.Text
SendKeys "{ENTER}"
SendKeys Text8.Text
SendKeys "{ENTER}"
SendKeys Text9.Text
SendKeys "{ENTER}"
SendKeys Text10.Text
SendKeys "{ENTER}"
End If
End Sub
