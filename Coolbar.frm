VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Most Simple CoolBar"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Coolbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   688
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make CoolBar!"
      Height          =   435
      Left            =   105
      TabIndex        =   2
      Top             =   1365
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Damian Janowski - E-mail: jano@sinai.com.ar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   1995
      Width           =   4485
   End
   Begin VB.Label Label2 
      Caption         =   "No subclassing or complicated processes, just a few lines of code and one function to make the CoolBar from a common toolbar."
      Height          =   750
      Left            =   1680
      TabIndex        =   3
      Top             =   1155
      Width           =   2955
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4095
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":1292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":13A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":14B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":15C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":16DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":17EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":18FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":1A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":1B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Coolbar.frx":1C34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "The Most Simple CoolBar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  'Toolbar Const
  Private Const WM_USER = &H400
  Private Const TBSTYLE_TRANSPARENT = &H8000
  Private Const TBSTYLE_FLAT = &H800
  Private Const TB_SETSTYLE = (WM_USER + 56)
  Private Const TB_GETSTYLE = (WM_USER + 57)
  Private Const TBSTYLE_LIST = &H1000
  Private Const CCS_NODIVIDER = &H40

  Private Declare Function FindWindowEx Lib "user32" _
          Alias "FindWindowExA" _
          (ByVal hWnd1 As Long, _
          ByVal hWnd2 As Long, _
          ByVal lpsz1 As String, _
          ByVal lpsz2 As String) As Long
  
  Private Declare Function SendTBMessage Lib "user32" _
          Alias "SendMessageA" _
          (ByVal hwnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Integer, _
          ByVal lParam As Any) As Long

Public Sub MakeToolbarFlat(Tb As Object)
  Dim Style As Long
  Dim lRet As Long
  Dim ToolbarHandle As Long

  ToolbarHandle = FindWindowEx(Tb.hwnd, 0&, "ToolbarWindow32", vbNullString)

  Style = SendTBMessage(ToolbarHandle, TB_GETSTYLE, 0&, 0&)
  Style = Style Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT Or CCS_NODIVIDER
  lRet = SendTBMessage(ToolbarHandle, TB_SETSTYLE, 0, Style)

  Tb.Refresh
End Sub

Private Sub Command1_Click()
    Call MakeToolbarFlat(Toolbar1)
End Sub
