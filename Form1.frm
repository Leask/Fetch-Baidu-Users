VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   14625
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text_Box 
      Height          =   3660
      Left            =   6165
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   4770
      Width           =   8250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   600
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   2940
   End
   Begin SHDocVwCtl.WebBrowser Web_Core 
      Height          =   4380
      Left            =   6210
      TabIndex        =   0
      Top             =   135
      Width           =   8250
      ExtentX         =   14552
      ExtentY         =   7726
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
   Begin VB.Label Label 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   270
      TabIndex        =   3
      Top             =   1710
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BD_UID As Long


Private Sub Command1_Click()

'BD_UID = 4566
BD_UID = 10000000

Web_Nav BD_UID

End Sub


Private Sub Web_Nav(BD_UID)
Label.Caption = BD_UID
Web_Core.Navigate "http://im.baidu.com/invite/groupauth.php?uid=" & BD_UID
End Sub


Private Sub Web_Core_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If BD_UID = 0 Then Exit Sub
Text_Box = Web_Core.Document.body.innertext
Open App.Path & "\baidu\" & BD_UID & ".txt" For Output As #1
    Print #1, Text_Box.Text
Close #1

If BD_UID = 70020000 Then Exit Sub

BD_UID = BD_UID + 1
 Web_Nav (BD_UID)
End Sub

