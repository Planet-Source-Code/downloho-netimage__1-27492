VERSION 5.00
Begin VB.Form frmNetImage 
   Caption         =   "NetImage Example"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   Icon            =   "frmNetImage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt 
      Caption         =   "URL"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtUrl 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "http://a.r.tv.com/cnet.1d/i/ftrs/ne/smrb.gif"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtLocal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "C:\WINDOWS\1stboot.bmp"
      Top             =   4560
      Width           =   3495
   End
   Begin VB.OptionButton opt 
      Caption         =   "Local File"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.OptionButton opt 
      Caption         =   "Microsoft"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.OptionButton opt 
      Caption         =   "Yahoo! Chat"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.OptionButton opt 
      Caption         =   "Google"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Image"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin prjNetImage.NetImage NetImage1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frmNetImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NetImage1.TempPath = App.Path
NetImage1.Stretch = False

If opt(0).Value = True Then
 NetImage1.Url = "http://www.google.com/images/title_homepage4.gif"
ElseIf opt(1).Value = True Then
 NetImage1.Url = "http://us.i1.yimg.com/us.yimg.com/i/us/ch/ch4.gif"
ElseIf opt(2).Value = True Then
 NetImage1.Url = "http://www.microsoft.com/library/homepage/images/mslogo-blue.gif"
ElseIf opt(3).Value = True Then
 NetImage1.Url = txtUrl.Text
ElseIf opt(4).Value = True Then
 NetImage1.Url = "file://" & txtLocal.Text
End If

NetImage1.GetImage
End Sub

Private Sub NetImage1_Done()
Caption = "NetImage Example - [Finished]"
End Sub

Private Sub NetImage1_Error(ByVal Object As String, ByVal Descripton As String)
Caption = Object$ & ": " & Description$
End Sub

Private Sub NetImage1_Size(ByVal h As Long, ByVal w As Long)
NetImage1.Height = h& * Screen.TwipsPerPixelY
NetImage1.Width = w& * Screen.TwipsPerPixelX
End Sub

Private Sub NetImage1_Status(ByVal Status As String)
Caption = Status$
End Sub

Private Sub opt_Click(Index As Integer)
txtUrl.Enabled = opt(3).Value
txtLocal.Enabled = opt(4).Value
End Sub
