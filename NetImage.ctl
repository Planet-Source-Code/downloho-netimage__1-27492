VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl NetImage 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   86
   Begin VB.Timer tmrOut 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox pic 
      Align           =   2  'Align Bottom
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   30
      Index           =   3
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   1290
      TabIndex        =   3
      Top             =   1380
      Width           =   1290
   End
   Begin VB.PictureBox pic 
      Align           =   4  'Align Right
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   1350
      Index           =   2
      Left            =   1260
      ScaleHeight     =   1350
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   30
      Width           =   30
   End
   Begin VB.PictureBox pic 
      Align           =   3  'Align Left
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1350
      Index           =   1
      Left            =   0
      ScaleHeight     =   1350
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   30
      Width           =   30
   End
   Begin VB.PictureBox pic 
      Align           =   1  'Align Top
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   30
      Index           =   0
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   1290
      TabIndex        =   0
      Top             =   0
      Width           =   1290
   End
   Begin VB.Image imgMain 
      Height          =   720
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   690
   End
   Begin VB.Image img 
      Height          =   240
      Left            =   120
      Picture         =   "NetImage.ctx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "NetImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Done()
Event Error(ByVal Object As String, ByVal Descripton As String)
Event Size(ByVal h As Long, ByVal w As Long)
Event Status(ByVal Status As String)

Dim mUrl As String, mPath As String

Dim Go As Boolean, lFileNum As Long, lTotal(1) As Long

Private Sub Connect(ByVal Url As String)
'Example: Call Connect("http://www.yahoo.com")
Dim sHost As String, sPath As String
Dim i As Integer

 If Left$(LCase$(Url$), 7) = "http://" Then Url$ = Mid$(Url$, 8)
 If InStr(Url$, "/") = 0 Then i% = Len(Url$) + 1 Else i% = InStr(Url$, "/")
 
 sHost$ = Left$(Url$, i% - 1)
 sPath$ = IIf(i% = Len(Url$) + 1, "/", Mid$(Url$, i%))

 With sckMain
 .Close
 .RemoteHost = sHost$
 .RemotePort = 80
 .Tag = sPath$
 .Connect
End With
RaiseEvent Status("Connecting")
End Sub

Private Function FileExist(ByVal File As String) As Boolean
On Error GoTo 1
 Call FileLen(File$)
 FileExist = True
Exit Function
1
 FileExist = False
End Function

Private Function FileExt() As String
Dim s$, v As Variant, b As Boolean
s$ = Right$(sckMain.Tag, 3)

For Each v In Array("bmp", "gif", "jpg")
 If LCase$(s$) = v Then b = True: Exit For
Next v

 If b = True Then FileExt$ = s$ Else FileExt$ = "jpg"
End Function

Public Sub GetImage(Optional ByVal Url As String, Optional ByVal TempPath As String)
If Url$ <> "" Then mUrl$ = Url$
If TempPath$ <> "" Then mPath$ = TempPath$
If mPath$ = "" Then RaiseEvent Error("Connect", "Temp Path not assigned, connection aborted."): Exit Sub
lTotal&(0) = 0: lTotal&(1) = 0
Go = False
If Left$(LCase$(mUrl$), 7) = "http://" Then
 Call Connect(mUrl$)
ElseIf Left$(LCase$(mUrl$), 6) = "ftp://" Then
 RaiseEvent Error("Connect", "FTP not supported, connection aborted.")
ElseIf Left$(LCase$(mUrl$), 7) = "file://" Then
 If FileExist(Mid$(mUrl$, 8)) = True Then _
    Set Me.Picture = LoadPicture(Mid$(mUrl$, 8)) Else _
    RaiseEvent Error("Connect", "File not found: " & Mid$(mUrl$, 8))
Else
 Call Connect(mUrl$)
End If
End Sub

Public Property Set Picture(p As Picture)
Dim i%
Set imgMain.Picture = p
If imgMain.Stretch = False Then RaiseEvent Size(imgMain.Height, imgMain.Width)

imgMain.Visible = True
img.Visible = False
For i% = 0 To 3
 pic(i%).Visible = False
Next i%
End Property

Private Function RetrieveHTML(ByVal Path As String, ByVal Host As String) As String
Dim sTxt As String
    
    sTxt$ = "GET " & Path$ & " HTTP/1.0" & vbCrLf
    sTxt$ = sTxt$ & "Referer: http://" & Host$ & vbCrLf
    sTxt$ = sTxt$ & "Connection: Keep-Alive" & vbCrLf
    sTxt$ = sTxt$ & "User-Agent: Default" & vbCrLf
    sTxt$ = sTxt$ & "Host: " & Host$ & vbCrLf
    sTxt$ = sTxt$ & "Accept: */*" & vbCrLf
    sTxt$ = sTxt$ & "Accept-Language: en" & vbCrLf
    RetrieveHTML$ = sTxt$ & vbCrLf
    
End Function

Public Property Get Stretch() As Boolean
Stretch = imgMain.Stretch
End Property

Public Property Let Stretch(b As Boolean)
imgMain.Stretch = b
End Property

Public Property Get TempPath() As String
TempPath$ = mPath$
End Property

Public Property Let TempPath(b As String)
mPath$ = b$
End Property

Public Property Get Url() As String
Url$ = mUrl$
End Property

Public Property Let Url(b As String)
mUrl$ = b$
End Property

Private Sub UserControl_Resize()
If Width < 480 Then Width = 480
If Height < 480 Then Height = 480
If imgMain.Stretch = True Then
 imgMain.Width = Width / Screen.TwipsPerPixelX
 imgMain.Height = Height / Screen.TwipsPerPixelY
End If
End Sub

Private Sub sckMain_Close()
On Error Resume Next
Dim i%
tmrOut.Enabled = False
Close lFileNum&: sckMain.Close
Set imgMain.Picture = LoadPicture(mPath$ & IIf(Right$(mPath$, 1) = "\", "", "\") & "tmp." & FileExt$)
img.Visible = False
imgMain.Visible = True
For i% = 0 To 3
 pic(i%).Visible = False
Next i%
If imgMain.Stretch = False Then RaiseEvent Size(imgMain.Height, imgMain.Width)
RaiseEvent Done
End Sub

Private Sub sckMain_Connect()
Call sckMain.SendData(RetrieveHTML$(sckMain.Tag, sckMain.RemoteHost))
lFileNum& = FreeFile()
Open mPath$ & IIf(Right$(mPath$, 1) = "\", "", "\") & "tmp." & FileExt$ For Binary Access Write As lFileNum&
tmrOut.Enabled = False
tmrOut.Enabled = True
RaiseEvent Status("Connected")
End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
Dim s As String
 Call sckMain.GetData(s$, vbString)

If mPath$ <> "" Then
 
 tmrOut.Enabled = False
 tmrOut.Enabled = True

 If Go = True Then
  Put lFileNum&, , s$
  lTotal&(1) = lTotal&(1) + Len(s$)
  If lTotal&(0) = lTotal&(1) Then Call sckMain_Close: Exit Sub
 Else

   If InStr(LCase$("f" & s$), "http") <> 0 And Mid(s$, 10, 3) <> "200" Then
     RaiseEvent Error("Socket", "File not found.")
     Close lFileNum&
     sckMain.Close
     Exit Sub
   End If

  If InStr(LCase$(s$), "content-type:") <> 0 Then
   If InStr(LCase$(s$), "content-length:") <> 0 Then
    Dim i%, d%
    i% = InStr(LCase$(s$), "content-length:") + Len("content-length:") + 1
    d% = InStr(i% + 1, LCase$(s$), vbCrLf)
    lTotal&(0) = CLng(Trim$(Mid$(s$, i%, d% - i%)))
    If lTotal&(0) = 0 Then lTotal&(0) = 10000000
   Else
    lTotal&(0) = 10000000
   End If
    Go = True
    Put lFileNum&, 1, Mid$(s$, InStr(s$, vbCrLf & vbCrLf) + 4)
    lTotal&(1) = lTotal&(1) + Len(Mid$(s$, InStr(s$, vbCrLf & vbCrLf) + 4))
  Else
    Go = True
    lTotal&(1) = lTotal&(1) + Len(s$)
    Put lFileNum&, 1, s$
    lTotal&(0) = 10000000
  End If
 End If
If lTotal&(0) = lTotal&(1) Then Call sckMain_Close: Exit Sub
Else
 sckMain.Close
 RaiseEvent Error("Socket", "No Temp Path set.")
 Exit Sub
End If
RaiseEvent Status("Recieving Data " & lTotal&(1) & " of " & IIf(lTotal&(0) = 10000000, "N/A", lTotal&(0)))
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call sckMain.Close
RaiseEvent Error("Socket", "d" & Description$)
End Sub

Private Sub tmrOut_Timer()
Call sckMain_Close
tmrOut.Enabled = False
End Sub
