VERSION 5.00
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "mshtml.dll"
Begin VB.Form Form1 
   Caption         =   "VTSTech-WebScraper v0.0.1-r00"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      CausesValidation=   0   'False
      Height          =   1140
      Left            =   75
      TabIndex        =   11
      Top             =   3600
      Width           =   7380
      Scrollbar       =   0   'False
      URL             =   "http://ad.a-ads.com/707814?size=468x60"
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Text            =   "capture-len-r"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Text            =   "capture-len-f"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "target"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "http://www.example.com/target/file.html"
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "scrape"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "GitHub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "www.vts-tech.org"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Written by VTSTech//Veritas Technical Solutions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1815
      TabIndex        =   8
      Top             =   3120
      Width           =   3510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1VTSgzD24bjkSGdD7kvauxkxHZ4yiwhdU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2055
      TabIndex        =   7
      Top             =   3360
      Width           =   3030
   End
   Begin VB.Label Label1 
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Build, x, y, z, tmp, target, link, output
Dim remotehost, httpdata, httpreq, ua, FSO, forward, backward
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function DoScrape(link)
Label1.Caption = "Downloading..."
If FSO.FileExists(VB.App.Path & "\wget.exe") Then
tmp = "--user-agent=" & ua & " --no-check-certificate " & link
Shell ("cmd.exe /c " & VB.App.Path & "\wget.exe " & tmp & " --output-document=" & VB.App.Path & "\temp.html"), vbHide
Sleep (1000)
Label1.Caption = "Parsing..."
a = DoParse()
End If
End Function
Function DoCleanup()
Label1.Caption = "Done." & z & " hits."
Shell ("cmd.exe /c del " & VB.App.Path & "\temp.html"), vbHide
End Function
Function DoParse()
z = 0
If FSO.FileExists(VB.App.Path & "\temp.html") Then
    Label1.Caption = "Checking for Target ..."
    Sleep (1000)
    Open VB.App.Path & "\temp.html" For Input As #1
        Do
            Line Input #1, tmp
            y = Len(tmp)
            For x = 1 To Len(tmp)
                If Mid$(tmp, x, Len(target)) = target Then
                    z = z + 1
                    Label1.Caption = "Target Found! Displaying results..."
                    'MsgBox "Found! At " & x
                    Text1.Text = Text1.Text & vbCrLf & "---" & vbCrLf & Mid(tmp, x, forward)
                    If (x - backward) > 0 Then Text1.Text = Text1.Text & vbrcrlf & Mid(tmp, x - backward, backward)
                End If
            Next x
    Loop While Not EOF(1)
    Close #1
    Sleep (1000)
End If
a = DoCleanup()
End Function
Private Sub Command1_Click()
Text1.Text = ""
link = Text2.Text
target = Text3.Text
forward = Text4.Text
backward = Text5.Text
If Len(link) > 1 Then
    If Mid(link, 1, 7) = "http://" Or Mid(link, 1, 8) = "https://" Then  'good link
        If Len(target) > 1 Then
            DoScrape (link)
            End If
    Else    'bad link
    MsgBox "link must start with http or https"
    End If
End If 'link len > 1
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Form_Load()
Scriptlet1.Width = 7025
Scriptlet1.Height = 905
Set FSO = CreateObject("Scripting.FileSystemObject")
Build = "0.0.1-r01"
ua = Chr(34) & "Mozilla/5.0 (Windows NT 5.1; WOW64; rv:52.0) Gecko/20100101 Firefox/52.0" & Chr(34)
Form1.Caption = "VTSTech-WebScraper v" & Build
Label1.Caption = "Status: Idle"
Text1.Text = "target is search string" & vbCrLf & "capture-len-f will output this many chars forward" & vbCrLf & "capture-len-r will output this many chars backward"
If FSO.FileExists(VB.App.Path & "\wget.exe") = False Then
    MsgBox "Fatal Error: wget.exe not found! Download and Install again!"
    Unload Form1
End If
End Sub

Private Sub Label2_Click()
Clipboard.SetText Label2.Caption
MsgBox "BitCoin Address copied to clipboard!"
End Sub

Private Sub Label4_Click()
Shell ("cmd.exe /c start http://www.vts-tech.org"), vbHide
End Sub

Private Sub Label5_Click()
Shell ("cmd.exe /c start http://www.github.com/Veritas83"), vbHide
End Sub
