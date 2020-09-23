VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "HTTP Subnet Scanner"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTimeOut 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   8
      Text            =   "2"
      Top             =   960
      Width           =   615
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1320
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1920
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdStop 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stop"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2280
      Top             =   960
   End
   Begin VB.TextBox TxtInfo 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton CmdGo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scan"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Text            =   "1"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Text            =   "163"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Text            =   "46"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Text            =   "80"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Time Out (In Seconds)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGo_Click()
If TxtIP(3).Text = 255 Then CmdStop_Click 'make sure it does not get past 255



TxtIP(3).Text = TxtIP(3).Text + 1 'change the ip

Winsock1.RemoteHost = TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text 'set the ip
Winsock1.RemotePort = 80
Winsock1.Connect 'try to connect to host

TimeOut.Enabled = False 'stop the time out (should already be stopped, but lets just make sure!)
TimeOut.Interval = TxtTimeOut.Text * 1000 'set the time out

TimeOut.Enabled = True 'enable the timeout timer

End Sub

Private Sub CmdStop_Click()
Winsock1.Close 'close winsock
TimeOut.Enabled = False 'stop the time out

End Sub

Private Sub TimeOut_Timer()
ConnectionClose 'goto connectionclose sub
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next 'if winsock connects...

TimeOut.Enabled = False 'disable the time out
site = Inet1.OpenURL("http://" + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text, icByteArray) 'set the site

servers = Inet1.GetHeader("server") 'grab the server header (if there is one)

'you can grab any header you want, as long as it's there!, normal headers include, Content-type, Content-length, and Expires
'if you want to grab all the headers, then just use servers = Inet1.GetHeader()
'the headers you can get from kazaa are,
'X-Kazaa-Username
'X-Kazaa-Network
'X-Kazaa-IP
'X-Kazaa-SupernodeIP



If servers = "" Then 'if there isn't a server header, chances are it's going to be kazaa

    site = Inet1.OpenURL("http://" + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text, icByteArray)
    
    user = Inet1.GetHeader("X-Kazaa-Username")
    'so we try and get the kazaa username, (nothing important, just somthing i thought could be fun?!)
    TxtInfo.Text = TxtInfo.Text + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text + "     " + "Kazaa Username: " + user + vbNewLine

Else
    'show the http server type
    TxtInfo.Text = TxtInfo.Text + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text + "     " + servers + vbNewLine
End If


TxtInfo.SelStart = Len(TxtInfo.Text) 'scroll down the txtinfo box

Winsock1.Close 'close winsock
CmdGo_Click 'start again


End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ConnectionClose 'call connection close sub

End Sub


Private Sub ConnectionClose()

TxtInfo.Text = TxtInfo.Text + TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & TxtIP(3).Text + "     " + "NO Server" + vbNewLine 'obvisuly there was no server
TxtInfo.SelStart = Len(TxtInfo.Text) 'scroll down
TimeOut.Enabled = False 'stop time out
Winsock1().Close 'close winsock (if not already done)
CmdGo_Click 'start again
End Sub

