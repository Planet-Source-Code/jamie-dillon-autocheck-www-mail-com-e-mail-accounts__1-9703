VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Check Mail.com"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Set Sound"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.wav|Wave files|"
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   1080
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   1920
      Max             =   60
      Min             =   1
      TabIndex        =   11
      Top             =   1080
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Text            =   "5"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check Now"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5040
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   15
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2415
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   4260
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save "
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1155
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Check Every:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1155
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Login:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code checks email accounts with the web based www.mail.com
'Its nothing but a cheap hack really, nothing educational here =)
'What it does it build a string which will log you into YOUR account.
'It then retrieves your inbox page and checks for strings specifying no
'new email found, such as "0 new messages" etc. It then alerts the
'users if any new mail is found. It will do this at intervals specified
' by the user, in minute increments
'It also checks if the user is connected 'to the internet
'(using a DUN - i may update this so it works for cable, LAN etc)
'Saves the password to the registry and encrypts it, for extra sexurity

'Jamie Dillon, benighted@post.com (part of mail.com) 12/7/00

Dim strHTML As String
Dim checkTime As Date
Dim strCheck As String
Dim strPass As String
Dim strWav As String



Private Sub Command1_Click()
SaveSetting App.EXEName, "Login", "Login", Text1.Text 'save login and password to registry
strPass = Encrypt(12, 1, 1, 9, 1, False, Text2.Text) 'encrpyt password
SaveSetting App.EXEName, "Login", "Passwd", strPass 'save encrypted password to registry
End Sub

Private Sub Command2_Click()
If IsConnected = True Then
    CheckEmail
    checkTime = Now
Else
    MsgBox "Please connect to the Internet first."
End If
End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "Wave Files(*.wav)|*.wav|"
CommonDialog1.ShowOpen
strWav = CommonDialog1.FileName
SaveSetting App.EXEName, "Login", "Sound", strWav
End Sub

Private Sub Form_Load()
Me.WindowState = 1
strWav = GetSetting(App.EXEName, "Login", "Sound")
strInterval = GetSetting(App.EXEName, "Login", "Interval") 'gets previously set time interval
If strInterval <> "" Then 'no interval has been set before
    Text3.Text = Val(strInterval)
End If
Me.Caption = "Check Mail.com - " & Text1.Text
CheckEmail
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If RunningInTray Then
    If x = 7725 Then 'double clicked tray icon with left mouse button
        Me.WindowState = vbNormal 'restore form from system tray
        Me.Show
        Me.SetFocus
        RemoveIcon Me
    End If
End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ' This code hides the Form and puts the icon in the tray.
        Me.Hide
        ShowIcon Me
    End If

End Sub

Function CheckEmail()
On Error Resume Next
If IsConnected Then
Text1.Text = GetSetting(App.EXEName, "Login", "Login")  'fill in login and password boxes from registry keys
strPass = GetSetting(App.EXEName, "Login", "Passwd") 'retrieves password
Text2.Text = Decrypt(12, 1, 1, 9, 1, False, strPass) 'decrypyts password into textbox
If Text1.Text = "" Or Text2.Text = "" Then
'havent filled in login/password, dont do anything
Else
    strHTML = "<form name=" & Chr(34) & "main" & Chr(34) & " action=" & Chr(34) & "http://www.mail.com/login/mailcom/login.jhtml" & Chr(34) & " method=" & Chr(34) & "POST" & Chr(34) & ">" & _
    "<input type=hidden name=sn val=em>" & _
    "<input type=text name=alias value=" & Chr(34) & Text1.Text & Chr(34) & " size=18 >" & _
    "<input type=password name=pw size=18 value = " & Text2.Text & ">" & _
    "<input type= Submit value=LOG IN>" & _
    "<INPUT type=hidden name=od value=www>" & _
    "<SCRIPT>main.submit()</SCRIPT>" & _
    "</form>"
    ' this builds a string to recreate the login page www.mail.com,
    'chopping all the unessacary stuff or course.
    'The values of login and password are also put into this string
    'the javascript line is used to stimulate clicking the submit button
    'thus automating the procedure
    Open App.path & "\Mail.htm" For Output As #1
        Print #1, strHTML 'string is saved into a file
    Close #1
    WebBrowser1.Navigate (App.path & "\Mail.htm") 'htm file consisting of
    'previously built string is opened into a webbrowser control.
    'beause of the java script in it, this will automatically login
    'into your account once the page loads
    strCheck = Inet1.OpenURL("http://www.mail.com/mailcom/mailbox.jhtml")
    'because you stimulated logging into your account using the webbrowser control,
    'going to http://www.mail.com/mailcom/mailbox.jhtml will show the contents
    'of your inbox, so retrieve that page using Inet
    If InStr(strCheck, "0 new messages") <> 0 Or InStr(strCheck, "/readoldmessages.gif") <> 0 Then
    'check retrieved page to see if it contains strings which would mean
    'no new email, ie. "0 new messages".
    'if above condition returns anything but 0, it means the string
    '"0 new messages" was found in the page, thus theres no new email.
    ' if it returns 0, it means it couldnt find "0 new messages", obviously if theres 0 new messages
    'there must be 1 or more new messages
    Else 'new email was found
        ret = sndPlaySound(strWav, 1) 'plays wav file using API
        Me.WindowState = vbNormal
        If MsgBox("New mail found for " & Text1.Text & vbCrLf & "Do you wish to goto mail.com?", vbYesNo, "NEW MAIL FOUND") = vbYes Then
        'msgbox asking if user wants to goto www.mail.com to read new mail
            ret = Shell("Start.exe " & "http://www.mail.com/", 0) 'yes they do, so open up page in browser
        End If
    End If
    Label3.Caption = "Last checked " & Text1.Text & " at " & DateTime.Time 'update time last checked
    Kill App.path & "\Mail.htm" 'deletes file, as its not needed anymore
End If
End If
checkTime = Now
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.path & "Mail.htm"
End Sub

Private Sub Text3_Change()

If Val(Text3.Text) > 120 Then 'sets max time of 2 hours between checks, no reason for this really...
    Text3.Text = 60
    Text3.SelStart = 2
End If
HScroll1.Value = Val(Text3.Text)
SaveSetting App.EXEName, "Login", "Interval", Text3.Text 'save interval to registry so it can be retireved next time program starts
End Sub

Private Sub HScroll1_Change()
Text3.Text = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
If DateAdd("n", Text3, checkTime) < Now Then 'checks to see if the
    'difference between the time e-mail was last checked and the current time
    'is equal to the number displayed in text3.
    CheckEmail 'if it is, then specified duration has elapsed, so check email again
End If
End Sub


