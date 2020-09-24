VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Msn Coding"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   1620
      Left            =   5280
      TabIndex        =   9
      Top             =   4800
      Width           =   4935
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   5295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   4560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "X"
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   2055
   End
   Begin MSComctlLib.ListView Online 
      Height          =   1575
      Left            =   5160
      TabIndex        =   4
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView Members 
      Height          =   2655
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer Verwerking 
      Interval        =   1
      Left            =   6120
      Top             =   840
   End
   Begin MSWinsockLib.Winsock Mainsock 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "64.4.13.58"
      RemotePort      =   1863
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4200
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   6240
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim un
Dim chg
Dim x(1000) As New Form2
Private Sub Command1_Click()
If Command1.Caption = "Connect" Then
  Mainsock.Connect
  Command1.Caption = "Disconnect"
 Else
  Command1.Caption = "Connect"
  Mainsock.Close
  Text1.Text = ""
  List1.Clear
End If
End Sub

Private Sub Form_Load()
un = 0
chg = 0
Members.ColumnHeaders.Add 1, , "Nicks"
Members.ColumnHeaders.Add 2, , "Emails"
Online.ColumnHeaders.Add 1, , "Nicks"
Online.ColumnHeaders.Add 2, , "Emails"
End Sub
Private Sub Mainsock_Connect()
  Mainsock.SendData "VER " & un & " MSNP5 MSNP4 CVRO" & vbCrLf
  Text1.Text = Text1.Text & "Client: " & "VER " & un & " MSNP5 MSNP4 CVRO" & vbCrLf
  Text1.SelStart = Len(Text1)
  un = un + 1
End Sub

Private Sub Mainsock_DataArrival(ByVal bytesTotal As Long)
Dim invoer As String
Mainsock.GetData invoer
Text1.Text = Text1.Text & "Server: " & invoer & vbCrLf
Text1.SelStart = Len(Text1)
For a = 1 To Len(invoer)
   b = InStr(invoer, vbCrLf)
   If b <> 0 Then
      c = Mid(invoer, 1, b - 1)
      invoer = Mid(invoer, b + 1)
      List1.AddItem c
      a = 1
   End If
Next a
End Sub
Private Sub Verwerking_Timer()
If List1.List(0) <> "" Then
  If Mid(List1.List(0), 1, 1) = Chr(10) Then
    List1.List(0) = Mid(List1.List(0), 2)
  End If
  
  If Mid(List1.List(0), 1, 3) = "VER" Then
     Mainsock.SendData "INF " & un & vbCrLf
     Text1.Text = Text1.Text & "Client: " & "INF " & un & vbCrLf
     Text1.SelStart = Len(Text1)
     un = un + 1
     ProgressBar1.Value = 10
  End If
  
  If Mid(List1.List(0), 1, 3) = "INF" And Mid(List1.List(0), 7, 3) = "MD5" Then
     Mainsock.SendData "USR " & un & " MD5  I " & Text2.Text & "@hotmail.com" & vbCrLf
     Text1.Text = Text1.Text & "Client: " & "USR " & un & " MD5  I " & Text2.Text & "@hotmail.com" & vbCrLf
     Text1.SelStart = Len(Text1)
     un = un + 1
     ProgressBar1.Value = 30
  End If
  
  If Mid(List1.List(0), 1, 3) = "XFR" And Mid(List1.List(0), 7, 2) = "NS" Then
       adress = Mid(List1.List(0), 10)
       For a = 1 To Len(adress)
           If Mid(adress, a, 1) = ":" Then
              ip = Left(adress, a - 1)
           End If
       Next a
       Mainsock.Close
       Mainsock.Connect ip, 1863
       un = 1
       ProgressBar1.Value = 0
  End If
  
  If Mid(List1.List(0), 1, 3) = "USR" And Mid(List1.List(0), 7, 5) = "MD5 S" Then
     hatch = Mid(List1.List(0), 13)
     Mainsock.SendData "USR " & un & " MD5 S " & MD5String(hatch & Text3.Text) & vbCrLf
     un = un + 1
     Text1.Text = Text1.Text & "Client: " & "USR " & un & " MD5 S " & MD5String(hatch & Text3.Text) & vbCrLf
     Text1.SelStart = Len(Text1)
     ProgressBar1.Value = 50
  End If
  
  If Mid(List1.List(0), 1, 3) = "MSG" Then
     Mainsock.SendData "CHG " & un & " NLN" & vbCrLf
     un = un + 1
     ProgressBar1.Value = 75
  End If
  
  If Mid(List1.List(0), 1, 3) = "ILN" And Mid(List1.List(0), 7, 3) = "NLN" Then
     user = Mid(List1.List(0), 11)
     For a = 1 To Len(user)
        If Mid(user, a, 1) = " " Then
              email = Left(user, a - 1)
              nick = Mid(user, a + 1)
              With Online
                     .ListItems.Add = nick
                     .ListItems.Item(.ListItems.Count).SubItems(1) = email
              End With
        End If
     Next a
     Mainsock.SendData "XFR " & un & " SB" & vbCrLf
     For e = 0 To List2.ListCount
        If List2.List(e) = email Then
            List2.RemoveItem (e)
        End If
     Next e
     List2.AddItem email
     un = un + 1
     ProgressBar1.Value = 80
  End If
  
  If InStr(List1.List(0), "CHG") > 0 Then
    If chg = 0 Then
     Mainsock.SendData "LST " & un & " RL" & vbCrLf
     un = un + 1
     ProgressBar1.Value = 90
     chg = chg + 1
    End If
  End If
  
  If Mid(List1.List(0), 1, 3) = "LST" And Mid(List1.List(0), 7, 2) = "RL" Then
      user = Mid(List1.List(0), 10)
      b = 0
      For a = 1 To Len(user)
        If b < 3 Then
         c = InStr(user, " ")
         user = Mid(user, c + 1)
         b = b + 1
        End If
      Next a
      b = 0
      For a = 1 To Len(user)
        If b < 1 Then
         c = InStr(user, " ")
         nick = Mid(user, c + 1)
         email = Mid(user, 1, c - 1)
         For c = 1 To Len(user)
           d = InStr(nick, "%20")
           If d > 0 Then
            nick = Mid(nick, 1, d - 1) & " " & Mid(nick, d + 3)
            End If
         Next c
         With Members
                     .ListItems.Add = nick
                     .ListItems.Item(.ListItems.Count).SubItems(1) = email
         End With
         b = b + 1
        End If
      Next a
      ProgressBar1.Value = 100
  End If
  
  If Mid(List1.List(0), 1, 7) = "NLN NLN" Then
      user = Mid(List1.List(0), 9)
      b = 0
      For a = 1 To Len(user)
        If b < 1 Then
         c = InStr(user, " ")
         nick = Mid(user, c + 1)
         email = Mid(user, 1, c - 1)
         b = b + 1
        End If
      Next a
      For a = 1 To Online.ListItems.Count
         If Online.ListItems.Item(a).SubItems(1) = email Then
               Online.ListItems.Remove (a)
         End If
      Next a
      With Online
        .ListItems.Add = nick
        .ListItems.Item(.ListItems.Count).SubItems(1) = email
      End With
     Mainsock.SendData "XFR " & un & " SB" & vbCrLf
     For e = 0 To List2.ListCount
        If List2.List(e) = email Then
            List2.RemoveItem (e)
        End If
     Next e
     List2.AddItem email
     un = un + 1
  End If
  
  If Mid(List1.List(0), 1, 3) = "FLN" Then
      For a = 1 To Online.ListItems.Count
         If Online.ListItems.Item(a).SubItems(1) = Mid(List1.List(0), 5) Then
               Online.ListItems.Remove (a)
         End If
      Next a
  End If
  
  If Mid(List1.List(0), 1, 3) = "XFR" And Mid(List1.List(0), 7, 2) = "SB" Then
     b = InStr(List1.List(0), "CKI")
     c = InStr(List1.List(0), ":")
     ip = Left(List1.List(0), c - 1)
     ip = Mid(ip, 10)
     cki = Mid(List1.List(0), b + 4)
     For a = 0 To List3.ListCount
       If List3.List(a) = List2.List(0) Then
           List3.RemoveItem (a)
       End If
     Next a
     List3.AddItem List2.List(0)
     x(a).unit.Connect ip, 1863
     x(a).Caption = List2.List(0)
     x(a).Tag = cki
     x(a).inout.Caption = "in"
     For a = 0 To List2.ListCount
       If List2.List(a) = List3.List(0) Then
           List2.RemoveItem (a)
       End If
     Next a
  End If
  
  If Mid(List1.List(0), 1, 3) = "RNG" Then
     b = InStr(List1.List(0), " ")
     ringer = Mid(List1.List(0), b + 1)
     b = InStr(ringer, " ")
     hatch = Mid(ringer, 1, b - 1)
     ringer = Mid(ringer, b + 1)
     c = InStr(ringer, ":")
     ip = Left(ringer, c - 1)
     b = InStr(ringer, "CKI")
     ringer = Mid(ringer, b + 4)
     c = InStr(ringer, " ")
     cki = Left(ringer, c - 1)
     cki = cki & " " & hatch
     ringer = Mid(ringer, c + 1)
     b = InStr(ringer, " ")
     email = Left(ringer, b - 1)
     nick = Mid(ringer, b + 1)
     For a = 0 To List3.ListCount
       If List3.List(a) = email Then
           List3.RemoveItem (a)
       End If
     Next a
     List3.AddItem email
     x(a).unit.Connect ip, 1863
     x(a).Caption = email
     x(a).Tag = cki
     x(a).inout.Caption = "out"
     For a = 0 To List2.ListCount
       If List2.List(a) = List3.List(0) Then
           List2.RemoveItem (a)
       End If
     Next a
  End If
  List1.RemoveItem (0)
End If
End Sub
