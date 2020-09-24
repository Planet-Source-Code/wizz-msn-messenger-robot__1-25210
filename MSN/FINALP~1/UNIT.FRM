VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6645
   LinkTopic       =   "Form2"
   ScaleHeight     =   3660
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Timer unitadd 
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin MSWinsockLib.Winsock unit 
      Left            =   0
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label inout 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim un

Private Sub Form_Load()
un = 1
End Sub

Private Sub unit_Connect()
  If inout.Caption = "in" Then
     unit.SendData "USR " & un & " " & Form1.Text2.Text & "@hotmail.com " & Me.Tag & vbCrLf
     Form1.Text1.Text = Form1.Text1.Text & "USR " & un & " " & Form1.Text2.Text & "@hotmail.com " & Me.Tag & vbCrLf
     Form1.Text1.SelStart = Len(Form1.Text1)
  End If
  
  If inout.Caption = "out" Then
     unit.SendData "ANS " & un & " " & Form1.Text2.Text & "@hotmail.com " & Me.Tag & vbCrLf
     Form1.Text1.Text = Form1.Text1.Text & "ANS " & un & " " & Form1.Text2.Text & "@hotmail.com " & Me.Tag & vbCrLf
     Form1.Text1.SelStart = Len(Form1.Text1)
  End If
End Sub

Private Sub unit_DataArrival(ByVal bytesTotal As Long)
Dim invoer As String
unit.GetData invoer
If InStr(invoer, MSG) Then
  List1.AddItem invoer
  Exit Sub
End If
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

Private Sub unitadd_Timer()
     Dim mess As String
     Dim mess2 As String
If List1.List(0) <> "" Then
  If Mid(List1.List(0), 1, 3) = "USR" And Mid(List1.List(0), 7, 2) = "OK" Then
     unit.SendData "CAL " & un & " " & Me.Caption & vbCrLf
     Form1.Text1.Text = Form1.Text1.Text & "CAL " & un & " " & Me.Caption & vbCrLf
     Form1.Text1.SelStart = Len(Form1.Text1)
     un = un + 1
  End If
  
  If Mid(List1.List(0), 1, 3) = "JOI" Then
     mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & "Hi i'am Wizzbot, your personal MSN bot" & vbCrLf & "Type: /Menu voor opties"
     mess2 = "MSG " & un & " N " & Len(mess) & vbCrLf & mess
     un = un + 1
     unit.SendData mess2
  End If
  
  If Mid(List1.List(0), 1, 3) = "IRO" Then
     mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & "You called me master??"
     mess2 = "MSG " & un & " N " & Len(mess) & vbCrLf & mess
     un = un + 1
     unit.SendData mess2
  End If
  
  If Mid(List1.List(0), 1, 3) = "MSG" Then
      bericht = List1.List(0)
     If InStr(bericht, "/info") Then
          mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & "Msn bot made by:" & vbCrLf & "Wilco Burggraaf" & vbCrLf & "mailto:xxxto_me@hotmail.com"
          mess2 = "MSG " & un & " N " & Len(mess) & vbCrLf & mess
          un = un + 1
          unit.SendData mess2
     End If
     If InStr(bericht, "/menu") Then
          mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & "--->MENU<---" & vbCrLf & "  (h) Nieuws" & vbCrLf & "  (h) Forum" & vbCrLf & "  (h) Notice" & vbCrLf & "  (h) Agenda" & vbCrLf & Date & " " & Time
          mess2 = "MSG " & un & " N " & Len(mess) & vbCrLf & mess
          un = un + 1
          unit.SendData mess2
     End If
     If InStr(bericht, "/nieuws") Then
          mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & "--->NIEUWS<---" & vbCrLf & " De Bot word druk aangepast dus hopelijk zal alles draaien in het begin van het nieuwe schooljaar," & vbCrLf & " gegroet Wizz"
          mess2 = "MSG " & un & " N " & Len(mess) & vbCrLf & mess
          un = un + 1
          unit.SendData mess2
     End If
     If InStr(bericht, "/notice") Then
        b = 0
        For a = 1 To Form1.Members.ListItems.Count
         If Form1.Members.ListItems.Item(a) <> "" Then
          memb = memb & "(h) " & a & " " & Form1.Members.ListItems.Item(a) & " " & Form1.Members.ListItems.Item(a).SubItems(1) & vbCrLf
          b = b + 1
          If b = 10 Or a = Form1.Members.ListItems.Count Then
            mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & memb
            mess2 = "MSG " & un & " N " & Len(mess) & vbCrLf & mess
            un = un + 1
            unit.SendData mess2
            memb = ""
            b = 0
          End If
         End If
        Next a
     End If
  End If
  
  If Mid(List1.List(0), 1, 3) = "BYE" Then
     For a = 0 To Form1.List3.ListCount
       If Form1.List3.List(a) = Mid(List1.List(0), 5) Then
           Form1.List3.RemoveItem (a)
           unit.Close
       End If
     Next a
     un = un + 1
  End If
  List1.RemoveItem (0)
End If
End Sub
