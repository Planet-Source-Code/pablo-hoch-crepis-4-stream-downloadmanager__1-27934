VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Crepis-Downloadmanager"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6375
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fDL 
      Caption         =   "Download wird ausgeführt"
      Height          =   3135
      Left            =   60
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   6255
      Begin MSWinsockLib.Winsock wDL 
         Index           =   1
         Left            =   240
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cSTOP 
         Caption         =   "Download anhalten"
         Height          =   375
         Left            =   3480
         TabIndex        =   44
         Top             =   2640
         Width           =   2655
      End
      Begin MSComctlLib.ProgressBar barSTREAM 
         Height          =   135
         Index           =   1
         Left            =   4080
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar barKOMPLETT 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar barSTREAM 
         Height          =   135
         Index           =   2
         Left            =   4080
         TabIndex        =   41
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar barSTREAM 
         Height          =   135
         Index           =   3
         Left            =   4080
         TabIndex        =   42
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar barSTREAM 
         Height          =   135
         Index           =   4
         Left            =   4080
         TabIndex        =   43
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSWinsockLib.Winsock wDL 
         Index           =   2
         Left            =   720
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wDL 
         Index           =   3
         Left            =   1200
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wDL 
         Index           =   4
         Left            =   1680
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wDL 
         Index           =   0
         Left            =   2280
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label st 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label lBYTES 
         Caption         =   "0 B"
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   39
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lBYTES 
         Caption         =   "0 B"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   38
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lBYTES 
         Caption         =   "0 B"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   37
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lBYTES 
         Caption         =   "0 B"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   36
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label st 
         Caption         =   "???"
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   35
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label st 
         Caption         =   "???"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   34
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label st 
         Caption         =   "???"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   33
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label st 
         Caption         =   "???"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Zentriert
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Zentriert
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Zentriert
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Bytes"
         Height          =   255
         Left            =   3480
         TabIndex        =   27
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Status"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Stream"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lRESUME 
         Caption         =   "???"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Forsetzen unterstützt:"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lNOCH 
         Caption         =   "???"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Verbleibende Zeit:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lSIZE 
         Caption         =   "0 Byte von 0 Byte"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Dateigröße:"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lSPEED 
         Caption         =   "???"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Geschwindigkeit:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lURL 
         Caption         =   "http://"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame fNEW 
      Caption         =   "Neuer Download-Auftrag"
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   6255
      Begin VB.CheckBox oRESUME 
         Caption         =   "Download fortsetzen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cSTARTDOWNLOAD 
         Caption         =   "&Download starten"
         Height          =   495
         Left            =   2880
         TabIndex        =   12
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Frame fHEADER 
         Caption         =   "Erweiterte Header-Optionen"
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
         Begin VB.TextBox tCOOKIEWERT 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox tCOOKIENAME 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox tBROWSER 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Text            =   "Crepis/0.1 (Crepis-Downloadmanager; www.melaxis.com)"
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Zentriert
            Caption         =   "="
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "Cookie:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "User-Agent:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox tLOCAL 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "c:\downloads\"
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox tURL 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "http://"
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "Speichern nach:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   0
      MouseIcon       =   "main.frx":1272
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "main.frx":13C4
      ToolTipText     =   "Crepis-Downloadmanager"
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim server As String, port As Integer, dateipfad As String
Dim Filesize As Long
Dim allbytes As Long

Dim PARTSTART(1 To 4) As Long
Dim PARTSTOP(1 To 4) As Long
Dim PARTSIZE As Long

Dim BytePos(1 To 4) As Long
Dim wHEADER(1 To 4) As Boolean
Dim lFILES(1 To 4) As String
Dim FILENUMS(1 To 4) As Long

Dim StartTime As Long
Dim tRATE As Single

Dim Fertig As Boolean

Private Sub cSTARTDOWNLOAD_Click()
If tURL = "http://" Then Exit Sub
Me.Caption = "Crepis-Downloadmanager"
lURL = tURL
lSPEED = "??"
lSIZE = "0 Byte von ???"
lNOCH = "???"
lRESUME = "???"
Fertig = False
barKOMPLETT.Value = 0
For i = 1 To 4
    st(i) = "???"
    lBYTES(i) = "0 B"
    barSTREAM(i).Value = 0
Next
For i = 0 To 4
    wDL(i).Close
Next
st(0) = ""
port = 80
server = cutStart(tURL, "http://")
d = Split(server, "/")
server = d(0)
dateipfad = cutStart(tURL, "http://" & server)
fDL.Visible = True
DoEvents
st(0) = "Dateiinfos abfragen..."
DoEvents
On Error GoTo conerr
wDL(0).Connect server, port
Exit Sub
conerr:
st(0) = "Verbindungsfehler."
On Error GoTo 0
End Sub

Private Sub cSTOP_Click()
st(0) = "Abbrechen..."
wDL(0).Close
DoEvents
For i = 1 To 4
    wDL(i).Close
    DoEvents
    st(i) = "Abgebrochen."
Next
st(0) = "Abgebrochen."
DoEvents
fDL.Visible = False
End Sub

Private Sub Form_Load()
Dim clip As String
If Command <> "" Then
    tURL = Command
    Exit Sub
End If
clip = Clipboard.GetText
If isStart(clip, "http://") Then
    tURL = clip
Else
    tURL = "http://www.melaxis.de/paradiseclient/paradise-setup.exe"
End If
End Sub

Private Sub Image1_Click()
StartMe Me, "http://www.melaxis.com/crepis/"
End Sub

Private Sub tURL_Change()
Dim Datei As String
d = Split(tURL, "/")
Datei = d(UBound(d))
If InStr(tLOCAL, "\") > 0 Then
    If Right(tLOCAL, 1) = "\" Then
        tLOCAL = tLOCAL & Datei
    Else
        d = Split(tLOCAL, "\")
        tLOCAL = Mid(tLOCAL, 1, Len(tLOCAL) - Len(d(UBound(d)))) & Datei
    End If
End If
End Sub

Private Sub wDL_Close(Index As Integer)
Dim mpfn As Long
If Index = 0 Then
    st(0) = "Download beginnen..."
    wDL(0).Close
    DoEvents
    PARTSIZE = Int(Filesize / 4)
    allbytes = 0
    StartTime = 0
    For i = 1 To 4
        st(i) = "Verbinden..."
        PARTSTART(i) = PARTSIZE * (i - 1)
        PARTSTOP(i) = (PARTSIZE * i) - 1
        wHEADER(i) = True
        lFILES(i) = App.Path & "\" & Round(Rnd * Timer) & ".tmp"
        wDL(i).Close
        wDL(i).Connect server, port
        DoEvents
    Next
    mpfn = FreeFile
    Open tLOCAL & ".crr" For Output As #mpfn
    Print #mpfn, tURL
    Print #mpfn, tLOCAL
    Print #mpfn, Filesize
    Print #mpfn, lFILES(1)
    Print #mpfn, lFILES(2)
    Print #mpfn, lFILES(3)
    Print #mpfn, lFILES(4)
    Close #mpfn
Else
    st(Index) = "Abgeschlossen."
    wDL(Index).Close
    DoEvents
    If Fertig Then Exit Sub
    For i = 1 To 4
        If wDL(i).State <> 0 Then Exit Sub
    Next
    st(0) = "Ausgabe erstellen..."
    lNOCH = "Fertig"
    Fertig = True
    DoEvents
    Dim myFN As Long, bytA() As Byte, bytP As Long
    myFN = FreeFile
    bytep = 1
    Open tLOCAL For Binary As #myFN
    For i = 1 To 4
        FILENUMS(i) = FreeFile
        Open lFILES(i) For Binary As #FILENUMS(i)
        ReDim bytA(0 To FileLen(lFILES(i)) - 1)
        Get #FILENUMS(i), , bytA
        Close #FILENUMS(i)
        Put #myFN, bytep, bytA
        bytep = bytep + FileLen(lFILES(i))
    Next
    Close #myFN
    DoEvents
    st(0) = "Download abgeschlossen."
    For i = 1 To 4
        Kill lFILES(i)
    Next
    Kill tLOCAL & ".crr"
    Me.Caption = "Fertig - Crepis-Downloadmanager"
    DoEvents
    fDL.Visible = False
End If
End Sub

Private Sub wDL_Connect(Index As Integer)
Fertig = False
If Index = 0 Then
    send "HEAD " & dateipfad & " HTTP/1.1", 0
    send "Host: " & server, 0
    send "User-Agent: " & tBROWSER, 0
    If tCOOKIENAME <> "" Then
        send "Cookie: " & urlencode(tCOOKIENAME) & "=" & urlencode(tCOOKIEWERT), 0
    End If
    send "Accept: */*, *,*", 0
    send "Connection: close", 0
    send "", 0
Else
    st(Index) = "Verbunden."
    wHEADER(Index) = True
    BytePos(Index) = 1
    send "GET " & dateipfad & " HTTP/1.1", Index
    send "Host: " & server, Index
    send "User-Agent: " & tBROWSER, Index
    If tCOOKIENAME <> "" Then
        send "Cookie: " & urlencode(tCOOKIENAME) & "=" & urlencode(tCOOKIEWERT), Index
    End If
    send "Accept: */*, *,*", Index
    send "Connection: close", Index
    send "Range: bytes=" & PARTSTART(Index) & "-" & IIf(Index <> 4, PARTSTOP(Index), ""), Index
    send "", Index
End If
End Sub

Private Sub wDL_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim s As String, z As String, lFILE As String, remme As String
If wDL(Index).State <> 7 Then Exit Sub
wDL(Index).GetData s
If Index = 0 Then
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, vbCrLf)
    d = Split(s, vbCrLf)
    st(0) = "Empfange Dateiinfos..."
    For i = 0 To UBound(d)
        z = Trim(d(i))
        If z <> "" Then
            If isStart(z, "HTTP/1.1 ") Then
                z = Trim(cutStart(z, "HTTP/1.1"))
                z = Mid(z, 1, 3)
                If z <> "200" Then
                    If z = "404" Then
                        st(0) = "Datei nicht gefunden."
                        MsgBox "Die Datei konnte nicht gefunden werden.", vbCritical, "Download"
                    End If
                    If z = "501" Then
                        st(0) = "Server inkompatibel."
                        MsgBox "Der Server ist nicht kompatibel.", vbCritical, "Download"
                    End If
                    If z = "403" Or z = "401" Then
                        st(0) = "Zugriff verweigert."
                        MsgBox "Der Zugriff auf die Datei wurde verweigert.", vbCritical, "Download"
                    End If
                    If z = "500" Then
                        st(0) = "Serverfehler."
                        MsgBox "Es ist ein interner Serverfehler aufgetreten.", vbCritical, "Download"
                    End If
                End If
                GoTo s0nh
            End If
            If isStart(z, "Content-Length:") Then
                z = Trim(cutStart(z, "Content-Length:"))
                Filesize = CLng(z)
                lSIZE = "0 Bytes von " & Filesize & " Bytes"
                If StartTime = 0 Then StartTime = Timer
                barKOMPLETT.Max = Filesize
                GoTo s0nh
            End If
        End If
s0nh:
    Next i
Else




If wHEADER(Index) Then
    '** Zeilenumbrüche auf vbCrLf bringen
    Debug.Print "** Header empfangen"
'    s = Replace(s, vbCrLf, vbLf)
'    s = Replace(s, vbCr, vbLf)
'    s = Replace(s, vbLf, vbCrLf)
    '** In einzelne Zeilen aufteilen
    d = Split(s, vbCrLf)
    remme = ""
    st(Index) = "Empfange Header..."
    For i = 0 To UBound(d)
        '** Alle Zeilen durchlaufen
        remme = remme & d(i) & vbCrLf
        If d(i) <> "" Then
            '** Uns interessiert nur die Länge der Datei
            Debug.Print "** Header: " & d(i)
            If isStart(CStr(d(i)), "HTTP/1.1 ") Then
                z = Trim(cutStart(CStr(d(i)), "HTTP/1.1 "))
                z = Mid(z, 1, 3)
                If z <> "206" Then
                    lRESUME = "Nein"
                Else
                    lRESUME = "Ja"
                End If
            End If
            If isStart(CStr(d(i)), "Content-Length:") Then
                z = Trim(cutStart(CStr(d(i)), "Content-Length:"))
                barSTREAM(Index).Max = z
                barSTREAM(Index).Value = 0
                Debug.Print "** Dateilänge: " & z
            End If
        Else
            '** Wenn Leerzeile, ist Header zu ende
            Debug.Print "** Ende des Headers"
            wHEADER(Index) = False
            s = cutStart(s, remme)
            st(Index) = "Empfange Datei..."
            If StartTime = 0 Then StartTime = Timer
            GoTo binfile
            Exit For
        End If
    Next
Else
binfile:
    '** Teil der Datei, binär!
    '** Datei zum Schreiben öffnen
    lFILE = lFILES(Index)
    FILENUMS(Index) = FreeFile
    Debug.Print FILENUMS(Index) & " opened for " & Index
    Open lFILE For Binary Access Write As #FILENUMS(Index)
    '** An Position schreiben
    Put #FILENUMS(Index), BytePos(Index), s
    '** Position erneuern, verbleibende Bytes berechnen
    BytePos(Index) = Seek(FILENUMS(Index))
    On Error Resume Next
    barSTREAM(Index).Value = BytePos(Index) - 1
    On Error GoTo 0
    '** Datei wieder schliessen
    Close #FILENUMS(Index)
'    '** Prozentzahl ausrechnen und runden
'    lPRO = Round((BytePos / ByteLen) * 100) & "%"
'    '** Balkenbreite ausrechnen
'    Shape2.Width = Round((BytePos / ByteLen) * Shape1.Width)
'    '** Transferrate ausrechnen
'    tRATE = Format(Int(BytePos / (Timer - StartTime)) / 1000, "####.00")
'    Debug.Print "** Binärdaten erhalten, Position: " & BytePos & ", Rate: " & tRATE & " KB/s"
    lBYTES(Index) = Round((BytePos(Index) - 1) / 1024, 1) & " KB"
    allbytes = BytePos(1) + BytePos(2) + BytePos(3) + BytePos(4) - 4
    lSIZE = Round(allbytes / 1024, 1) & " KB von " & Round(Filesize / 1024, 1) & " KB"
    On Error Resume Next
    barSTREAM(Index).Value = BytePos(Index) - 1
    barKOMPLETT.Value = allbytes
    tRATE = Format(Int(allbytes / (Timer - StartTime)) / 1000, "####.00")
    lSPEED = tRATE & " KB/s"
    'lNOCH = ConvertTime(Int((((Filesize - allbytes) - allbytes) / 1024) / tRATE))
    lNOCH = ConvertTime(Int((((Filesize - allbytes)) / 1024) / tRATE))
    Me.Caption = Round((allbytes / Filesize) * 100) & "%" & " - Crepis-Downloadmanager"
    On Error GoTo 0
End If






End If
End Sub

Private Sub wDL_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wDL(Index).Close
st(Index) = "Verbindungsfehler."
End Sub

Sub send(DATA As String, Index As Integer)
On Error Resume Next
wDL(Index).SendData DATA & vbCrLf
DoEvents
On Error GoTo 0
End Sub


Function ConvertTime(TheTime As Single)
    Dim NewTime As String
    Dim Sec As Single
    Dim Min As Single
    Dim H As Single

    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If


    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function
