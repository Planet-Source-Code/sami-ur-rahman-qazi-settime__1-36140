VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSetTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Synchronize PC Clock"
   ClientHeight    =   3240
   ClientLeft      =   3210
   ClientTop       =   2280
   ClientWidth     =   5040
   Icon            =   "SetTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5040
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   2760
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "cboServers"
      ToolTipText     =   "Remote server to get time from. "
      Top             =   240
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   2760
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2685
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "&Adjust"
      Height          =   375
      Left            =   1605
      TabIndex        =   3
      ToolTipText     =   "Synchronize PC clock."
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Go get remote time."
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timings"
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      Begin VB.TextBox txtDifference 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "txtDifference"
         ToolTipText     =   "Difference between Local time & Server time."
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtLocalTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "txtLocalTime"
         ToolTipText     =   "Local time."
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtServerTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "txtServerTime"
         ToolTipText     =   "Server time in UTC/GMT"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Difference:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Local Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Server Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      Caption         =   "&Servers:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const TIME_ZONE_ID_UNKNOWN = 0
Const TIME_ZONE_ID_STANDARD = 1
Const TIME_ZONE_ID_DAYLIGHT = 2

Private RemoteTime As String        'the 32bit time stamp returned by the server
Private UTCTime As Date
Private TimeDelay As Single         'the time between the acknowledgement of the
                                    'connection and the data received. we compensate
                                    'by adding half of the round trip latency.
Private ZoneFactor As Long          'Adding this to UTC time will give us loacal time

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function ntohl Lib "WSOCK32.DLL" (ByVal hostlong As Long) As Long

Private Sub cmdCheck_Click()
  'clear the string used for incoming data
   RemoteTime = Empty
   
   cmdCheck.Enabled = False
   Timer2.Enabled = False
   txtServerTime.Text = ""
   txtDifference.Text = "Calculating time difference..."
   ZoneFactor = 60 * AdjustTimeForTimeZone
   
  'connect
   With Socket
      If .State <> sckClosed Then .Close
      .RemoteHost = cboServers.Text
      .RemotePort = 37  'port 37 is the timserver port
      .Connect
   End With
End Sub

Private Sub cmdAdjust_Click()

   Dim ST As SYSTEMTIME
   
     Timer2.Enabled = False
     
     'fill a SYSTEMTIME structure with the appropriate values
      With ST
         .wYear = Year(UTCTime)
         .wMonth = Month(UTCTime)
         .wDay = Day(UTCTime)
         .wHour = Hour(UTCTime)
         .wMinute = Minute(UTCTime)
         .wSecond = Second(UTCTime)
      End With
     Timer2.Enabled = True
     'and call the API with the new date & time
      If SetSystemTime(ST) Then
         txtDifference.Text = "PC Clock synchronised"
         cmdAdjust.Enabled = False
      Else
         txtDifference.Text = "Pc Clock not synchronised!"
      End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    With cboServers
        .AddItem "time-a.timefreq.bldrdoc.gov"
        .AddItem "time-b.timefreq.bldrdoc.gov"
        .AddItem "time-c.timefreq.bldrdoc.gov"
        .AddItem "utcnist.colorado.edu"
        .AddItem "time-nw.nist.gov"
        .AddItem "nist1.nyc.certifiedtime.com"
        .AddItem "nist1.dc.certifiedtime.com"
        .AddItem "nist1.sjc.certifiedtime.com"
        .AddItem "nist1.datum.com"
        .AddItem "ntp2.cmc.ec.gc.ca"
        .AddItem "ntps1-0.uni-erlangen.de"
        .AddItem "ntps1-1.uni-erlangen.de"
        .AddItem "ntps1-2.uni-erlangen.de"
        .AddItem "ntps1-0.cs.tu-berlin.de"
        .AddItem "time.ien.it"
        .AddItem "ptbtime1.ptb.de"
        .AddItem "ptbtime2.ptb.de"
        .ListIndex = 0
    End With
    txtServerTime.Text = ""
    txtLocalTime.Text = Format(Now, "hh:mm:ss") + "  " + Format(Now, "ddd mmm d, yyyy")
    txtDifference.Text = ""
    Timer1.Enabled = True
    cmdAdjust.Enabled = False
End Sub

Private Sub Form_Load()
    frmSetTime.Show
    cmdCheck.SetFocus
End Sub

Private Sub Socket_Close()
    Dim NTPTime As Double
    Dim LocalTime As Date
    Dim dwSecondsSince1990 As Long
    Dim Difference As Long
    
    Socket.Close
   
    RemoteTime = Trim(RemoteTime)
    If Len(RemoteTime) = 4 Then
        'since the data was returned in a string,
        'format it back into a numeric value
        NTPTime = Asc(Left$(RemoteTime, 1)) * (256 ^ 3) + _
                  Asc(Mid$(RemoteTime, 2, 1)) * (256 ^ 2) + _
                  Asc(Mid$(RemoteTime, 3, 1)) * (256 ^ 1) + _
                  Asc(Right$(RemoteTime, 1))
        
        'calculate round trip delay
        TimeDelay = (Timer - TimeDelay)
'        Debug.Print TimeDelay
        'and create a valid date based on
        'the seconds since January 1, 1990
        dwSecondsSince1990 = NTPTime - 2840140800# + CDbl(TimeDelay)
        UTCTime = DateAdd("s", CDbl(dwSecondsSince1990), #1/1/1990#)
        
        'convert UTC time to local time and get the difference
        LocalTime = DateAdd("s", CDbl(ZoneFactor), UTCTime)
        Difference = DateDiff("s", Now, LocalTime)
        
        Timer2.Enabled = True
        cmdAdjust.Enabled = True
        
        If Difference < 0 Then
            txtDifference.Text = "PC Clock " + CStr(-Difference) + " sec ahead."
        ElseIf Difference > 0 Then
            txtDifference.Text = "PC Clock " + CStr(Difference) + " sec behind."
        Else
            txtDifference.Text = "PC & Server Clock matched."
        End If
    Else
      txtServerTime.Text = "Time received not valid."
      Timer2.Enabled = False
      cmdAdjust.Enabled = False
    End If
    cmdCheck.Enabled = True
End Sub

Private Sub Socket_Connect()
    TimeDelay = Timer
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
   Dim sData As String
   Socket.GetData sData, vbString
   RemoteTime = RemoteTime & sData
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txtServerTime.Text = Description
    cmdCheck.Enabled = True
End Sub

Private Sub Timer1_Timer()
    txtLocalTime.Text = Format(Now, "hh:mm:ss") + "  " + Format(Now, "ddd mmm d, yyyy")
End Sub

Private Sub Timer2_Timer()
    txtServerTime.Text = Format(UTCTime, "hh:mm:ss") + "  " + Format(UTCTime, "ddd mmm d, yyyy")
    UTCTime = DateAdd("s", CDbl(1), UTCTime)
End Sub
'Returns the amount of adjustment in seconds necessary from UTC time for the
'current system by checking the system's time zone and daylight savings properties
Private Function AdjustTimeForTimeZone() As Long

    Dim TZI As TIME_ZONE_INFORMATION
    Dim RetVal As Long
    Dim ZoneCorrection As Long
    
    RetVal = GetTimeZoneInformation(TZI)
    ZoneCorrection = TZI.Bias
    If RetVal = TIME_ZONE_ID_STANDARD Then
        ZoneCorrection = ZoneCorrection + TZI.StandardBias
    ElseIf RetVal = TIME_ZONE_ID_DAYLIGHT Then
            ZoneCorrection = ZoneCorrection + TZI.DaylightBias
    Else
        MsgBox "Unable to get zone information.", vbExclamation, "Error"
    End If
    AdjustTimeForTimeZone = -ZoneCorrection     'correction in minutes
End Function

