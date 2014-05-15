VERSION 5.00
Begin VB.Form frmVB6serialAPItest 
   Caption         =   "Test for class for one serial API port"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameLines 
      Caption         =   "Control Lines"
      Height          =   1095
      Left            =   1800
      TabIndex        =   26
      Top             =   3750
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton cmdRTS0 
         Caption         =   "RTS0"
         Height          =   345
         Left            =   1560
         TabIndex        =   32
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdDTR0 
         Caption         =   "DTR0"
         Height          =   345
         Left            =   840
         TabIndex        =   31
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdBRK0 
         Caption         =   "BRK0"
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdRTS1 
         Caption         =   "RTS1"
         Height          =   345
         Left            =   1560
         TabIndex        =   29
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdDTR1 
         Caption         =   "DTR1"
         Height          =   345
         Left            =   840
         TabIndex        =   28
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdBRK1 
         Caption         =   "BRK1"
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   660
         Width           =   615
      End
   End
   Begin VB.Frame FrameErrors 
      Caption         =   "ERRORS"
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   4950
      Width           =   8415
      Begin VB.TextBox txbErrors 
         Height          =   525
         Left            =   120
         TabIndex        =   25
         Top             =   330
         Width           =   8175
      End
      Begin VB.CommandButton butClearErrorReport 
         Caption         =   "Clear Error box"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox tbxWrite 
      Height          =   405
      Left            =   1800
      TabIndex        =   21
      Text            =   "Hello serial API sender"
      Top             =   1920
      Width           =   6615
   End
   Begin VB.TextBox txbVersion 
      Height          =   405
      Left            =   5400
      TabIndex        =   14
      Top             =   6000
      Width           =   3015
   End
   Begin VB.TextBox txbStatus 
      Height          =   405
      Left            =   1800
      TabIndex        =   13
      Text            =   "NOT SET UP (1)"
      Top             =   3210
      Width           =   6615
   End
   Begin VB.TextBox txbRead 
      Height          =   525
      Left            =   1800
      TabIndex        =   12
      Top             =   2550
      Width           =   6615
   End
   Begin VB.TextBox txbBaud 
      Height          =   345
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   810
      Width           =   975
   End
   Begin VB.TextBox txbPortNumber 
      Height          =   345
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   210
      Width           =   975
   End
   Begin VB.CommandButton butSetUpBaud 
      Caption         =   "Setup Baud"
      Height          =   490
      Left            =   120
      TabIndex        =   7
      Top             =   690
      Width           =   1220
   End
   Begin VB.CommandButton butInit 
      Caption         =   "Init"
      Height          =   490
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1220
   End
   Begin VB.CommandButton cmdFlush 
      Caption         =   "Flush"
      Height          =   490
      Left            =   120
      TabIndex        =   5
      Top             =   3780
      Width           =   1220
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Show Status"
      Height          =   490
      Left            =   120
      TabIndex        =   4
      Top             =   3150
      Width           =   1220
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   490
      Left            =   120
      TabIndex        =   3
      Top             =   4410
      Width           =   1220
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   490
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1220
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   490
      Left            =   120
      TabIndex        =   1
      Top             =   1860
      Width           =   1220
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   490
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   1220
   End
   Begin VB.Label labToClose 
      Caption         =   "To Close: Press close button first then X above."
      Height          =   495
      Left            =   6240
      TabIndex        =   34
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label labCrLf 
      Caption         =   "Put \r for <CR> and \n for <LF>"
      Height          =   315
      Left            =   6000
      TabIndex        =   33
      Top             =   1620
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "<<"
      Height          =   225
      Left            =   1440
      TabIndex        =   22
      Top             =   2010
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   ">>"
      Height          =   225
      Left            =   1440
      TabIndex        =   20
      Top             =   3270
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   ">>"
      Height          =   225
      Left            =   1440
      TabIndex        =   19
      Top             =   2670
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "<<"
      Height          =   225
      Left            =   1440
      TabIndex        =   18
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "<<"
      Height          =   225
      Left            =   1440
      TabIndex        =   17
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "clsVB6serialAPI version"
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   6060
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "frmVB6serialAPItest"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   6030
      Width           =   1935
   End
   Begin VB.Label labBaud 
      Caption         =   "Baud eg 2400 or 9600 etc"
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   810
      Width           =   1935
   End
   Begin VB.Label labPortNumber 
      Caption         =   "Port Number eg 1 or 18"
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   210
      Width           =   1935
   End
End
Attribute VB_Name = "frmVB6serialAPItest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program language used: Microsoft Visual Basic Six (VB6).
'This form provides a test for serial API port handling class clsVB6serialAPI
'and also demonstrates use of the public functions of the class.
'Class version control is at the top of the class code.

'Port 4 is default on startup as laptops without serial ports can use a USB to
'serial adapter that often is allocated to this port.

'CAUTION: When using the VB6 development system always close an open port then
' close the form (click top RHS X on form).
' Selecting Run/End when a port is open leads to failure to open the port next time.
' If this is experienced exit the development system (File/Exit) and restart.

'NB In this test code only one serial port is set up see Note 1
'however if several are needed it may be convenient to address them as an array
'as the following example shows...
'Public clSerAPI(PORT_FIRST To PORT_LAST) As clsVB6serialAPI
'Dim iPortNum As Integer
'For iPortNum = PORT_FIRST To PORT_LAST
'  'Go through every serial port that could be used
'  Set clSerAPI(iPortNum) = New clsVB6serialAPI 'Make a new instance
'  clSerAPI(iPortNum).clSerialAPIinitialise iPortNum 'Initialise the port
'Next iPortNum

'Modification record (latest float on top)
'CX003 Add line handling 14May2014
'CX002 clSerialAPIflush added to return string 13May2014
'CX001 add test if class object exists 07May2014

Private clSerAPI As clsVB6serialAPI

Private Sub Form_Load()
  txbStatus.Text = "Port not set up"
  'Put in suggested starting port number and baud
  txbPortNumber.Text = "4"
  txbBaud.Text = "9600"
End Sub

Private Sub butInit_Click()
  'Init port allocating port number but baud etc not yet set up
  Dim sErrorReturn, sPortNumber As String
  Dim iSerialPortToTest As Integer
  
  sPortNumber = txbPortNumber.Text
  If IsNumeric(sPortNumber) = False Then
    txbErrors.Text = "Type in port number to use as a number"
    Exit Sub
  End If
  iSerialPortToTest = Int(sPortNumber)
  
  If Not clSerAPI Is Nothing Then
    'ie if class is something then... ie class was previously instantiated
    txbErrors.Text = "Serial API port number " & sPortNumber & " is already in use"
    Exit Sub
  End If
  
  Set clSerAPI = New clsVB6serialAPI 'make a new instance Note 1
  
  'Check class was instantiated
  If clSerAPI Is Nothing Then
    'Class was not instantiated
    txbErrors.Text = "Error: Class was not set up"
    Exit Sub
  End If
  
  sErrorReturn = clSerAPI.clSerialAPIinitialise(iSerialPortToTest) 'Add serial port class initialisation
  If sErrorReturn <> "" Then
    txbErrors.Text = sErrorReturn
  Else
    'All set up ok
    frameLines.Visible = True 'allow line states to be changed
  End If
  cmdStatus_Click 'show status
  txbVersion.Text = clSerAPI.clSerialAPIgetVersion
End Sub

Private Sub butSetUpBaud_Click()
  'Set baud and port description
  Dim sRetn, sBaud As String
  sRetn = ""
  
  'CX001 add test if class object exists
  If clSerAPI Is Nothing Then
    'API class needs setting up
    txbErrors.Text = "Initialise before setting Baud"
    Exit Sub
  End If
  
  sBaud = txbBaud.Text
  If IsNumeric(sBaud) = False Then
    txbErrors.Text = "Type in baud to use"
    Exit Sub
  End If
  
  'Here class is something then... ie class was instantiated
  sRetn = clSerAPI.clSerialAPIsetBaud(sBaud)
  If sRetn <> "" Then
    txbErrors.Text = sRetn
  End If
  cmdStatus_Click 'show status
End Sub

Private Sub cmdStatus_Click()
  'Test open close function and show status of port
  Dim iBaud As Integer
  Dim bIsOpen As Boolean
  Dim sOpenClosd, sPortNumber, sLineStates As String
  
  If clSerAPI Is Nothing Then
    'ie if class was not instantiated
    txbErrors.Text = "Initialise before reading status"
    Exit Sub
  End If
  
  'Get port number in use
  sPortNumber = txbPortNumber.Text
  If IsNumeric(sPortNumber) = False Then
    txbErrors.Text = "Type in port number to use"
    Exit Sub
  End If

  iBaud = clSerAPI.clSerialAPIgetBaud
  bIsOpen = clSerAPI.clSerialAPIgetIsOpen
  If bIsOpen = True Then
    sLineStates = clSerAPI.clSerialAPIgetLines
    sOpenClosd = "OPEN (" & sLineStates & ")"
  Else
    sOpenClosd = "CLOSED"
  End If
  txbStatus.Text = Format(sPortNumber, "00") & " " & iBaud & " baud " & " " & sOpenClosd
End Sub

Private Sub cmdOpen_Click()
  Dim sErrorReturn As String
  
  sErrorReturn = ""
  'CX001 add test if class object exists
  If clSerAPI Is Nothing Then
    'API class needs setting up
    txbErrors.Text = "Initialise before opening the port"
    Exit Sub
  End If
  'Here class is something then... ie class was instantiated
  sErrorReturn = clSerAPI.clSerialAPIopen()
  If sErrorReturn <> "" Then
    txbErrors.Text = sErrorReturn
  End If
  cmdStatus_Click 'show status
End Sub

Private Sub cmdRead_Click()
  Dim sStringReadIn As String
  Dim sErrorReturn As String
  
  sErrorReturn = ""
  'CX001 add test if class object exists
  If Not clSerAPI Is Nothing Then
   'ie if class is something then... ie class was instantiated
    sStringReadIn = clSerAPI.clSerialAPIread(64, sErrorReturn)
  Else
    'API class needs setting up
    sErrorReturn = "Initialise before Read"
  End If
  If sErrorReturn <> "" Then
    'Problem seen
    txbErrors.Text = sErrorReturn
    txbRead.Text = ""
  Else
    'No problems returned
    txbRead.Text = sStringReadIn
  End If
End Sub

Private Sub cmdWrite_Click()
  ' Write data to serial port.
  Dim sErrorReturn, sTextToWrite As String
  sErrorReturn = ""
  sTextToWrite = tbxWrite.Text
  sTextToWrite = Replace(sTextToWrite, "\r", vbCr)
  sTextToWrite = Replace(sTextToWrite, "\n", vbLf)
  'CX001 add test if class object exists
  If Not clSerAPI Is Nothing Then
    'ie if class is something then... ie class was instantiated
    sErrorReturn = clSerAPI.clSerialAPIwrite(sTextToWrite) 'Returns blank string if okay else message
  Else
    'API class needs setting up
    sErrorReturn = "Initialise before Write"
  End If
  If sErrorReturn <> "" Then
    'Problem seen
    txbErrors.Text = sErrorReturn
  End If
End Sub

Private Sub cmdClose_Click()
  'This just closes the port, no shutdown
  clSerAPI.clSerialAPIclose
  cmdStatus_Click 'show status
End Sub

'Test Flush
Private Sub cmdFlush_Click()
  'CX001 add test if class object exists
  Dim sRetn As String
  If clSerAPI Is Nothing Then
    'API class needs setting up
    txbErrors.Text = "Initialise before flushing"
    Exit Sub
  End If
  sRetn = clSerAPI.clSerialAPIflush
  If sRetn <> "" Then
    'Problem ocurred
    txbErrors.Text = sRetn
  End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
  'Closes comports and forms and shuts down.
  Dim sStr As String
  If Not clSerAPI Is Nothing Then
    'ie if class is something then... ie class was instantiated
    sStr = clSerAPI.clSerialAPIclose
    showInfo "*** CLOSING... *** " & sStr
  End If
  Close   ' Close all open files.

  'If code below is not executed the application won't close properly
  Dim frmCurrent As Form
  For Each frmCurrent In Forms
    If Not "frmSAPIMain" = frmCurrent.Name Then
      Unload frmCurrent
      Set frmCurrent = Nothing
    End If
  Next
End Sub

Private Sub butClearErrorReport_Click()
  'Clear Error box
  txbErrors.Text = ""
End Sub

Sub showInfo(ByVal sTeext As String)
  'Show information in pop up box
  MsgBox sTeext, vbInformation
End Sub

'Set BREAK DTR and RTS lines True(1) or False(0) - six routines
Private Sub cmdBRK0_Click()
  'Test below is not essential as frame with line buttons in is not visible on start up.
  If clSerAPI Is Nothing Then
    'Class was not instantiated
    txbErrors.Text = "Error: Class not yet set up"
    Exit Sub
  End If
  txbErrors.Text = clSerAPI.clSerialAPIsetBREAK(False)
  cmdStatus_Click 'show status
End Sub

Private Sub cmdBRK1_Click()
  txbErrors.Text = clSerAPI.clSerialAPIsetBREAK(True)
  cmdStatus_Click 'show status
End Sub

Private Sub cmdDTR0_Click()
  txbErrors.Text = clSerAPI.clSerialAPIsetDTR(False)
  cmdStatus_Click 'show status
End Sub

Private Sub cmdDTR1_Click()
  txbErrors.Text = clSerAPI.clSerialAPIsetDTR(True)
  cmdStatus_Click 'show status
End Sub

Private Sub cmdRTS0_Click()
  txbErrors.Text = clSerAPI.clSerialAPIsetRTS(False)
  cmdStatus_Click 'show status
End Sub

Private Sub cmdRTS1_Click()
  txbErrors.Text = clSerAPI.clSerialAPIsetRTS(True)
  cmdStatus_Click 'show status
End Sub
'--------From end

