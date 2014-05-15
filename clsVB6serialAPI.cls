VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVB6serialAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Force explicit variable declaration.

'********************************************************************************
'This class represents a serial port with API comms
'Language is Microsoft Visual Basic Six (VB6).
'See...
'  http://stackoverflow.com/questions/4929414/serial-port-programming-vb6-via-win32-api
'Serial Port Communication: perform serial port I/O without using the Microsoft Comm Control component
'Above link leads to...
'  http://www.thescarms.com/vbasic/CommIO.aspx
'
' This class is based on content from original CommIO.bas modCOMM module
' from the above reference.
' This is a collection of routines to perform serial port I/O without
' using the Microsoft Comm Control component.  This module uses the Windows API
' to perform the overlapped I/O operations necessary for serial communications.
' See API serial port public functions below for possible options.
'********************************************************************************
Const scVersion = "20140515_2040" 'Version of this class

Private z_SerialPortNumber As Integer 'Unique key
Private z_iBaud As Integer 'Also 0 means port is not set up
Private z_bIsOpen As Boolean 'True if serial port was opened

'-------------------------------------------------------------------------------
' Constants
'-------------------------------------------------------------------------------

' Output Control Lines (CommSetLine)
Private Const LINE_BREAK = 1
Private Const LINE_DTR = 2
Private Const LINE_RTS = 3

' Input Control Lines  (CommGetLine)
Private Const LINE_CTS = &H10&
Private Const LINE_DSR = &H20&
Private Const LINE_RING = &H40&
Private Const LINE_RLSD = &H80&
Private Const LINE_CD = &H80&

'-------------------------------------------------------------------------------
' System Constants
'-------------------------------------------------------------------------------
Private Const ERROR_IO_INCOMPLETE = 996&
Private Const ERROR_IO_PENDING = 997
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const OPEN_EXISTING = 3

' COMM Functions
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&
Private Const PURGE_RXABORT = &H2
Private Const PURGE_RXCLEAR = &H8
Private Const PURGE_TXABORT = &H1
Private Const PURGE_TXCLEAR = &H4

' COMM Escape Functions
Private Const CLRBREAK = 9
Private Const CLRDTR = 6
Private Const CLRRTS = 4
Private Const SETBREAK = 8
Private Const SETDTR = 5
Private Const SETRTS = 3

'-------------------------------------------------------------------------------
' System Structures
'-------------------------------------------------------------------------------
Private Type COMSTAT
        fBitFields As Long ' See Comment in Win32API.Txt
        cbInQue As Long
        cbOutQue As Long
End Type

Private Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type

'
' The DCB structure defines the control setting for a serial
' communications device.
'
Private Type DCB
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long ' See Comments in Win32API.Txt
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer 'Reserved; Do Not Use
End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


'-------------------------------------------------------------------------------
' Program Structures
'-------------------------------------------------------------------------------

Private Type COMM_ERROR
    lngErrorCode As Long
    strFunction As String
    strErrorMessage As String
End Type

Private Type COMM_PORT
    lngHandle As Long
    blnPortOpen As Boolean
    udtDCB As DCB
End Type

'-------------------------------------------------------------------------------
' Program Storage
'-------------------------------------------------------------------------------

Private udtCommOverlap As OVERLAPPED
Private udtCommError As COMM_ERROR
'was Private udtPorts(1 To MAX_API_PORTS) As COMM_PORT
Private udtPort1 As COMM_PORT

'-------------------------------------------------------------------------------
' System Functions begin
'-------------------------------------------------------------------------------
'
' Fills a specified DCB structure with values specified in
' a device-control string.
'http://msdn.microsoft.com/en-us/library/windows/desktop/aa363143%28v=vs.85%29.aspx
'
Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" _
    (ByVal lpDef As String, lpDCB As DCB) As Long
'
' Retrieves information about a communications error and reports
' the current status of a communications device. The function is
' called when a communications error occurs, and it clears the
' device's error flag to enable additional input and output
' (I/O) operations.
Private Declare Function ClearCommError Lib "kernel32" _
    (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
'
' Closes an open communications device or file handle.
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
' Creates or opens a communications resource and returns a handle
' that can be used to access the resource.
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
'
' Directs a specified communications device to perform a function.
Private Declare Function EscapeCommFunction Lib "kernel32" _
    (ByVal nCid As Long, ByVal nFunc As Long) As Long
'
' Formats a message string such as an error string returned
' by anoher function.
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
'
' Retrieves modem control-register values.
Private Declare Function GetCommModemStatus Lib "kernel32" _
    (ByVal hFile As Long, lpModemStat As Long) As Long
'
' Retrieves the current control settings for a specified
' communications device.
Private Declare Function GetCommState Lib "kernel32" _
    (ByVal nCid As Long, lpDCB As DCB) As Long
'
' Retrieves the calling thread's last-error code value.
Private Declare Function GetLastError Lib "kernel32" () As Long
'
' Retrieves the results of an overlapped operation on the
' specified file, named pipe, or communications device.
Private Declare Function GetOverlappedResult Lib "kernel32" _
    (ByVal hFile As Long, lpOverlapped As OVERLAPPED, _
    lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
'
' Discards all characters from the output or input buffer of a
' specified communications resource. It can also terminate
' pending read or write operations on the resource.
Private Declare Function PurgeComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwFlags As Long) As Long
'
' Reads data from a file, starting at the position indicated by the
' file pointer. After the read operation has been completed, the
' file pointer is adjusted by the number of bytes actually read,
' unless the file handle is created with the overlapped attribute.
' If the file handle is created for overlapped input and output
' (I/O), the application must adjust the position of the file pointer
' after the read operation.
Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    lpOverlapped As OVERLAPPED) As Long
'
' Configures a communications device according to the specifications
' in a device-control block (a DCB structure). The function
' reinitializes all hardware and control settings, but it does not
' empty output or input queues.
Private Declare Function SetCommState Lib "kernel32" _
    (ByVal hCommDev As Long, lpDCB As DCB) As Long
'
' Sets the time-out parameters for all read and write operations on a
' specified communications device.
Private Declare Function SetCommTimeouts Lib "kernel32" _
    (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
'
' Initializes the communications parameters for a specified
' communications device.
Private Declare Function SetupComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
'
' Writes data to a file and is designed for both synchronous and a
' synchronous operation. The function starts writing data to the file
' at the position indicated by the file pointer. After the write
' operation has been completed, the file pointer is adjusted by the
' number of bytes actually written, except when the file is opened with
' FILE_FLAG_OVERLAPPED. If the file handle was created for overlapped
' input and output (I/O), the application must adjust the position of
' the file pointer after the write operation is finished.
Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
    lpOverlapped As OVERLAPPED) As Long
    
' System Functions end
'------------------------------------------------------------------

'-------------------------------------------------------------------------------
' GetSystemMessage - Gets system error text for the specified error code.
'-------------------------------------------------------------------------------
Private Function GetSystemMessage(lngErrorCode As Long) As String

  Dim intPos As Integer
  Dim strMessage As String, strMsgBuff As String * 256

    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff, 255, 0)

    intPos = InStr(1, strMsgBuff, vbNullChar)
    If intPos > 0 Then
        strMessage = Trim$(Left$(strMsgBuff, intPos - 1))
    Else
        strMessage = Trim$(strMsgBuff)
    End If
    
    GetSystemMessage = strMessage
    
End Function

'-------------------------------------------------------------------------------
' CommOpen - Opens/Initializes serial port.
'
'
' Parameters:
'   strPort     - COM port name. (COM1, COM2, COM3, COM4)
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=N data=8 stop=1"
'
' Returns:
'   Error Code  - 0 = No Error.
'
'-------------------------------------------------------------------------------
Private Function CommOpen(strPort As String, _
    strSettings As String) As Long
  '
    
  Dim lngStatus       As Long
  Dim udtCommTimeOuts As COMMTIMEOUTS

    On Error GoTo Routine_Error
    
    ' See if port already in use.
    If udtPort1.blnPortOpen Then
        lngStatus = -1
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = "Port in use."
        End With
        
        GoTo Routine_Exit
    End If

    ' Open serial port.
    udtPort1.lngHandle = CreateFile(strPort, GENERIC_READ Or _
        GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If udtPort1.lngHandle = -1 Then
        lngStatus = SetCommError("CommOpen (CreateFile)")
        GoTo Routine_Exit
    End If

    udtPort1.blnPortOpen = True

    ' Setup device buffers (1K each).
    lngStatus = SetupComm(udtPort1.lngHandle, 1024, 1024)
    
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetupComm)")
        GoTo Routine_Exit
    End If

    ' Purge buffers.
    lngStatus = PurgeComm(udtPort1.lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (PurgeComm)")
        GoTo Routine_Exit
    End If

    ' Set serial port timeouts.
    With udtCommTimeOuts
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutMultiplier = 0
        .ReadTotalTimeoutConstant = 1000
        .WriteTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutMultiplier = 1000
    End With

    lngStatus = SetCommTimeouts(udtPort1.lngHandle, udtCommTimeOuts)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommTimeouts)")
        GoTo Routine_Exit
    End If

    ' Get the current state (DCB).
    lngStatus = GetCommState(udtPort1.lngHandle, _
        udtPort1.udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (GetCommState)")
        GoTo Routine_Exit
    End If

    ' Modify the DCB to reflect the desired settings.
    lngStatus = BuildCommDCB(strSettings, udtPort1.udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    ' Set the new state.
    lngStatus = SetCommState(udtPort1.lngHandle, _
        udtPort1.udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommOpen = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommOpen"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

Private Function SetCommError(strFunction As String) As Long
    
    With udtCommError
        .lngErrorCode = Err.LastDllError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        SetCommError = .lngErrorCode
    End With
    
End Function

Private Function SetCommErrorEx(strFunction As String, lngHnd As Long) As Long
  Dim lngErrorFlags As Long
  Dim udtCommStat As COMSTAT
    
    With udtCommError
        .lngErrorCode = GetLastError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
    
        Call ClearCommError(lngHnd, lngErrorFlags, udtCommStat)
    
        .strErrorMessage = .strErrorMessage & "  COMM Error Flags = " & _
                Hex$(lngErrorFlags)
        
        SetCommErrorEx = .lngErrorCode
    End With
    
End Function

'-------------------------------------------------------------------------------
' CommSet - Modifies the serial port settings.
'
' Parameters:
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=N data=8 stop=1"
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommSet(strSettings As String) As Long
  '
  Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommState(udtPort1.lngHandle, _
        udtPort1.udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (GetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = BuildCommDCB(strSettings, udtPort1.udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    lngStatus = SetCommState(udtPort1.lngHandle, _
        udtPort1.udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommSet = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSet"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommClose - Close the serial port.
'
' Parameters:
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommClose() As Long
  'Ensure this is executed on close down to release the handle
  Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    If udtPort1.blnPortOpen Then
        lngStatus = CloseHandle(udtPort1.lngHandle)
    
        If lngStatus = 0 Then
            lngStatus = SetCommError("CommClose (CloseHandle)")
            GoTo Routine_Exit
        End If
    
        udtPort1.blnPortOpen = False
    End If

    lngStatus = 0

Routine_Exit:
    CommClose = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommClose"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommRead - Read serial port input buffer.
'
' Parameters:
'   strInData     - Data buffer.
'   lngSize     - Maximum number of bytes to be read.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommRead(strInData As String, _
    lngSize As Long) As Long
  'Please note that this routine CommRead does not wait for data to be received.
  'A common scheme is to set a timer and periodically read the port.

  Dim lngStatus As Long
  Dim lngRdSize As Long, lngBytesRead As Long
  Dim lngRdStatus As Long, strRdBuffer As String * 1024
  Dim lngErrorFlags As Long, udtCommStat As COMSTAT
    
    On Error GoTo Routine_Error

    strInData = ""
    lngBytesRead = 0
    DoEvents
    
    ' Clear any previous errors and get current status.
    lngStatus = ClearCommError(udtPort1.lngHandle, lngErrorFlags, _
        udtCommStat)

    If lngStatus = 0 Then
        lngBytesRead = -1
        lngStatus = SetCommError("CommRead (ClearCommError)")
        GoTo Routine_Exit
    End If
        
    If udtCommStat.cbInQue > 0 Then
        If udtCommStat.cbInQue > lngSize Then
            lngRdSize = udtCommStat.cbInQue
        Else
            lngRdSize = lngSize
        End If
    Else
        lngRdSize = 0
    End If

    If lngRdSize Then
        lngRdStatus = ReadFile(udtPort1.lngHandle, strRdBuffer, _
            lngRdSize, lngBytesRead, udtCommOverlap)

        If lngRdStatus = 0 Then
            lngStatus = GetLastError
            If lngStatus = ERROR_IO_PENDING Then
                ' Wait for read to complete.
                ' This function will timeout according to the
                ' COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
                ' Every time it times out, check for port errors.

                ' Loop until operation is complete.
                While GetOverlappedResult(udtPort1.lngHandle, _
                    udtCommOverlap, lngBytesRead, True) = 0
                                    
                    lngStatus = GetLastError
                                        
                    If lngStatus <> ERROR_IO_INCOMPLETE Then
                        lngBytesRead = -1
                        lngStatus = SetCommErrorEx( _
                            "CommRead (GetOverlappedResult)", _
                            udtPort1.lngHandle)
                        GoTo Routine_Exit
                    End If
                Wend
            Else
                ' Some other error occurred.
                lngBytesRead = -1
                lngStatus = SetCommErrorEx("CommRead (ReadFile)", _
                    udtPort1.lngHandle)
                GoTo Routine_Exit
            
            End If
        End If
    
        strInData = Left$(strRdBuffer, lngBytesRead)
    End If

Routine_Exit:
    CommRead = lngBytesRead
    Exit Function

Routine_Error:
    lngBytesRead = -1
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommRead"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommWrite - Output data to the serial port.
'
' Parameters:
'   strData     - Data to be transmitted.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommWrite(strData As String) As Long
    
  Dim i As Integer
  Dim lngStatus As Long, lngSize As Long
  Dim lngWrSize As Long, lngWrStatus As Long
    
    On Error GoTo Routine_Error
    
    ' Get the length of the data.
    lngSize = Len(strData)

    ' Output the data.
    lngWrStatus = WriteFile(udtPort1.lngHandle, strData, lngSize, _
        lngWrSize, udtCommOverlap)

    ' Note that normally the following code will not execute because the driver
    ' caches write operations. Small I/O requests (up to several thousand bytes)
    ' will normally be accepted immediately and WriteFile will return true even
    ' though an overlapped operation was specified.
        
    DoEvents
    
    If lngWrStatus = 0 Then
        lngStatus = GetLastError
        If lngStatus = 0 Then
            GoTo Routine_Exit
        ElseIf lngStatus = ERROR_IO_PENDING Then
            ' We should wait for the completion of the write operation so we know
            ' if it worked or not.
            '
            ' This is only one way to do this. It might be beneficial to place the
            ' writing operation in a separate thread so that blocking on completion
            ' will not negatively affect the responsiveness of the UI.
            '
            ' If the write takes long enough to complete, this function will timeout
            ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
            ' At that time we can check for errors and then wait some more.

            ' Loop until operation is complete.
            While GetOverlappedResult(udtPort1.lngHandle, _
                udtCommOverlap, lngWrSize, True) = 0
                                
                lngStatus = GetLastError
                                    
                If lngStatus <> ERROR_IO_INCOMPLETE Then
                    lngStatus = SetCommErrorEx( _
                        "CommWrite (GetOverlappedResult)", _
                        udtPort1.lngHandle)
                    GoTo Routine_Exit
                End If
            Wend
        Else
            ' Some other error occurred.
            lngWrSize = -1
                    
            lngStatus = SetCommErrorEx("CommWrite (WriteFile)", _
                udtPort1.lngHandle)
            GoTo Routine_Exit
        
        End If
    End If
    
    For i = 1 To 10
        DoEvents
    Next
    
Routine_Exit:
    CommWrite = lngWrSize
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommWrite"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

Private Function getState(ByVal iLineType As Integer, ByVal sTipe As String, ByRef sErrRetrn As String) As String
  'User supplies sErrRetrn so any error message can be returned,
  'blank signifies no problems.
  Const FNREF = "getState "
  On Error GoTo ErorHandlr
  Dim lErRetn As Long
  Dim bState As Boolean
  lErRetn = CommGetLine(iLineType, bState)
  If lErRetn = 0 Then
    'OK return
    sErrRetrn = "" 'say No problems
    If bState = True Then
      getState = sTipe & "=1 "
    Else
      getState = sTipe & "=0 "
    End If
  Else
    'Error seen amcc debug
    sErrRetrn = FNREF & sTipe & " Error: " & Str(lErRetn) & " " & GetSystemMessage(lErRetn)
    getState = sTipe & "=? "
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  sErrRetrn = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

'-------------------------------------------------------------------------------
' CommGetLine - Get the state of selected serial port control lines.
'
' Parameters:
'   intLine     - Serial port line. CTS, DSR, RING, RLSD (CD)
'   blnState    - Returns state of line (Cleared or Set).
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommGetLine(ByVal intLine As Integer, _
   ByRef blnState As Boolean) As Long
    
  Dim lngStatus As Long
  Dim lngComStatus As Long, lngModemStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommModemStatus(udtPort1.lngHandle, lngModemStatus)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommReadCD (GetCommModemStatus)")
        GoTo Routine_Exit
    End If

    If (lngModemStatus And intLine) Then
        blnState = True
    Else
        blnState = False
    End If
        
    lngStatus = 0
        
Routine_Exit:
    CommGetLine = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommReadCD"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function


'-------------------------------------------------------------------------------
' CommGetError - Get the last serial port error message.
'
' Parameters:
'   strMessage  - Error message from last serial port error.
'
' Returns:
'   Error Code  - Last serial port error code.
'-------------------------------------------------------------------------------
Private Function CommGetError(strMessage As String) As Long
    
    With udtCommError
        CommGetError = .lngErrorCode
        strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & _
            " - " & .strErrorMessage
    End With
    
End Function

'-------------------------------------------------------------------------------
' CommFlush - Flush the send and receive serial port buffers.
'
' Parameters:
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommFlush() As Long
    
  Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = PurgeComm(udtPort1.lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommFlush (PurgeComm)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommFlush = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommFlush"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

Private Function fixLine(ByVal iLineType As Integer, ByVal bInStaet As Boolean) As String
  'Returns empty string if ok else error message
  Const FNREF = "fixLine "
  On Error GoTo ErorHandlr
  Dim lRetn As Long
  fixLine = "" 'Assume no problems
  lRetn = CommSetLine(iLineType, bInStaet)
  If lRetn <> 0 Then
    fixLine = "Problem setting line, Error: " & Str(lRetn) & " " & GetSystemMessage(lRetn)
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  fixLine = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

'-------------------------------------------------------------------------------
' CommSetLine - Set the state of selected serial port control lines.
'
' Parameters:
'   intLine     - Serial port line. BREAK, DTR, RTS
'                 Note: BREAK actually sets or clears a "break" condition on
'                 the transmit data line.
'   blnState    - Sets the state of line (Cleared or Set).
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Private Function CommSetLine(intLine As Integer, _
   blnState As Boolean) As Long
   
  Dim lngStatus As Long
  Dim lngNewState As Long
    
    On Error GoTo Routine_Error
    
    If intLine = LINE_BREAK Then
        If blnState Then
            lngNewState = SETBREAK
        Else
            lngNewState = CLRBREAK
        End If
    
    ElseIf intLine = LINE_DTR Then
        If blnState Then
            lngNewState = SETDTR
        Else
            lngNewState = CLRDTR
        End If
    
    ElseIf intLine = LINE_RTS Then
        If blnState Then
            lngNewState = SETRTS
        Else
            lngNewState = CLRRTS
        End If
    End If

    lngStatus = EscapeCommFunction(udtPort1.lngHandle, lngNewState)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSetLine (EscapeCommFunction)")
        GoTo Routine_Exit
    End If

    lngStatus = 0
        
Routine_Exit:
    CommSetLine = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSetLine"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

Private Sub class_initialize()
  z_SerialPortNumber = 0 'See Note 1
End Sub

'-----------------------------------------------------------------------------------
'-------------API serial port public functions Begin
'-----------------------------------------------------------------------------------
'The following functions communicate with the API serial port...
' clSerialAPIinitialise* (Sets port number and initialises this class E1)
' clSerialAPIsetBaud* (Sets baud rate for closed port E1)
' clSerialAPIopen* (Opens the port E1)
' clSerialAPIread* (Reads the API port E2)
' clSerialAPIwrite* (Writes to the API port E1)
' clSerialAPIclose* (Closes or ensures port is closed E1)
' clSerialAPIflush (Flush the send and receive serial port buffers E1)
' clSerialAPIsetBREAK (Set BREAK line E1)
' clSerialAPIsetDTR (Set DTR line E1)
' clSerialAPIsetRTS (Set RTS line E1)

'The following functions provide information about API serial port...
' clSerialAPIgetBaud (returns a integer Baud rate)
' clSerialAPIgetIsOpen (returns boolean true if port is open)
' clSerialAPIgetLines (returns string of 1 or 0 else error message)
' clSerialAPIgetVersion (returns the build version of serial API class)

'NOTES:
'  * Starred functions comprise the minimum set useful.
'  E1 Returns error message else blank string if ok
'  E2 User must supply a string as a parameter otherwise return is as E1
'----------------------------------------------------------------------

Public Function clSerialAPIinitialise(ByVal iSerPortNumbre As Integer) As String
  'Initialise serial port class
  'Returns error message else blank string if ok
  Const FNREF = "clSerialAPIinitialise "
  On Error GoTo ErorHandlr
  
  'Note 1 Ensure port numbers are greater than 0 as this is used to show port is not initialised.
  If iSerPortNumbre < 1 Then
    clSerialAPIinitialise = "API serial port number must exceed 0"
    Exit Function
  End If
  
  'Test for port already being used
  If z_SerialPortNumber = iSerPortNumbre Then
    clSerialAPIinitialise = "API serial port " & iSerPortNumbre & " is already in use"
    Exit Function
  End If
  
  clSerialAPIinitialise = "" 'All is well
  z_SerialPortNumber = iSerPortNumbre
  z_iBaud = 0 'To indicate port is not set up
  z_bIsOpen = False
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIinitialise = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIsetBaud(ByVal iBaudIn As Integer) As String
  'This sets up an unopened port. It does not open the port.
  'Returns empty string if ok else error message.
  Const FNREF = "clSerialAPIsetBaud "
  On Error GoTo ErorHandlr
  
  'Only allow a closed port to have baud changed
  If z_bIsOpen = True Then
    clSerialAPIsetBaud = "Close port first then set baud"
    Exit Function
  End If
  z_iBaud = iBaudIn
  clSerialAPIsetBaud = ""
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIsetBaud = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIopen() As String
  'Returns blank string if okay else message containing the problem
  'The caller is advised to check the return string.
  'Parity, data bits and stop bits are hardcoded below.
  'Baud rate must be set before this function is called.
  'Note O1: The COM port name format can vary. Any of the following formats
  'may be valid depending on the system and serial port driver.
  'Try the others if the port won't open.
  '  COM1
  '  COM1:
  '  \\.\COM1

  Const FNREF = "clSerialAPIopen "
  On Error GoTo ErorHandlr
  
  Dim lngStaatus As Long
  Dim strError  As String
  clSerialAPIopen = "" 'Default to no problems seen

  If z_iBaud = 0 Then
    clSerialAPIopen = "Port not set up so cannot open it"
    Exit Function
  End If

  If z_bIsOpen = True Then
    clSerialAPIopen = "Port is already open"
    Exit Function
  End If

  ' Initialize Communications (See above Note O1)
  lngStaatus = CommOpen("COM" & CStr(z_SerialPortNumber), _
      "baud=" & z_iBaud & " parity=N data=8 stop=1")
  
  If lngStaatus <> 0 Then
    lngStaatus = CommGetError(strError) 'function fills in strError
    clSerialAPIopen = "COM Error: " & strError
    z_bIsOpen = False 'Define port being closed
  Else
    z_bIsOpen = True
    ' Set modem control lines.
    lngStaatus = CommSetLine(LINE_RTS, True)
    lngStaatus = CommSetLine(LINE_DTR, True)
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIopen = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIread(ByVal iMaxBytesIn As Long, ByRef sReturnError As String) As String
  'Returns blank string if okay else message
  'INPUT iMaxBytesIn is maximum number of bytes to be read in
  'OUTPUT The user must supply a string variable (parameter sReturnError)
  'for any error messages this will be set blank if no problems found.
  Const FNREF = "clSerialAPIread "
  On Error GoTo ErorHandlr
  Dim lngStaatus As Long
  Dim strData As String
  clSerialAPIread = ""
  sReturnError = "" 'Assume no problems
  
  If z_iBaud = 0 Then
    sReturnError = "Attempt to read an unitialised port "
    Exit Function
  End If
  
  If z_bIsOpen = False Then
    sReturnError = "Attempt to read a closed port "
    Exit Function
  End If
  
  ' Read maximum of iMaxBytesIn bytes from serial port.
  lngStaatus = CommRead(strData, iMaxBytesIn)
  If lngStaatus > 0 Then
    clSerialAPIread = strData ' Get data. All okay.
  ElseIf lngStaatus < 0 Then
    sReturnError = "API read handle error" ' Handle error.
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  sReturnError = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIwrite(ByVal strData As String) As String
  'Returns blank string if okay else message
  Const FNREF = "clSerialAPIwrite "
  On Error GoTo ErorHandlr
  Dim lngSize, lngStaatus As Long
  
  clSerialAPIwrite = "" 'Assume no problems
  If z_iBaud = 0 Then
    clSerialAPIwrite = "Attempt to write to an unitialised port "
    Exit Function
  End If
  
  lngSize = Len(strData)
  lngStaatus = CommWrite(strData)
  If lngStaatus <> lngSize Then
    clSerialAPIwrite = "API write handle error"
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIwrite = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIclose() As String
  'Returns blank string if okay else message
  Const FNREF = "clSerialAPIclose "
  On Error GoTo ErorHandlr
  Dim lRetn As Long
  'Close communications.
  'More than one close is not harmful so no need for complications
  clSerialAPIclose = "" 'Assume no problems
  lRetn = CommClose 'returns 0 if ok
  If lRetn = 0 Then
    z_bIsOpen = False 'To say closed
  Else
    clSerialAPIclose = "Problem closing " & Str(lRetn)
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIclose = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIflush() As String
  'Returns blank string if okay else message
  Const FNREF = "clSerialAPIclose "
  On Error GoTo ErorHandlr
  Dim lRetn As Long
  clSerialAPIflush = "" 'Assume operation is ok
  
  If z_bIsOpen = False Then
    clSerialAPIflush = "Attempt to flush a closed port"
    Exit Function
  End If
    
  lRetn = CommFlush
  If lRetn <> 0 Then
    clSerialAPIflush = "Problem flushing, Error: " & Str(lRetn) & " " & GetSystemMessage(lRetn)
  End If

  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIflush = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Function clSerialAPIsetBREAK(ByVal bInStaet As Boolean) As String
  clSerialAPIsetBREAK = fixLine(LINE_BREAK, bInStaet)
End Function

Public Function clSerialAPIsetDTR(ByVal bInStaet As Boolean) As String
  clSerialAPIsetDTR = fixLine(LINE_DTR, bInStaet)
End Function

Public Function clSerialAPIsetRTS(ByVal bInStaet As Boolean) As String
  clSerialAPIsetRTS = fixLine(LINE_RTS, bInStaet)
End Function

'The following get information about the port
Public Property Get clSerialAPIgetBaud() As Integer
  clSerialAPIgetBaud = z_iBaud
End Property

Public Property Get clSerialAPIgetIsOpen() As Boolean
  clSerialAPIgetIsOpen = z_bIsOpen
End Property

Public Function clSerialAPIgetLines() As String
  'Returns a series of 1 or 0 to represent line states of CD CTS DSR RING RLSD
  'or an error message
  Const FNREF = "clSerialAPIgetLines "
  On Error GoTo ErorHandlr
  Dim sSts, sErrRetn As String
  
  If z_bIsOpen = False Then
    clSerialAPIgetLines = "Open port before changing line states"
  End If
  
  sSts = ""
  sErrRetn = "" 'No error assumed, this will be passed to getState to populate if error seen
  
  Do While True
    'A once through only loop
    sSts = sSts + getState(LINE_CD, "CD", sErrRetn)
    If sErrRetn <> "" Then
      Exit Do
    End If
    sSts = sSts + getState(LINE_CTS, "CTS", sErrRetn)
    If sErrRetn <> "" Then
      Exit Do 'with error in sErrRetn
    End If
    sSts = sSts + getState(LINE_DSR, "DSR", sErrRetn)
    If sErrRetn <> "" Then
      Exit Do
    End If
    sSts = sSts + getState(LINE_RING, "RING", sErrRetn)
    If sErrRetn <> "" Then
      Exit Do
    End If
    sSts = sSts + getState(LINE_RLSD, "RLSD", sErrRetn)
'    If sErrRetn <> "" Then
'      Exit Do
'    End If
    Exit Do 'Important as once only
  Loop
  If sErrRetn = "" Then
    clSerialAPIgetLines = sSts 'No problems
  Else
    clSerialAPIgetLines = sErrRetn 'Problem(s) seen
  End If
  Exit Function ' Exit to avoid handler.
ErorHandlr:   ' Error-handling routine.
  clSerialAPIgetLines = FNREF & "Error " & Err.Number & ", " & Err.Description
End Function

Public Property Get clSerialAPIgetVersion() As String
  'Use as for example... MsgBox(YourClass.clSerialAPIgetVersion)
  clSerialAPIgetVersion = scVersion
End Property

'-------------API serial port public functions End



