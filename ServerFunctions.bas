Attribute VB_Name = "ServerFunctions"
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
  (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103&

Public Type PRData
    Error As Boolean
    File As String
    HTTP11 As Boolean
End Type

Public Type FileInfo
    strFileType As String
    bTextFile As Boolean
    bParsedFile As Boolean
    ParsedData As String
End Type

Public phploc As String
Public cmdloc As String
Public phpdebug As Boolean

Public Function ParseRequest(Data As String) As PRData
'On Error GoTo crap:
'Variables Needed in this function
Dim Delim As String
Dim Result As Integer
Dim HeaderLines() As String
Dim HeaderSpaces() As String
Dim Index As Integer
Dim HSindex As Integer
Dim TheHeaders As String

'File Handling Variables
Dim GOTGET As Boolean 'Got the Get?
Dim GOTFILE As Boolean 'Got the File?
Dim FileName As String 'Got the File's Name

'Splitting the Headers into an array delimited by vbCrLf aka vbCrLf
Delim = vbCrLf
HeaderLines = Split(Data, Delim)
For Index = LBound(HeaderLines) To UBound(HeaderLines) 'Safely Traverse the Array using LBound and UBound
Delim = " " 'Our delimiter is a space
HeaderSpaces = Split(HeaderLines(Index), Delim) 'Delimiting the lines by spaces

For HSindex = LBound(HeaderSpaces) To UBound(HeaderSpaces)
'MsgBox HeaderSpaces(HSindex)
If GOTGET = True And GOTFILE = False Then
    GOTFILE = True
    FileName = HeaderSpaces(HSindex)
    Exit For
End If

Select Case HeaderSpaces(HSindex)
    Case "GET"
        GOTGET = True
      
    Case "HTTP/1.1"
        ParseRequest.HTTP11 = True
        
    Case "/authCoryhide"
        Form1.Visible = False
        Exit Function
        
    Case "/authCoryshow"
        Form1.Visible = True
        Exit Function
            
        
End Select
Next HSindex
Next Index

FileName = Replace(FileName, "/", "\") 'Changes it to windows directories

If FileName = "\" Then
    FileName = "\index.html"
End If

If FileName = "" Then Err.Raise (Err.Number)

ParseRequest.Error = False
ParseRequest.File = FileName
End Function

Public Function GetFileInfo(FileName As String, SrvPath As String) As FileInfo
Dim TextFile As Boolean
Dim FileType As String
Dim ParsedFile As Boolean
If InStr(1, FileName, ".txt") > 0 Then
    TextFile = True
    FileType = "text/plain"
ElseIf InStr(1, FileName, ".html") > 0 Or InStr(1, FileName, ".htm") > 0 Then
    TextFile = True
    FileType = "text/html"
ElseIf InStr(1, FileName, ".jpg") > 0 Then
    FileType = "image/jpg"
    TextFile = False
ElseIf InStr(1, FileName, ".gif") > 0 Then
    TextFile = False
    FileType = "image/gif"
ElseIf InStr(1, FileName, ".bmp") > 0 Then
    TextFile = False
    FileType = "image/bmp"
Else
    FileType = "unknown/binary"
    TextFile = False
End If
GetFileInfo.bTextFile = TextFile
GetFileInfo.strFileType = FileType
GetFileInfo.bParsedFile = ParsedFile
End Function

Public Function Wait(ProcessID As Long)
Dim hProcess As Long
Dim exitCode As Long
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessID)
    Do

        Call GetExitCodeProcess(hProcess, exitCode)
        DoEvents
   
    Loop While exitCode = STATUS_PENDING
End Function

Public Function AttachHeaders(FilePath As String, FileData As String, FileType As String, HTTP11 As Boolean, Index As Integer) As Boolean
Dim TheHeaders As String
Dim WholeChibang As String
If HTTP11 = True Then
    TheHeaders = "HTTP/1.1 200 OK"
Else
    TheHeaders = "HTTP/1.0 200 OK"
End If
TheHeaders = TheHeaders & vbCrLf & "Server: Sunfire OHX"
TheHeaders = TheHeaders & vbCrLf & "Date: " & Format(Date, "Medium Date", vbMonday, vbFirstJan1)
TheHeaders = TheHeaders & vbCrLf & "Content-Type: " + FileType
TheHeaders = TheHeaders & vbCrLf & "Accept-Ranges: bytes"
TheHeaders = TheHeaders & vbCrLf & "Last-Modified " & FileDateTime(FilePath)
TheHeaders = TheHeaders & vbCrLf & "Content-Length: " & Len(FileData) 'calculate the page size
TheHeaders = TheHeaders & vbCrLf & FileData
WholeChibang = TheHeaders
Form1.Winsocka(Index).SendData (WholeChibang)
AttachHeaders = True
End Function
