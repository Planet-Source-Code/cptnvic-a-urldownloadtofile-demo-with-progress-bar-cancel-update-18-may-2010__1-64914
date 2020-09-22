VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VicsDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--> declare api for downloading file
'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'-- the above declare is not needed since is implemented in the type lib. but I leave it here
'-- incase you take bits/pieces out of the class.

'--> declare api for deleting an existing file from IE's cache
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
'--> needed constants
Const ERROR_SUCCESS As Long = 0
Const BINDF_GETNEWESTVERSION As Long = &H10
Const INTERNET_FLAG_RELOAD As Long = &H80000000
Const S_OK As Long = &H0

'--> Some quickly constructed IBindStatusCallback events for you to use... or not as you wish
'--> See the events/uses in the frmDnLoad code

Public Event VicDLStart()  'notify when current download is started
Public Event VicDLDone()  'notify when current download is finished
Public Event VicDLCrash(ByVal VicErrNum As Long, VicErrDescr As String, Cancelled As Boolean) 'notify when download was interrupted
Public Event VicDLCancelled()  'notify when a cancel request was received
Public Event VicDLProg(ByVal VicBytesIn As Long, ByVal VicTotalBytes As Long) 'notify of progress during download

Implements olelib.IBindStatusCallback 'initialize the IBindStatusCallback interface...

Private VicBinding As IBinding  ' set-up so I can access as I wish

Private Sub Class_Terminate()
    If VicBinding Is Nothing Then
        Exit Sub
    Else
        Set VicBinding = Nothing
    End If
End Sub

Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As olelib.BINDSTATUS, ByVal szStatusText As Long)
    'the OnProgress event is what keeps your app from freezing during download ala URLDownloadToFile
    '--> just an fyi... IBindStatusCallback_OnProgress returns results in terms of bytes...
    '--> ulProgressMax = the total # bytes to be downloaded
    '--> ulProgress = the # of bytes downloaded thus far
    '--> ...you need to base your calculations (if any) in those terms
    
    'if you downloaded the original version of this project you will remember that ulProgressMax can return null
    'which really screws up a progress bar... so protect against that possibility here... and
    'handle the ulProgress error possibilities in the VicDLProg event
    If ulProgressMax > 0 Then
        RaiseEvent VicDLProg(ulProgress, ulProgressMax) 'report progress to the event
    End If
    'this is a good place to let a cancel request get in line for processing
    DoEvents
    'leave this sub here even if you don't utilize it!
    'all events need to be exposed... and so on for all IBind subs below
    'see: http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/moniker/reference/ifaces/urlmon_ref_ifaces_entry.asp
End Sub

Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As olelib.IBinding)
    Set VicBinding = pib 'get something to read from callback... and get in the "stream" of things - so to speak.
    RaiseEvent VicDLStart
    '(binding has started.  See the frmDnLoad code for some reasons to use this info)
    'leave this sub here even if you don't utilize it!
End Sub

Private Sub IBindStatusCallback_OnStopBinding(ByVal hresult As Long, ByVal szError As Long)
    'Binding has stopped... figure out why if you want to know...
    If hresult = 1 Or hresult = S_OK Then
        'Download Completed normally...
        RaiseEvent VicDLDone
    Else
        'Download was aborted/or cancelled:
        'You would think that szError would be the logical number to interpret...
        'But, it's not!  hresult is the error number to interpret...
        'if it's a failure... it's most likely a server/connection problem/or file not found
        Dim Cancelled As Boolean
        If VicBinding Is Nothing Then
            Cancelled = True
        Else
            Cancelled = False
        End If
        RaiseEvent VicDLCrash(hresult, GetVicError(hresult), Cancelled) 'report the failure to the event for handling
    End If
    Set VicBinding = Nothing
    'leave this sub here even if you don't utilize it!
    'See the frmDnLoad code for some reasons to use this info
End Sub

Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As olelib.BINDF, pbindinfo As olelib.BINDINFO)
    frmDemo.Text3.Text = frmDemo.Text3.Text & "GetBindInfo." & vbCrLf
    'leave this sub here even if you don't utilize it!
End Sub

Private Function IBindStatusCallback_GetPriority() As Long
    frmDemo.Text3.Text = frmDemo.Text3.Text & "GetPriority." & vbCrLf
    'leave this sub here even if you don't utilize it!
End Function

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As olelib.BSCF, ByVal dwSize As Long, pformatetc As olelib.FORMATETC, pStgmed As olelib.STGMEDIUM)
    frmDemo.Text3.Text = frmDemo.Text3.Text & "OnDataAvailable." & vbCrLf
    'leave this sub here even if you don't utilize it!
End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)
    frmDemo.Text3.Text = frmDemo.Text3.Text & "OnLowResource." & vbCrLf
    'leave this sub here even if you don't utilize it!
End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(riid As olelib.UUID, ByVal pUnk As stdole.IUnknown)
    frmDemo.Text3.Text = frmDemo.Text3.Text & "OnObjectAvailable." & vbCrLf
    'leave this sub here even if you don't utilize it!
End Sub

Public Function StartTheStinkinDownLoad(ByVal File2DownLoad As String, ByVal File2Save As String)
    Dim DownLoadResult As Long
    '--> Start the download...
    '--> You should note that URLDownloadToFile overwrites any existing file without notification!
    If VicBinding Is Nothing Then
        DownLoadResult = olelib.URLDownloadToFile(Nothing, File2DownLoad, File2Save, 0, Me)
        '--> Report the results of the download attempt
        StartTheStinkinDownLoad = (DownLoadResult = olelib.S_OK)
    End If
End Function

Public Sub KillVic()
    On Error GoTo NoBind2
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++ Gabe, this is the part you were interested in... sort of! ++
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    VicBinding.Abort
    
    'More comments to bore you:
    
    'E_ABORT (I think) is the only clean method of aborting the download while in binding
    'According to MSDN, there are 3 possibilities for hresult of this attempt:
    '1) S_OK    - Abort Was Successful
    '2) E_FAIL  - Bind Operation Could Not Be Aborted... Why would this be possible????
    '3) S_FALSE - Bind Operation Already Aborted... (Binding was not underway)
    
    'Since it is possible for this request to fail, you should not bail out in this sub...
    'but let binding cease in it's own due course, otherwise a case of limbo could exist.
    'In this demo, I'll send notification of the cancel request in: VicDLCancelled()
    'then actually stop operations in: VicDLCrash(...)
    
    RaiseEvent VicDLCancelled
    Set VicBinding = Nothing
Exit Sub 'protect the error handler
NoBind2:
    frmDemo.Text3.Text = frmDemo.Text3.Text & "Binding Error Ignored" & vbCrLf
    Resume Next
End Sub

Private Function GetVicError(ByVal VicErrNum As Long) As String
    Select Case VicErrNum
        Case INET_E_AUTHENTICATION_REQUIRED
            GetVicError = "Authentication Failure."
        Case INET_E_CANNOT_CONNECT
            GetVicError = "Cannot Connect"
        Case INET_E_CANNOT_INSTANTIATE_OBJECT
            GetVicError = "Cannot Instantiate Object."
        Case INET_E_CANNOT_LOAD_DATA
            GetVicError = "Cannot Load Data."
        Case INET_E_CANNOT_LOCK_REQUEST
            GetVicError = "Cannot Lock Request."
        Case INET_E_CANNOT_REPLACE_SFP_FILE
            GetVicError = "Cannot Replace SFP File."
        Case INET_E_CODE_DOWNLOAD_DECLINED
            GetVicError = "Code Download Declined."
        Case INET_E_CONNECTION_TIMEOUT
            GetVicError = "Connection Timeout."
        Case INET_E_DATA_NOT_AVAILABLE
            GetVicError = "Data Not Available."
        Case INET_E_DEFAULT_ACTION
            GetVicError = "Default Action."
        Case INET_E_DOWNLOAD_FAILURE
            GetVicError = "Download Failure."
        Case INET_E_INVALID_REQUEST
            GetVicError = "Invalid Request."
        Case INET_E_INVALID_URL
            GetVicError = "Invalid URL."
        Case INET_E_NO_SESSION
            GetVicError = "No Session."
        Case INET_E_NO_VALID_MEDIA
            GetVicError = "No Valid Media."
        Case INET_E_OBJECT_NOT_FOUND
            GetVicError = "File Not Found."
        Case INET_E_QUERYOPTION_UNKNOWN
            GetVicError = "QueryOption Unknown."
        Case INET_E_REDIRECT_FAILED
            GetVicError = "Redirect Failed."
        Case INET_E_REDIRECT_TO_DIR
            GetVicError = "Redirect To Dir."
        Case INET_E_REDIRECTING
            GetVicError = "Redirecting."
        Case INET_E_RESOURCE_NOT_FOUND
            GetVicError = "Resource Not Found."
        Case INET_E_RESULT_DISPATCHED
            GetVicError = "Result Dispatched."
        Case INET_E_SECURITY_PROBLEM
            GetVicError = "Security Problem."
        Case INET_E_UNKNOWN_PROTOCOL
            GetVicError = "Unknown Protocol."
        Case Else
            GetVicError = "Unknown."
   End Select
End Function

Public Function DeleteVicCache(Filename As String) As Long
    DeleteVicCache = DeleteUrlCacheEntry(Filename)
End Function
