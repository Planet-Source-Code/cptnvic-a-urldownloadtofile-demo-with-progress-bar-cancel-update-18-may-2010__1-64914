VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDnLoad 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   Enabled         =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "frmDnLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ PLEASE, PLEASE, PLEASE Read the readme.txt files that accompany this download before attempting to use this!                                ++
'++ The type library that you need is NOT NOW or ever going to be found in this download... but can be gotten freely and easily!                ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Author: CptnVic                                                                                                                             ++
'++ Updated: 15 June 2006:                                                                                                                      ++
'++     [] Added cancel ability at request of Gabe.                                                                                             ++
'++     [] Put the IBindStatusCallback stuff in the VicsDL class for ease of use                                                                ++
'++     [] Built few other modest exposures for your coding pleasure... enjoy!                                                                  ++
'++ NOTE: The abort procedure in the cancel routine below is the only (clean) way I know of to cancel URLDownloadToFile while in progress.      ++
'++ However, URLDownloadToFile downloads files in large chunks... so stopping the download is not always possible - especially when small files ++
'++ are being downloaded.  Once binding gets started on larger files, intercepting the download in the IBindStatusCallback_OnProgress event is  ++
'++ fairly straight forward, however, when multiple small files are being downloaded... a heck of alot of stuff is going on... and the order in ++
'++ which events occur are difficult to predict.  Anyway, it seems to work o.k.  Let me know if you have bad experiences with it!               ++
'++ Otherwise, use it (as always) as you see fit.                                                                                               ++
'++ Your most ardent admirer, CptnVic                                                                                                           ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ This code is placed on PSC for the 3 or 4 coders on PSC that actually read, try, and then                                                   ++
'++ leave CONSTRUCTIVE comments/suggestions (and even ocasionally vote!) for code on PSC.                                                       ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ I GOT ALOT OF HELP FROM:                                                                                                                    ++
'++ http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/moniker/reference/ifaces/ibindstatuscallback/ibindstatuscallback.asp ++
'++ AND OF COURSE, Edamo's OLE interfaces & functions v1.81, available freely at:                                                               ++
'++ http://www.mvps.org/emorcillo/download/vb6/tl_ole.zip                                                                                       ++
'++ He has simple re-use requirements... see them at:                                                                                           ++
'++ http://www.mvps.org/emorcillo/en/index.shtml                                                                                                ++
'++ He has some other great stuff at: http://www.mvps.org/emorcillo/en/index.shtml, sadly, they have given up on VB long ago.                   ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ To save the inevitible question... NO YOU DON'T HAVE TO DISTRIBUTE THE TYPE LIBRARY!                                                        ++
'++ When you compile your project, VB will take what it needs and compile it in your *.exe.  You only need the TLB for compiling.               ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ SEE THE README.TXT FILES to set this up!  Otherwise, email someone who cares!                                                               ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private WithEvents mydl As VicsDL
Attribute mydl.VB_VarHelpID = -1
Dim IsCancelRequested As Boolean

Private Sub Form_Unload(Cancel As Integer)
    Set mydl = Nothing
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ The Code that shows the progress bar, draws the form, etc. follows...
'++ First, I'll build the 5 events exposed by the class.
'++
'++ Once you understand when and how these events are fired, you can delete those you don't
'++ want or add new ones as long as you excercise reasonable care.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub mydl_VicDLStart()
    '++ Raised In: IBindStatusCallback_OnStartBinding event
    'this event is fired when the CURRENT download is confirmed to have begun.
    
    'I don't have much use for it in this demo, however, it is very useful information!
    'For example, I have used it in other projects which have multiple progress bars tracking
    'multiple downloads: 1 for the download in progress, 1 for the number of downloads remaining.
    
    'Another good use is to start a timer if you want to time-out a download.  In this scenario,
    'you would start a timer in this event, disable it in the VicDLDone event that follows, OR
    'cause a cancel to happen if the timer times out.  I hope that makes sense?
    
    'For the purposes of this demo, I'll just record the event occurring
    frmDemo.Text3.Text = frmDemo.Text3.Text & "VicDLStart Event Occurred" & vbCrLf
End Sub
Private Sub mydl_VicDLDone()
    '++ Raised In: IBindStatusCallback_OnStopBinding event
    'This event is fired when the CURRENT download is confirmed to have been completed.
    'Most likely used in conjunction with the VicDLStart event above
    
    'I don't need it for this demo... so I'll just document it's existence
    frmDemo.Text3.Text = frmDemo.Text3.Text & "VicDLDone Event Occurred" & vbCrLf
End Sub
Private Sub mydl_VicDLCrash(ByVal VicErrNum As Long, VicErrDescr As String, Cancelled As Boolean)
    '++ Raised In: IBindStatusCallback_OnStopBinding event
    'This event is fired when the download was NOT completed for any reason...
    'some failure, or was maybe cancelled.
    
    'It reveals what happened... if you care!
    If Cancelled Then
        'failure was caused by a cancel... you may not care? ...
        'but I want to use this event to unload this form, prevent further downloads, and
        'get outta here!  Setting IsCancelRequested = True will get me out of the loop... if I'm still in it.
        
        frmDemo.Text3.Text = frmDemo.Text3.Text & "Download was cancelled successfully." & vbCrLf
        IsCancelRequested = True
    Else
        'do a message box here?
        'VicErrNum = the error number returned
        'VicErrDescr =  the text description
        Dim Msg As String, Title As String
        Title = "Holy Moly!  A Crappy Thing Has Happened!"
        Msg = "The Bad Thing's Number is: " & VicErrNum & vbCrLf & "The Bad Thing Was:" & vbCrLf & VicErrDescr
        MsgBox Msg, vbOKOnly + vbExclamation, Title
        '--> you may want to abort further download action here?????? Or Continue?????
    End If
    
End Sub
Private Sub mydl_VicDLCancelled()
    '++ Raised In: KillVic sub
    'this event is fired when you have passed the cancel (abort) request.
    'I'd just use this event to notify your user that you are attempting a cancel... as in:
    
    Label1.Caption = "Cancelling..."
    
    'You can deal with this as you like, for now, I'll report it to frmDemo
    frmDemo.Text3.Text = frmDemo.Text3.Text & "Cancel Request Received" & vbCrLf
    
End Sub
Private Sub mydl_VicDLProg(ByVal VicBytesIn As Long, ByVal VicTotalBytes As Long)
    On Error GoTo OhCrap
    '++ Raised In: IBindStatusCallback_OnProgress event
    'use this event's info to update your progress bar, etc.
    
    'VicBytesIn = # of BYTES downloaded so far = ulProgress
    'VicTotalBytes = Total # of BYTES to ultimately be downloaded = ulProgressMax
    '--> URLDownloadToFile sometimes freaks out here... so control the damage...
    '    and it will catch back up with itself.
    'Here are a few combinations that I have observed while debugging:
    'ulProgress = 0: ulProgressMax = 0 -> Set ProgressBar1.Max = 0 fires error
    'ulProgress > ulProgressMax -> Set ProgressBar1.Value > ProgressBar1.Max fires error
    'I've already trapped the ulProgressMax errors in IBindStatusCallback_OnProgress
    'so all that's left to guard against is:
    
    'handle the ulProgress error possibilities
    If VicBytesIn >= 0 And VicBytesIn <= VicTotalBytes Then
        ProgressBar1.Max = VicTotalBytes ' set/re-set the progress bar's max value after it is known for sure
        '-->Be sure to set the max value before assigning the bar value!
        ProgressBar1.Value = VicBytesIn ' set the current level of progress
        DoEvents 'force a refresh... even though this slows things down
    End If
Exit Sub
OhCrap:
    'this shouldn't ever fire... but there's no sense in letting your progress bar screw
    'things up now!  To my way of thinking, it's better to have a slightly mis-informed
    'user than a bad download and crash.  An error here is caused by the Progress bar.
    'I guess, if your using this to download movies (or similar) your progress bar's Max
    'and Value limits could be exceeded... so if you need to, you can use this handler to
    'hide the progress bar and switch to a text only progress update?
    Resume Next
End Sub

Public Sub EchoCancelRequest()
    'Echo a cancel request (if you can) from here so that this form will receive the events
    'Note that this is a public sub so the request from frmDemo can get to it.
    mydl.KillVic
End Sub

Public Function ShowDownLoad(FileList As String, CallingForm As Form, Optional Owner As Object)
    Set mydl = New VicsDL 'implement the class on this form
    'this would usually be in the form_load event... but I do not use that event in this project
    
    'be sure the focus is set on the calling form so download can be cancelled easier
    CallingForm.SetFocus
    IsCancelRequested = False
    DoEvents
    DoFormStuff 'draw the form
    If IsMissing(Owner) = False Then
        Me.Show 'You're better off to show without owner form... otherwise the function will wait till you close the form before it does anything... :(
    Else
        Me.Show vbModeless, Owner 'I leave this incase you want response.
    End If
    Me.Refresh 'force form to be displayed 1st before processing the code that follows
    'split files to download from FileList
    Dim i, X As Integer
    Dim File2DownLoad As String, File2Save As String, DeleteCache As Boolean, TopLimit As Integer, TempDelete As String, OffSet As Integer
    i = Split(FileList, ",")
    TopLimit = (UBound(i) - 2) / 3 'filelist comes in as:File2DownLoad,File2Save,DeleteCache
    OffSet = 0
    For X = 0 To TopLimit '<-- start the processing loop and check occasionally for a cancel
        
        '--> Before beginning, check to see if a cancel request was received
        If IsCancelRequested Then Exit For 'if so, leave this loop - otherwise, more files could be downloaded
        
        '--> no cancel request exists... so, start processing files <--
        File2DownLoad = i(OffSet)
        File2Save = i(OffSet + 1)
        TempDelete = i(OffSet + 2)
        If TempDelete = "1" Then
            DeleteCache = True
        Else
            DeleteCache = False
        End If
        OffSet = OffSet + 3 'increment the offset for next file
        ProgressBar1.Value = 0 'initialize the progress bar
        
        'inform the calling form of action for purposes of this demo
        frmDemo.Text3.Text = frmDemo.Text3.Text & "Starting Download..." & vbCrLf
        frmDemo.Text3.Text = frmDemo.Text3.Text & File2DownLoad & vbCrLf
        
        'You may want to download the file from IE's cache... if so... set DeleteCache = False
        'however, this may result in an old file being "downloaded" from the cache and not the internet
        'Note that the remote URL is passed since this is the name that the cached file is known by.
        'This does NOT delete the file from the remote server... ONLY the local machine copy
        'Deleting the cached copy (if it exists) forces a new copy to be downloaded from internet
        
        If DeleteCache Then
            If mydl.DeleteVicCache(File2DownLoad) = 1 Then 'file was found and deleted
                frmDemo.Text3.Text = frmDemo.Text3.Text & "Found Cached File and Deleted It..." & vbCrLf
            Else
                frmDemo.Text3.Text = frmDemo.Text3.Text & "Did Not Find Cached Copy Of Requested File" & vbCrLf 'no local copy existed
            End If
        End If
        Label1.Caption = File2DownLoad
        
        'proceed with the download part
        If mydl.StartTheStinkinDownLoad(File2DownLoad, File2Save) Then
            frmDemo.Text3.Text = frmDemo.Text3.Text & File2DownLoad & " Download Completed!" & vbCrLf
            ShowDownLoad = True
            '-->you may want some other notification back to calling form here
        Else
            frmDemo.Text3.Text = frmDemo.Text3.Text & "File Download Failed!" & vbCrLf
            ShowDownLoad = False
            '-->you may want some other notification back to calling form here
        End If
    Next
BailingOut:
    frmDemo.Text3.Text = frmDemo.Text3.Text & "Ending Download(s)..." & vbCrLf 'report leaving for the heck of it
    frmDemo.cmdCancel.Visible = False
    Set mydl = Nothing 'free memory
    Unload Me 'report the last ShowDownLoad state to owner if any
    Set frmDnLoad = Nothing 'loose every shred
End Function

Private Sub DoFormStuff()
    '--> this sub just sets the form up... nothing download wise happens in here.
    'Draw a black box around the virtual title bar
    Me.Line (0, 0)-(Me.ScaleWidth - 1, 32), &H0&, B
    'draw the title gradient
    DrawGradient Me, 175, 177, 166, False, 0, 0, Me.ScaleWidth - 1, 16
    DrawGradient Me, 175, 177, 166, True, 0, 17, Me.ScaleWidth - 1, 31
    DrawGradient Me, 175, 177, 166, True, 1, 34, Me.ScaleWidth - 1, Me.ScaleHeight - 1
    'Draw the form border according to the colorscheme
    Me.ForeColor = &H0
    RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
    Me.ForeColor = &HA7A7A7
    RoundRect Me.hdc, 1, 1, (Me.Width / Screen.TwipsPerPixelX) - 2, (Me.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
    '-->clip rounded corners transparent
    SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), 25, 25), True
    'get usable screen dimensions
    GetScreenInfo
    With Me
        .ForeColor = &H0
        .FontBold = True
        'I'll put the form on bottom right corner of screen... above toolbar if there
        'You can put it anywhere you like
        .Left = (ScreenDimensions.Right - Me.ScaleWidth) * Screen.TwipsPerPixelX
        .Top = (ScreenDimensions.Bottom - Me.ScaleHeight) * Screen.TwipsPerPixelY
        .CurrentX = 10
        .CurrentY = 10
    End With
    Me.Print "Downloading... Please Wait..." 'print caption here
    Label1.Caption = "Checking For Cached Copy..."
    
End Sub

