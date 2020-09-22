VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdDoMultipleDownLoads 
      Caption         =   "Multiple Files"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmDemo.frx":0000
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton cmdDoSingleDownLoad 
      Caption         =   "Single File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Debug Type Window:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Local FileName To Save To:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter URL Of File To DownLoad:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ This is just a demo form so you can see how things work... ++
'++ Sorry I didn'nt put any time into making this prettier     ++
'++ for you... NOT!                                            ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdCancel_Click()
    'so you want to cancel eh?
    frmDnLoad.EchoCancelRequest 'call the cancel request from frmDnLoad so it will receive the events
    DoEvents 'give the cancel request a chance to get in-line
    
End Sub

Private Sub cmdDoMultipleDownLoads_Click()
    'do mutliple file downloads
    '--> I'm putting stuff here... but don't jazz their servers too much!
    '--> BE NICE!!!!!!!!!!!!!
    Dim FileList As String
    FileList = "http://www.planet-source-code.com/" & "," & "c:/planet-source-code.html" & "," & "1" & ","
    FileList = FileList & "http://www.foxnews.com/" & "," & "c:/fox_news.html" & "," & "1" & ","
    FileList = FileList & "http://www.cnn.com/" & "," & "c:/cnn.html" & "," & "1" & ","
    FileList = FileList & "http://sportsillustrated.cnn.com/" & "," & "c:/SIllustrated.html" & "," & "1" & ","
    FileList = FileList & "http://www.weather.com/" & "," & "c:/weatherChannel.html" & "," & "1" '<<<<< NO TRAILING COMMA!!!!!
    
    cmdCancel.Visible = True
    Text3.Text = ""
    DoEvents
    Call frmDnLoad.ShowDownLoad(FileList, Me, Me)
End Sub

Private Sub cmdDoSingleDownLoad_Click()
    'do a single file download with form waiting for response from function
    Dim FileList As String
    FileList = Text1.Text & "," & Text2.Text & "," & "1"
    cmdCancel.Visible = True
    Text3.Text = ""
    DoEvents
    Call frmDnLoad.ShowDownLoad(FileList, Me, Me)
End Sub

Private Sub Form_Load()
    'This is a pretty weird sample mp3!  I have no clue what it's supposed to represent!
    Text1.Text = "http://www.archive.org/download/GOD_FOOTSTEPSsamplemp3/heavybeat.mp3"
    Text2.Text = "c:/God_Footsteps.mp3"
    Text3.Text = ""
End Sub


