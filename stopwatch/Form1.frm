VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "High Performance StopWatch Demo"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3495
      TabIndex        =   8
      Top             =   1980
      Width           =   1245
   End
   Begin VB.CommandButton CmdSplit 
      Caption         =   "Split"
      Height          =   450
      Left            =   2475
      TabIndex        =   6
      Top             =   1830
      Width           =   930
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   450
      Left            =   2475
      TabIndex        =   5
      Top             =   1380
      Width           =   930
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   255
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1710
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   690
      Left            =   1185
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   3645
   End
   Begin VB.CommandButton CmdInit 
      Caption         =   "Initialize"
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   840
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      Height          =   210
      Left            =   1320
      TabIndex        =   10
      Top             =   1530
      Width           =   840
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   1305
      TabIndex        =   9
      Top             =   1755
      Width           =   945
   End
   Begin VB.Label Label4 
      Caption         =   "Elapsed time (s):"
      Height          =   210
      Left            =   3495
      TabIndex        =   7
      Top             =   1725
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Select a StopWatch:"
      Height          =   420
      Left            =   210
      TabIndex        =   4
      Top             =   1275
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Valid StopWatch Handles:"
      Height          =   270
      Left            =   1230
      TabIndex        =   2
      Top             =   255
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdInit_Click()
Dim rethandle As Long

'get a handle to an initialized timer
rethandle = StopWatchInitialize

Text1.Text = Text1.Text & "  " & rethandle

'add handle to list
Combo1.AddItem rethandle, rethandle - 1
Combo1.ListIndex = rethandle - 1

End Sub

Private Sub CmdStart_Click()
Dim retval As Long

'start selected stopwatch
retval = StopWatchStart(Combo1.Text)

If retval = 1 Then
'successful
Else
'not (not initialized?)
End If

Call Combo1_Click

End Sub


Private Sub CmdSplit_Click()
Dim retval As Single

'Get elapsed time
retval = StopWatchSplit(Combo1.Text)

If retval > 0 Then
    'successful
    Text2.Text = retval
ElseIf retval = 0 Then
    Text2.Text = "not started"
ElseIf retval = -1 Then
    'not initialized
End If

End Sub

Private Sub Combo1_Click()
Dim retval As Long

retval = GetStopWatchStatus(Combo1.Text)

If retval = 1 Then
    Label5.Caption = "started"
ElseIf retval = 0 Then
    Label5.Caption = "stopped"
Else
    Label5.Caption = "Invalid"
End If

End Sub

Private Sub form_load()

'Initialize a stopwatch to start with
Call CmdInit_Click

End Sub
