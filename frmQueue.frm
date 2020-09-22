VERSION 5.00
Begin VB.Form frmQueue 
   Caption         =   "Queue Example"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   300
      Left            =   3225
      TabIndex        =   4
      Top             =   1170
      Width           =   1605
   End
   Begin VB.CommandButton cmdDequeue 
      Caption         =   "&Dequeue"
      Height          =   300
      Left            =   3225
      TabIndex        =   3
      Top             =   735
      Width           =   1605
   End
   Begin VB.CommandButton cmdEnqueue 
      Caption         =   "&Enqueue"
      Height          =   300
      Left            =   3225
      TabIndex        =   2
      Top             =   285
      Width           =   1605
   End
   Begin VB.TextBox txtStuff 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   1140
      Width           =   2400
   End
   Begin VB.TextBox txtID 
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   2400
   End
   Begin VB.Label lblStatus 
      Caption         =   "Items Available: 0"
      Height          =   330
      Left            =   30
      TabIndex        =   7
      Top             =   1650
      Width           =   4950
   End
   Begin VB.Label Label1 
      Caption         =   "Other Info"
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   885
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Type ID"
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   2400
   End
End
Attribute VB_Name = "frmQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' This form demonstrates the use of the CQueue Class
' Created By: AAA (aaa_001@hotmail.com)
'==========================================================

Private Type QUEUE_TEST ' Just a random type, this can be any fixed type of
    lID As Long         ' data like arrays or even VB's standard data types
    lStuff As Long      ' except for Variants & Dynamic Strings
End Type

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&

Private DaQueue As CQueue           ' Our Queue
Private DaTestStruct As QUEUE_TEST  ' Test structure we will use in the queue

Private Sub SetNumberText(ByVal hWnd As Long)
    '======================================================
    ' A better way to filter out non-numeric characters.
    '======================================================
    SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Or ES_NUMBER
End Sub

Private Sub cmdClear_Click()
    '======================================================
    ' This simply clears the queue of all its contents
    '======================================================
    DaQueue.Clear
    lblStatus.Caption = "Items Available: " & CStr(DaQueue.Count)
End Sub

Private Sub cmdDequeue_Click()
    '======================================================
    ' Reads the next item in the queue
    '======================================================
    If DaQueue.Dequeue(VarPtr(DaTestStruct)) Then
    
        ' Output the result in the textboxes
        txtID.Text = CStr(DaTestStruct.lID)
        txtStuff.Text = CStr(DaTestStruct.lStuff)
        lblStatus.Caption = "Items Available: " & CStr(DaQueue.Count)
    Else
        MsgBox "The queue is empty!"
    End If
End Sub

Private Sub cmdEnqueue_Click()
    '======================================================
    ' Add an item to the end of the queue
    '======================================================
    DaTestStruct.lID = CLng(txtID.Text)
    DaTestStruct.lStuff = CLng(txtStuff.Text)
    
    ' Add the item!
    DaQueue.Enqueue VarPtr(DaTestStruct)
    lblStatus.Caption = "Items Available: " & CStr(DaQueue.Count)
End Sub

Private Sub Form_Load()
    '======================================================
    ' Start up!
    '======================================================
    SetNumberText txtID.hWnd
    SetNumberText txtStuff.hWnd
 
    Set DaQueue = New CQueue    ' Create the object
    
    ' We must initialize it with the size of our structure
    DaQueue.Initialize Len(DaTestStruct)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '======================================================
    ' End the program
    '======================================================
    Set DaQueue = Nothing   ' Dereference the queue, we must do this to clean
                            ' up after ourselves
End Sub
