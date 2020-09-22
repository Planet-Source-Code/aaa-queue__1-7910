VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' CQueue -> A queue data structure, in the form of an
'           object.
'
'   This class internally maintains all the data passed to
' it. Use this class for fixed data types (eg. Longs, fixed
' arrays, fixed strings, types, etc...)
'
'   This implementation of a queue is somewhat of a hybrid
' between a linked list and a queue. In it's core it can
' be said that the object behaves more like a linked list
' but the interface to it is of a Queue. Therefore, you can
' easily modify this code to operate like a linked list.
'
' =====================================
' Created By: AAA (aaa_001@hotmail.com)
' =====================================
'
'   This code is published under the public domain. You may
' freely use it, modify it... however you may not call it
' your own. Please give credit where credit IS due. If you
' intend on using this code please be kind and give me some
' credit for it.
'
' USAGE:
'  1) Create an instance of the object
'  2) Call Initialize and pass the class the size of a
'     single element in your type structure, or whatever
'     data you intend to use.
'  3) Use Enqueue with a pointer to the data you wish to
'     store to add an item to the queue.
'  4) Use Dequeue with a pointer to a buffer to retrieve
'     the next item in the queue. It also removes the item
'     from the queue.
'  5) When done using the class you _MUST_ dereference the
'     Object!
'
' BUGS:
'  1) The Queue can only hold ~32767 items due to a Windows
'     limit. Put some sort of error check for this just in
'     case.
'
' REVISION HISTORY:
'  1) 05/05/2000 -> Created class
'
'==========================================================
' Conditional compilation constants for enhanced functionality.
' Turn these options on if you require them in the program,
' otherwise turn them off to reduce program size.
'
' 1 = On    0 = Off
'==========================================================
#Const OPTION_COUNT = 1 ' Enable code to return the number of items in the queue
#Const OPTION_CLEAR = 1 ' Enable code to clear the queue. If this is turned off
                        ' you must make sure the queue is empty before exiting
                        ' the program or dereferencing the object. For the sake
                        ' of simplicity, leave this option turned on!

Private Type QUEUE_ITEM
    lpNext As Long  ' Pointer to next item in the queue
    lpData As Long  ' Pointer to the data
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (dest As Any, src As Any, ByVal numBytes As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (dest As Any, ByVal numBytes As Long)
Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)

Private Const ERR_DESCRIPTION_INIT = "Class is not properly initialized!"
Private Const ERR_DESCRIPTION_ALREADYINIT = "Class is already initialized!"
Private Const ERR_SOURCE = "CQueue Class"

Private m_lSize As Long         ' The size of each item of the queue
Private m_lHead As QUEUE_ITEM   ' First item in the queue

#If OPTION_COUNT Then
Private m_lCount As Long        ' Number of items in the queue
#End If ' OPTION_COUNT

Public Sub Initialize(ByVal lSize As Long)
    '======================================================
    ' Initialize -> Sets up the size of a single queue item
    '               so that the queue can be used. This sub
    '               _MUST_ be called prior to calling any
    '               other methods.
    '======================================================
    If m_lSize = 0 Then
        ' Only initialize once.
        m_lSize = lSize ' Set the size (bytes) of a single item
                        ' in the queue
        #If OPTION_COUNT Then
            m_lCount = 0
        #End If
    Else
        Err.Raise 55554, ERR_SOURCE, ERR_DESCRIPTION_ALREADYINIT
    End If
End Sub

#If OPTION_CLEAR Then

Public Sub Clear()
    '======================================================
    ' Clear -> Frees all memory used by the queue. Resets
    '          the queue to empty state.
    '======================================================
    Dim qItem As QUEUE_ITEM
    Dim lTemp As Long
    
    If m_lSize > 0 Then ' Have we initialized?
    
        ' Start at the head
        RtlMoveMemory qItem, m_lHead, Len(m_lHead)
        
        Do While qItem.lpData > 0
            LocalFree qItem.lpData      ' Free memory
            qItem.lpData = 0
            
            ' Retrive the pointer to the next item
            lTemp = qItem.lpNext
            
            If lTemp > 0 Then
                ' Get a copy of the next item in the list
                RtlMoveMemory qItem, ByVal lTemp, Len(m_lHead)
                
                ' Free windows' copy of the next item
                LocalFree lTemp
            End If
        Loop
        
        ' Clear the head
        m_lHead.lpData = 0
        m_lHead.lpNext = 0
        
        #If OPTION_COUNT Then
            m_lCount = 0
        #End If
    Else
        Err.Raise 55555, ERR_SOURCE, ERR_DESCRIPTION_INIT
    End If
End Sub

#End If ' OPTION_CLEAR

Public Sub Enqueue(ByVal lpData As Long)
    '======================================================
    ' Enqueue -> Add an Item to the end of the queue.
    '
    '   Inputs:
    '       lpData -> pointer to the data you want to add
    '                 to the queue. (Use VarPtr)
    '======================================================
    Dim qItem As QUEUE_ITEM
    Dim qLastItem As QUEUE_ITEM
    Dim lTempPtr As Long
        
    If m_lSize > 0 Then ' Have we initialized?
                
        ' Create the buffer space for the data to copy
        qItem.lpData = LocalAlloc(LPTR, m_lSize)

        ' Save the data
        RtlMoveMemory ByVal qItem.lpData, ByVal lpData, m_lSize
        
        If m_lHead.lpData = 0 Then
            ' The queue is empty right now, add new item to the head
            RtlMoveMemory m_lHead, qItem, Len(m_lHead)
            
        Else
            ' Find the last item in the queue and place the new item
            ' after it.
            lTempPtr = VarPtr(m_lHead)
            
            ' Copy the first item
            RtlMoveMemory qLastItem, ByVal lTempPtr, Len(m_lHead)
            
            ' Get to the last item
            Do While qLastItem.lpNext > 0
                lTempPtr = qLastItem.lpNext
                RtlMoveMemory qLastItem, ByVal lTempPtr, Len(m_lHead)
            Loop
            
            ' Allocate the memory for new queue item
            qLastItem.lpNext = LocalAlloc(LPTR, Len(m_lHead))
            
            ' Finally copy back the info
            RtlMoveMemory ByVal qLastItem.lpNext, qItem, Len(m_lHead)
            RtlMoveMemory ByVal lTempPtr, qLastItem, Len(m_lHead)
        End If
        
        #If OPTION_COUNT Then
            m_lCount = m_lCount + 1 ' Increment counter
        #End If
        
    Else
        Err.Raise 55555, ERR_SOURCE, ERR_DESCRIPTION_INIT
    End If
End Sub

Public Function Dequeue(ByVal lpData As Long) As Boolean
    '======================================================
    ' Dequeue -> Read & remove the first item in the queue
    '            and move the rest of the items up.
    '
    '   Inputs:
    '       lpData -> pointer to a buffer to receive the
    '                 info.
    '
    '   Output:
    '       True  -> If there is an item to retrieve.
    '       False -> The queue is already empty. Contents
    '                of lpData are cleared as well.
    '======================================================
    Dim lTemp As Long
    
    If m_lSize > 0 Then ' Have we initialized?
    
        If m_lHead.lpData > 0 Then
            ' There's at least one item in the list.
            ' copy it to the pointer supplied by lpData
            RtlMoveMemory ByVal lpData, ByVal m_lHead.lpData, m_lSize
            
            ' Clean up a bit
            LocalFree m_lHead.lpData
            
            ' Move the next item up (if one exists) to the head
            ' of the queue
            lTemp = m_lHead.lpNext
            If lTemp > 0 Then
                RtlMoveMemory m_lHead, ByVal lTemp, Len(m_lHead)
                LocalFree lTemp
            Else
                m_lHead.lpData = 0
                m_lHead.lpNext = 0
            End If
            
            #If OPTION_COUNT Then
                m_lCount = m_lCount - 1 ' Decrement counter
            #End If
            
            Dequeue = True
        Else
            ' The queue is empty, clear out the lpData structure
            RtlZeroMemory ByVal lpData, Len(m_lHead)
            Dequeue = False
        End If
    Else
        Err.Raise 55555, ERR_SOURCE, ERR_DESCRIPTION_INIT
    End If
End Function

#If OPTION_COUNT Then

Public Property Get Count() As Long
    '======================================================
    ' Count -> Retrieve the number of items in the queue.
    '          Read only property
    '======================================================
    Count = m_lCount
End Property

#End If ' OPTION_COUNT

Private Sub Class_Terminate()
    '======================================================
    ' Clean up the queue
    '======================================================
    #If OPTION_CLEAR Then
        Call Clear
    #End If
End Sub
