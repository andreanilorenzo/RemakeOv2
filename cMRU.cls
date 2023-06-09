VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     cMRU
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 January 2000
'
' Copyright � 2000 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' A simple MRU list which persists its state to the
' registry.
'
' ======================================================================================

Private m_iMaxCount As Long
Private m_iCount As Long
Private m_sItems() As String

Public Property Get MaxCount() As Long
   MaxCount = m_iMaxCount
End Property
Public Property Let MaxCount(ByVal lMaxCount As Long)
   m_iMaxCount = lMaxCount
   If m_iCount > m_iMaxCount Then
      m_iCount = m_iMaxCount
   End If
End Property

Public Property Get Count() As Long
   Count = m_iCount
End Property

Public Property Get Item(ByVal nIndex As Long)
   Item = m_sItems(nIndex)
End Property

Public Sub Add(ByVal sItem As String)
Dim i As Long
Dim iIdx As Long
   
   For i = 1 To m_iCount
      If m_sItems(i) = sItem Then
         iIdx = i
         Exit For
      End If
   Next i
   
   If iIdx > 0 Then
      ' swap from 1 -> iIdx-1
      For i = iIdx - 1 To 1 Step -1
         m_sItems(i + 1) = m_sItems(i)
      Next i
   Else
      If m_iCount < m_iMaxCount Then
         m_iCount = m_iCount + 1
         ReDim Preserve m_sItems(1 To m_iCount) As String
      End If
      For i = m_iCount - 1 To 1 Step -1
         m_sItems(i + 1) = m_sItems(i)
      Next i
   End If
   m_sItems(1) = sItem
   
End Sub

Private Sub Class_Initialize()
   m_iMaxCount = 16
End Sub
