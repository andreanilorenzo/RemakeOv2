VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' cSplitter
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

' If we used WithEvents on these items, then we not need to
' write any code at all in the form being split.  However, this
' means you have a a WithEvents reference to the form here, whilst
' the form also has a WithEvents reference to the splitter.
' This seems to cause immediate crash (VB dev environment disappears!)
' when trying to terminate the splitter class under VB5 (SP2 and above)
Private m_picSplitter() As PictureBox
Attribute m_picSplitter.VB_VarHelpID = -1
Private m_frmParent As Form
Attribute m_frmParent.VB_VarHelpID = -1

Private m_bSplitting As Boolean
Private m_lSplitOffset As Long
Private m_lBorder As Long
Private m_eOrientation As ESPLTOrientationConstants
Private m_iIndex As Integer
Private m_OffsetBasso As Integer
Private m_objOffsetBasso As Object

Public Enum ESPLTOrientationConstants
    cSPLTOrientationHorizontal = 1
    cSPLTOrientationVertical = 2
End Enum

Public Event DoSplit(bSplit As Boolean)
Public Event SplitComplete()

Property Let Orientation(eOrientation As ESPLTOrientationConstants)
    Dim cnt As Integer
    m_eOrientation = eOrientation
    
    For cnt = 0 To UBound(m_picSplitter)
        If Not (m_picSplitter(cnt) Is Nothing) Then
            If (eOrientation = cSPLTOrientationHorizontal) Then
                m_picSplitter(cnt).MousePointer = vbSizeNS
            Else
                m_picSplitter(cnt).MousePointer = vbSizeWE
            End If
        End If
    Next
    
End Property
Property Get Orientation() As ESPLTOrientationConstants
    Orientation = m_eOrientation
End Property

Property Let BorderSize(lSize As Long)
    m_lBorder = lSize
End Property
Property Get BorderSize() As Long
    BorderSize = m_lBorder
End Property

Property Let OffsetBasso(obj As Control)
    Set m_objOffsetBasso = obj
    CalcolaOffsetBasso
End Property

Private Sub CalcolaOffsetBasso()
    Dim iTop As Integer
    
    If m_objOffsetBasso Is Nothing Then Exit Sub
    
    iTop = m_objOffsetBasso.Top
    
    If iTop > (m_frmParent.ScaleHeight - m_lBorder) Then
        m_OffsetBasso = 0
    Else
        m_OffsetBasso = m_frmParent.ScaleHeight - m_objOffsetBasso.Top
    End If
    
End Sub

Public Sub Initialise(ByRef picSplitter As Object, ByRef frmParent As Form, ByVal eOrientation As ESPLTOrientationConstants, Optional ByVal lSize As Long = -1)
    Dim cnt As Long
    
    ReDim Preserve m_picSplitter(picSplitter.Count - 1)
    
    Set m_frmParent = frmParent
    
    For cnt = 0 To picSplitter.Count - 1
        Set m_picSplitter(cnt) = picSplitter(cnt)
        With m_picSplitter(cnt)
            .BorderStyle = 0
            .ZOrder 1
            .MousePointer = vbSizeWE
            .BackColor = &H80000002
            .Visible = True
        End With
    Next
    
    Orientation = eOrientation
    
    If lSize >= -1 Then
        m_lBorder = 40
    Else
        m_lBorder = lSize
    End If
    
End Sub

Public Sub MouseDown(ByVal index As Integer, X As Single, Y As Single)
    Dim Pos As Single
    Dim bSplit As Boolean
    
    bSplit = True
    RaiseEvent DoSplit(bSplit)
    If Not (bSplit) Then Exit Sub
    
    m_iIndex = index
    m_bSplitting = True
    
    With m_picSplitter(index)
        .BackColor = &H80000010
        .ZOrder 0
        .BorderStyle = 1
        If m_eOrientation = cSPLTOrientationVertical Then
            Pos = X
            .Width = 4 * Screen.TwipsPerPixelX
        ElseIf m_eOrientation = cSPLTOrientationHorizontal Then
            .Height = 4 * Screen.TwipsPerPixelY
            Pos = Y
        End If
    End With
    m_lSplitOffset = Pos
    
    SetCapture m_frmParent.hwnd
    
End Sub

Public Sub MouseMove(ByVal X As Single, ByVal Y As Single)
    Dim Pos As Single
    
    If (m_bSplitting) Then
        If (m_eOrientation = cSPLTOrientationHorizontal) Then
            ' Horizontal orientation:
            Pos = Y
            If (Pos < m_frmParent.ScaleHeight - m_lBorder - m_picSplitter(m_iIndex).Height) And (Pos > m_lBorder) Then
                Screen.MousePointer = vbSizeNS
                m_picSplitter(m_iIndex).Move m_picSplitter(m_iIndex).Left, Pos
            Else
                Screen.MousePointer = vbNoDrop
            End If
            
        Else
            ' Vertical orientation:
            Pos = X
            If (Pos < m_frmParent.ScaleWidth - (m_lBorder * 2) - m_picSplitter(m_iIndex).Width) And (Pos > m_lBorder) Then
                Screen.MousePointer = vbSizeWE
                m_picSplitter(m_iIndex).Move Pos
            Else
                Screen.MousePointer = vbNoDrop
            End If
        End If
    End If
    
End Sub

Public Function MouseUp(ByRef X As Single, ByRef Y As Single) As Boolean
    Dim Pos As Single
    Dim lRealPos As Long

    If (m_bSplitting) Then
        ' End the moving:
        ReleaseCapture
        With m_picSplitter(m_iIndex)
            .BackColor = &H80000002
            .BorderStyle = 0
            
            If (m_eOrientation = cSPLTOrientationHorizontal) Then
                Pos = Y
            Else
                Pos = X
            End If
            
        
            ' Move to a position within bounds if we are out of bounds:
            If (Pos < m_lBorder) Then
                Pos = m_lBorder + 100
                
            ElseIf (m_eOrientation = cSPLTOrientationVertical) And (Pos > (m_frmParent.ScaleWidth - m_lBorder)) Then
                Pos = m_frmParent.ScaleWidth - (m_lBorder * 2) - 100
                
            ElseIf (m_eOrientation = cSPLTOrientationHorizontal) And (Pos > (m_frmParent.ScaleHeight - m_lBorder)) Then
                Pos = m_frmParent.ScaleHeight - m_lBorder - 100
                
            End If
            
            If (m_eOrientation = cSPLTOrientationHorizontal) Then
                If (Pos > (m_frmParent.ScaleHeight - m_lBorder)) Then
                    Pos = m_frmParent.ScaleHeight - m_lBorder
                End If
            Else
                If (Pos > (m_frmParent.ScaleWidth - m_lBorder)) Then
                    Pos = m_frmParent.ScaleWidth - m_lBorder
                End If
            End If
            
            ' Now drop the splitter:
            Pos = Pos - m_lSplitOffset
            If (m_eOrientation = cSPLTOrientationHorizontal) Then
                Y = Pos
                .Move .Left, Pos
            Else
                X = Pos
                .Move Pos
            End If
            .ZOrder 1
            
        End With
        
        m_bSplitting = False
        Screen.MousePointer = vbNormal
        
        MouseUp = True
        
        RaiseEvent SplitComplete
    End If
    
End Function

Public Sub Resize(Optional CtlAltoSX As Object, Optional CtlBassoSX As Object, Optional CtlDestra As Object, Optional ByVal index As Integer = 0, Optional ByRef CtlOpt1 As Object, Optional ByRef CtlOpt2 As Object)
    Dim lM As Long
    Dim lW As Long
    Dim ll As Long
    Dim lH As Long
    Dim Bordo As Long
    
    On Error Resume Next
    
    Bordo = m_lBorder
    CalcolaOffsetBasso
    
    ' Controllo che la m_picSplitter non sia stata posizionata fuori dai margini
    If (Not CtlAltoSX Is Nothing) Then
        If m_picSplitter(index).Top < CtlAltoSX(0).Top Then
            m_picSplitter(index).Move m_picSplitter(index).Left, CtlAltoSX(0).Top
        End If
        If UBound(m_picSplitter) = 1 Then
            If m_picSplitter(1).Top < CtlAltoSX(0).Top Then
                m_picSplitter(1).Move m_picSplitter(1).Left, CtlAltoSX(0).Top
            End If
        End If
    End If
    
    
    If (CtlBassoSX Is Nothing) And (m_eOrientation = cSPLTOrientationVertical) Then
        With CtlAltoSX(0)
           ' Calcolo l'altezza
           lH = m_frmParent.ScaleHeight - .Top - (Bordo * 1.6) - m_OffsetBasso
           ' Calcolo la larghezza
           lM = m_picSplitter(index).Left - 2 * Screen.TwipsPerPixelX
           .Move Bordo, .Top, lM, lH
           ll = lM + m_picSplitter(index).Width + 2 * Screen.TwipsPerPixelX
           CtlDestra(0).Move ll, .Top, m_frmParent.ScaleWidth - ll - 4 * Screen.TwipsPerPixelX, lH
        End With
        With m_picSplitter(index)
           .Move .Left, CtlAltoSX(0).Top, .Width, lH
           .ZOrder
        End With
             
             
             
    ElseIf (Not CtlBassoSX Is Nothing) And (m_eOrientation = cSPLTOrientationVertical) Then
        If (GetNomeObj(CtlAltoSX) <> "picSplitH") Then
            With CtlAltoSX(0)
               'lH = Me.ScaleHeight - .Top * 2 '- sbrMain.Height * -1 * sbrMain.Visible
               ' Calcolo l'altezza
               lH = CtlBassoSX(0).Top - .Top
               ' Calcolo la larghezza
               lM = m_picSplitter(index).Left
               .Move Bordo, .Top, lM, lH
               ll = lM + m_picSplitter(index).Width
               CtlDestra(0).Move ll, .Top, m_frmParent.ScaleWidth - ll - Bordo, lH
            End With
        
            With m_picSplitter(index)
               .Move .Left, CtlAltoSX(0).Top, .Width, lH
               .ZOrder
            End With
            
        Else
            With CtlBassoSX(0)
                ' Calcolo l'altezza
                lH = m_frmParent.ScaleHeight - CtlAltoSX(0).Top - CtlAltoSX(0).Height - Bordo
                ' Calcolo la larghezza
                lM = m_picSplitter(index).Left
                .Move Bordo, CtlAltoSX(0).Top + CtlAltoSX(0).Height, lM, lH
            End With
            
            With m_picSplitter(index)
               .Move .Left, CtlBassoSX(0).Top, .Width, CtlBassoSX(0).Height
               .ZOrder
            End With
            
            With CtlDestra(0)
                ' Calcolo l'altezza
                lH = CtlBassoSX(0).Height
                ' Calcolo la larghezza
                lM = m_frmParent.ScaleWidth - m_picSplitter(index).Left - m_picSplitter(index).Width - Bordo
                .Move m_picSplitter(index).Left + m_picSplitter(index).Width, CtlBassoSX(0).Top, lM, lH
            End With
            
            
        End If



    ElseIf (GetNomeObj(CtlBassoSX) <> "picSplitH") And (Not CtlBassoSX Is Nothing) And (m_eOrientation = cSPLTOrientationHorizontal) Then
        If (CtlDestra Is Nothing) And (Not CtlAltoSX Is Nothing) Then
            With CtlAltoSX(0)
               ' Calcolo l'altezza
               lH = (m_frmParent.ScaleHeight - (m_frmParent.ScaleHeight - m_picSplitter(index).Top)) - .Top
               ' Calcolo la larghezza
               lM = m_frmParent.ScaleWidth - (Bordo * 2)
               .Move Bordo, .Top, lM, lH
            End With
        End If
        
        If (Not CtlDestra Is Nothing) Then
            If (GetNomeObj(CtlDestra) = "picSplitV") Then
                With CtlBassoSX(0)
                   lH = m_frmParent.ScaleHeight - m_picSplitter(index).Top - Bordo
                   lM = m_picSplitter(index).Top + m_picSplitter(index).Height
                   
                   .Move Bordo, lM, m_frmParent.ScaleWidth - (m_frmParent.ScaleWidth - CtlDestra(0).Left) - Bordo, lH - (Bordo * 1.6)
                End With
                
                With CtlAltoSX(0)
                    lH = m_picSplitter(index).Top - .Top
                    .Move CtlBassoSX(0).Left, .Top, CtlBassoSX(0).Width, lH
                End With
                
                With CtlDestra(0)
                    .Move .Left, CtlAltoSX(0).Top, .Width, CtlAltoSX(0).Height + CtlBassoSX(0).Height + m_picSplitter(index).Height
                    .ZOrder
                End With
                
                If index = 0 And UBound(m_picSplitter) = 1 Then
                    With m_picSplitter(1)
                        .Move CtlDestra(0).Left + CtlDestra(0).Width, .Top, m_frmParent.ScaleWidth - CtlDestra(0).Left - CtlDestra(0).Width - (Bordo * 2), .Height
                        .ZOrder
                    End With
                End If
                    
            End If
        Else
            With CtlBassoSX(0)
               lH = m_frmParent.ScaleHeight - m_picSplitter(index).Top - Bordo - m_OffsetBasso
               lM = m_picSplitter(index).Top + m_picSplitter(index).Height
               .Move Bordo, lM, m_frmParent.ScaleWidth - (Bordo * 2), lH - (Bordo * 1.6)
            End With
        End If
        
        With m_picSplitter(index)
           .Move CtlBassoSX(0).Left, CtlBassoSX(0).Top - m_picSplitter(index).Height, CtlBassoSX(0).Width, .Height
           .ZOrder
        End With
        
        If index = 0 And UBound(m_picSplitter) = 1 Then
            If (Not CtlOpt1 Is Nothing) Then
                With CtlOpt1(0)
                    lW = m_picSplitter(1).Width
                    lH = m_frmParent.ScaleHeight - (m_frmParent.ScaleHeight - m_picSplitter(1).Top) - CtlDestra(0).Top
                    .Move CtlDestra(0).Left + CtlDestra(0).Width, CtlAltoSX(0).Top, lW, lH
                End With
            End If
            
            If (Not CtlOpt2 Is Nothing) Then
                With CtlOpt2(0)
                    lW = m_picSplitter(1).Width
                    lH = (m_frmParent.ScaleHeight - m_picSplitter(1).Top) - m_picSplitter(1).Height - Bordo
                    .Move CtlDestra(0).Left + CtlDestra(0).Width, m_picSplitter(1).Top + m_picSplitter(1).Height, lW, lH
                End With
            End If
        End If


    ElseIf (GetNomeObj(CtlBassoSX) = "picSplitH") And (Not CtlBassoSX Is Nothing) And (m_eOrientation = cSPLTOrientationHorizontal) Then
        If (CtlDestra Is Nothing) And (Not CtlAltoSX Is Nothing) Then
            With CtlAltoSX(0)
               ' Calcolo l'altezza
               lH = (m_frmParent.ScaleHeight - (m_frmParent.ScaleHeight - m_picSplitter(index).Top)) - .Top
               ' Calcolo la larghezza
               lM = m_frmParent.ScaleWidth - (Bordo * 2)
               .Move Bordo, .Top, lM, lH
            End With
        End If
        
        With m_picSplitter(index)
           .Move CtlAltoSX(0).Left, .Top, CtlAltoSX(0).Width
           .ZOrder
        End With
        
    End If
    
End Sub

Private Function GetNomeObj(ByVal obj As Object, Optional ByVal index As Integer = 0) As String
    On Error Resume Next
    GetNomeObj = Trim$(obj(index).Name)
End Function

Private Sub Class_Initialize()
    m_eOrientation = cSPLTOrientationVertical
    m_OffsetBasso = 0
End Sub

Private Sub Class_Terminate()
    Dim cnt As Integer
    
    m_bSplitting = False
    Set m_frmParent = Nothing
    
    For cnt = 0 To UBound(m_picSplitter)
        Set m_picSplitter(cnt) = Nothing
    Next
    
End Sub

