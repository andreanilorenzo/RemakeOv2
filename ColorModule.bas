Attribute VB_Name = "ColorModule"
Option Explicit

Public Function ConvertDelphiColor(ByVal sDelphiColor As String) As Variant
    'clBlack, clMaroon, clGreen, clOlive, clNavy,
    'clPurple, clTeal, clGray, clSilver, clRed, clLime,
    'clYellow, clBlue, clFuchsia, clAqua, clLtGray,
    'clDkGray, clWhite, clScrollBar, clBackground,
    'clActiveCaption, clInactiveCaption, clMenu, clWindow,
    'clWindowFrame, clMenuText, clWindowText, clCaptionText,
    'clActiveBorder, clInactiveBorder, clAppWorkSpace,
    'clHighlight, clHighlightText, clBtnFace, clBtnShadow,
    'clGrayText, clBtnText, clInactiveCaptionText,
    'clBtnHighlight, cl3DDkShadow, cl3DLight, clInfoText,
    'clInfoBk
    Dim cVbColor As ColorConstants
    Dim cSiSolor As SystemColorConstants
    
    ConvertDelphiColor = vbGrayText
    sDelphiColor = LCase(sDelphiColor)
    
    Select Case sDelphiColor
        Case LCase("clBlack")
            ConvertDelphiColor = vbBlack
    
        Case LCase("clGreen")
            ConvertDelphiColor = vbGreen
    
        Case LCase("clInfoBk")
            ConvertDelphiColor = vbBlack
        
        Case LCase("clBtnFace")
            ConvertDelphiColor = vbButtonFace
        
    End Select

End Function


