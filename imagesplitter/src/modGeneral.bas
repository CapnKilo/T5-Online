Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public FileName() As String
Public Extension As String
Public TotalSpritesSaved As Long

Private SizeX As Long
Private SizeY As Long

Public Sub AddToLog(Msg As String)
Static NumLines As Long

    NumLines = NumLines + 1
    
    If NumLines >= 500 Then
        frmMain.txtEvents.Text = vbNullString
        NumLines = 0
    End If
    
    frmMain.txtEvents.Text = frmMain.txtEvents.Text & vbNewLine & Msg
    frmMain.txtEvents.SelStart = Len(frmMain.txtEvents.Text)
    
End Sub

Sub UpdateProperties()
    frmMain.lblWidth.Caption = "Width: " & frmMain.picImage.ScaleWidth
    frmMain.lblHeight.Caption = "Height: " & frmMain.picImage.ScaleHeight
    frmMain.lblSprites.Caption = "Sprites: " & frmMain.picImage.ScaleHeight / Val(frmMain.txtSizeY.Text)
    If TotalSpritesSaved = 0 Then TotalSpritesSaved = Val(frmMain.txtStartAt.Text)
End Sub

Sub SplitImage()
Dim i As Long

    frmMain.picConversion.Width = frmMain.picImage.Width
    frmMain.picConversion.Height = Val(frmMain.txtSizeY) * Screen.TwipsPerPixelY
    
    frmMain.Height = 4635
    
    DoEvents
    
    For i = 0 To ((frmMain.picImage.ScaleHeight) / Val(frmMain.txtSizeY)) - 1
        BitBlt frmMain.picConversion.hDC, 0, 0, frmMain.picImage.ScaleWidth, Val(frmMain.txtSizeY), frmMain.picImage.hDC, 0, (i * Val(frmMain.txtSizeY)), vbSrcCopy
            
        SavePicture frmMain.picConversion.Image, App.Path & "\converted\" & TotalSpritesSaved & Extension
        AddToLog TotalSpritesSaved & Extension & " saved."
        
        TotalSpritesSaved = TotalSpritesSaved + 1
        frmMain.txtStartAt.Text = TotalSpritesSaved
        
        frmMain.picConversion.Cls
    Next
    
End Sub
