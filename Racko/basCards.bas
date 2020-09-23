Attribute VB_Name = "basCards"
'Racko
'Paul Bahlawan
'Jan 2011
'
'-------------
'basCards
'For drawing cards

Option Explicit

'Card styles
Public Enum cStyle
    cBack
    cBackHighlight
    cFront
    cFrontHighlight
    cNone
End Enum

Public Declare Function GdiTransparentBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean


Public Sub DrawDeck(Style As cStyle, Optional Value As Long)
    Select Case Style
        Case cBack
            GdiTransparentBlt frmMain.hDC, 220, 60, 195, 120, frmMain.picBack.hDC, 0, 0, 195, 120, vbMagenta
        
        Case cBackHighlight
            GdiTransparentBlt frmMain.hDC, 220, 60, 195, 120, frmMain.picBackHighlight.hDC, 0, 0, 195, 120, vbMagenta
        
        Case cFrontHighlight
            GdiTransparentBlt frmMain.hDC, 220, 60, 195, 120, frmMain.picFrontHighlight.hDC, 0, 0, 195, 120, vbMagenta
            frmMain.CurrentX = 222 + 2 * Value
            frmMain.CurrentY = 60
            frmMain.Print Value
        
        Case cNone
            GdiTransparentBlt frmMain.hDC, 220, 60, 195, 120, frmMain.hDC, 220, 0, 195, 50, vbMagenta
            frmMain.Line (230, 70)-(405, 170), vbBlack, B
    End Select
End Sub

Public Sub DrawDiscard(Style As cStyle, Value As Long)
    Select Case Style
        Case cFront
            GdiTransparentBlt frmMain.hDC, 220, 200, 195, 120, frmMain.picFront.hDC, 0, 0, 195, 120, vbMagenta
        
        Case cFrontHighlight
            GdiTransparentBlt frmMain.hDC, 220, 200, 195, 120, frmMain.picFrontHighlight.hDC, 0, 0, 195, 120, vbMagenta
    End Select
    
    frmMain.CurrentX = 222 + 2 * Value
    frmMain.CurrentY = 200
    frmMain.Print Value
End Sub

'For drawing a card in a rack
Public Sub DrawCard(player As Long, yPos As Long, Style As cStyle, Value As Long)
Dim x As Long
    If player = 1 Then
        x = 20
    Else
        x = 20 + 200 * player
    End If
    
    Select Case Style
        Case cBack
            GdiTransparentBlt frmMain.hDC, x, 30 * yPos, 195, 120, frmMain.picBack.hDC, 0, 0, 195, 120, vbMagenta
        
        Case cBackHighlight
            GdiTransparentBlt frmMain.hDC, x, 30 * yPos, 195, 120, frmMain.picBackHighlight.hDC, 0, 0, 195, 120, vbMagenta
        
        Case cFront
            GdiTransparentBlt frmMain.hDC, x, 30 * yPos, 195, 120, frmMain.picFront.hDC, 0, 0, 195, 120, vbMagenta
            frmMain.CurrentX = x + 2 * Value
            frmMain.CurrentY = 30 * yPos
            frmMain.Print Value
        
        Case cFrontHighlight
            GdiTransparentBlt frmMain.hDC, x, 30 * yPos, 195, 120, frmMain.picFrontHighlight.hDC, 0, 0, 195, 120, vbMagenta
            frmMain.CurrentX = x + 2 * Value
            frmMain.CurrentY = 30 * yPos
            frmMain.Print Value
    End Select

End Sub

'Draw all cards in a rack
' if ySel > 0  card will be highlighted
Public Sub DrawRack(player As Long, ySel As Long, Value() As Long)
Dim i As Long
    For i = 1 To 10
        If player = 1 Then
            If i = ySel Then
                DrawCard player, i, cFrontHighlight, Value(player, i)
            Else
                DrawCard player, i, cFront, Value(player, i)
            End If
        Else
            If i = ySel Then
                DrawCard player, i, cBackHighlight, Value(player, i)
            Else
                DrawCard player, i, cBack, Value(player, i)
            End If
        End If
    Next i
End Sub

'Replace deck and discard pile with "RACKO"
Public Sub DrawRacko()
Dim i As Long
    GdiTransparentBlt frmMain.hDC, 220, 60, 195, 400, frmMain.hDC, 220, 0, 195, 50, vbMagenta
    frmMain.FontSize = 72
    For i = 1 To 5
        frmMain.ForeColor = &H404040
        frmMain.CurrentX = 279
        frmMain.CurrentY = 80 * i - 82
        frmMain.Print Mid$("RACKO", i, 1)
        frmMain.ForeColor = vbRed
        frmMain.CurrentX = 275
        frmMain.CurrentY = 80 * i - 86
        frmMain.Print Mid$("RACKO", i, 1)
    Next i
End Sub
