Attribute VB_Name = "modStar"
' Name          : modStar
' Description   : Module for drawing the star
' Lines of Code : 187
'
' Modified      : 10/10/2001
'
' --------------------------------------------------
Option Explicit

    'Declare PI & Rad
    Const m_cPI = 3.141592654
    Const m_cRAD = m_cPI / 180

Public Sub GeoStar(frm As Form, pbx1 As PictureBox, lngLength As Long, lngWidth As Long, dblSides As Double, lngDensity As Long, lngFTR As String, lngFTG As String, lngFTB As String, blnShadow As Boolean, strShadowPos As String, intLineSize As Integer, intShadowSize As Integer, lngSR As Long, lngSG As Long, lngSB As Long)

    'Declare Variables
    Dim dblX            As Double
    Dim dblX1           As Double
    Dim dblX2           As Double
    Dim dblY1           As Double
    Dim dblY2           As Double
    Dim dblB1           As Double
    Dim dblB2           As Double
    Dim dblY            As Double
    Dim dblAngle        As Double
    Dim dblMidX         As Double
    Dim dblMidY         As Double
    Dim intShadowX      As Integer
    Dim intShadowY      As Integer
    Dim lngFR           As Long
    Dim lngFG           As Long
    Dim lngFB           As Long
    
    'Error Handler
    On Error GoTo PROC_ERR
    
    'Convert Colours
    If LCase(lngFTR) <> "lr" And LCase(lngFTR) <> "sr" Then lngFR = CLng(Val(lngFTR))
    If LCase(lngFTG) <> "lr" And LCase(lngFTG) <> "sr" Then lngFG = CLng(Val(lngFTG))
    If LCase(lngFTB) <> "lr" And LCase(lngFTB) <> "sr" Then lngFB = CLng(Val(lngFTB))
    
    'Determine shadow co-ordinates
    Select Case LCase(strShadowPos)
    
        Case "t" 'Shadow is above the star
            intShadowX = 0
            intShadowY = -15 * intShadowSize
        
        Case "b" 'Shadow is below the star
            intShadowX = 0
            intShadowY = 15 * intShadowSize
            
        Case "l" 'Shadow is to the left of the star
            intShadowX = -15 * intShadowSize
            intShadowY = 0
            
        Case "r" 'Shadow is to the right of the star
            intShadowX = 15 * intShadowSize
            intShadowY = 0
            
        Case "tl" 'Shadow is to the top-left of the star
            intShadowX = -15 * intShadowSize
            intShadowY = -15 * intShadowSize
            
        Case "tr" 'Shadow is to the top-right of the star
            intShadowX = 15 * intShadowSize
            intShadowY = -15 * intShadowSize
            
        Case "bl" 'Shadow is to the bottom-left of the star
            intShadowX = -15 * intShadowSize
            intShadowY = 15 * intShadowSize
            
        Case "br" 'Shadow is to the bottom-right of the star (Default)
            intShadowX = 15 * intShadowSize
            intShadowY = 15 * intShadowSize
            
        Case Else 'Make the shadow in the bottom-right
            intShadowX = 15 * intShadowSize
            intShadowY = 15 * intShadowSize
            
    End Select
    
    'Clear Screen
    pbx1.Cls
    
    'Change # of sides to an angle
    dblAngle = 360 / dblSides
    
    'Find middle of PictureBox
    dblMidX = frm.pbx1.Width / 2
    dblMidY = frm.pbx1.Height / 2
    
    'Draw Shadow
    If blnShadow Then
        
        pbx1.DrawWidth = intShadowSize
        
        'See which shape to draw
        For dblY = 0 To 360 Step (360 / lngDensity)
            
            'Get coordinates
            dblB1 = lngWidth * Cos(m_cRAD * dblY) + dblMidY + intShadowY
            dblB2 = lngWidth * Sin(m_cRAD * dblY) + dblMidX + intShadowX
            dblX2 = dblB2
            dblY2 = dblB1
            
            'Draw shape
            For dblX = 0 To 360 / dblAngle Step 1
                
                'Get coordinates
                dblX1 = dblX2
                dblY1 = dblY2
                
                dblX2 = lngLength * Cos(m_cRAD * (dblAngle * dblX - dblY)) + dblB2
                dblY2 = lngLength * Sin(m_cRAD * (dblAngle * dblX - dblY)) + dblB1
                
                'Draw Line
                If dblX <> 0 Then
                    frm.pbx1.Line (dblX1, dblY1)-(dblX2, dblY2), RGB(lngSR, lngSG, lngSB)
                End If
                
            Next dblX
            
        Next dblY
        
    End If
    
    'Draw actual star
    For dblY = 0 To 360 Step (360 / lngDensity)
        
        pbx1.DrawWidth = intLineSize
        
        'Get coordinates
        dblB1 = lngWidth * Cos(m_cRAD * dblY) + dblMidY
        dblB2 = lngWidth * Sin(m_cRAD * dblY) + dblMidX
        dblX2 = dblB2
        dblY2 = dblB1
        
            'Check to see if random colours should be used or not
            Randomize
            If lngFTR = "sr" Then lngFR = CLng(255 * Rnd)
            
            Randomize
            If lngFTG = "sr" Then lngFG = CLng(255 * Rnd)
            
            Randomize
            If lngFTB = "sr" Then lngFB = CLng(255 * Rnd)
        
        'Draw Shape
        For dblX = 0 To 360 / dblAngle Step 1
            
            Randomize
            
            'Get coordinates
            dblX1 = dblX2
            dblY1 = dblY2
            
            dblX2 = lngLength * Cos(m_cRAD * (dblAngle * dblX - dblY)) + dblB2
            dblY2 = lngLength * Sin(m_cRAD * (dblAngle * dblX - dblY)) + dblB1
            
            'Check to see if random colours should be used or not
            Randomize
            If lngFTR = "lr" Then lngFR = CLng(255 * Rnd)
            
            Randomize
            If lngFTG = "lr" Then lngFG = CLng(255 * Rnd)
            
            Randomize
            If lngFTB = "lr" Then lngFB = CLng(255 * Rnd)
            
            'Draw line
            If dblX <> 0 Then
                frm.pbx1.Line (dblX1, dblY1)-(dblX2, dblY2), RGB(lngFR, lngFG, lngFB)
            End If
            
        Next dblX
        
    Next dblY
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub
