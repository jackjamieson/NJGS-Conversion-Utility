Attribute VB_Name = "modProcs"
Option Explicit
'variables for ddmmss <=> decimaldegrees; also shared with lattoat <=> attolat
'and lat/lon <=> state plane.
    Public dblLatdd As Double, dblLondd As Double
    Public dblLatdecimal As Double, dblLondecimal As Double
    Public dblLatDeg As Double, dblLatMin As Double, dblLatSec As Double
    Public dblLonDeg As Double, dblLonMin As Double, dblLonSec As Double
    Public strLatMin As String, strLatSec As String
    Public strLonMin As String, strLonSec As String
    Public blnNAD83 As Boolean

'variables for datum
    Public strPath As String, strQnum As String, strQnam As String
    Public dblNwlat As Double, dblNwlon As Double, dblNw As Double, dblS3 As Double
    Public dblNelat As Double, dblNelon As Double, dblNe As Double, dblS4 As Double
    Public dblSwlat As Double, dblSwlon As Double, dblSw As Double, dblS1 As Double
    Public dblSelat As Double, dblSelon As Double, dblSe As Double, dblS2 As Double
    Dim dblAlatd As Double, dblAlatm As Double, dblAlats As Double, dblAslat As Double
    Dim dblAlond As Double, dblAlonm As Double, dblAlons As Double, dblAslon As Double
    Dim dblNwlatd As Double, dblNwlatm As Double, dblNwlats As Double, dblSnwlat As Double
    Dim dblNwlond As Double, dblNwlonm As Double, dblNwlons As Double, dblSnwlon As Double
    Dim dblSelatd As Double, dblSelatm As Double, dblSelats As Double, dblSselat As Double
    Dim dblSelond As Double, dblSelonm As Double, dblSelons As Double, dblSselon As Double
    Dim dblA As Double, dblB As Double, dblC As Double, dblD As Double
    Dim dblSplat As Double, dblSplon As Double
    Public dblLatDatum As Double, dblLonDatum As Double
    
'variables for LatToAt <=> AtToLat
    Public strAscCoor As String
    Dim sngSouth As Single, sngEast As Single
    Dim sngLatMin As Single, sngLonMin As Single
    Dim intBlkSouth As Integer, intBlkEast As Integer, intBlk As Integer
    Dim intR1 As Integer, intR2 As Integer, intR3 As Integer, intSheet As Integer
    Dim strSheet As String, strBlk As String, strRect As String

'variables for lat/lon <=> state plane
    Public dblNorthing As Double, dblEasting As Double
    Dim dblNorth As Double, dblEast As Double, dblOm As Double, dblPi As Double
    Dim dblFoot As Double, dblSinf As Double, dblCosf As Double
    Dim dblTn As Double, dblTs As Double, dblEts As Double
    Dim dblRn As Double, dblQ As Double, dblQs As Double
    Dim dblB2 As Double, dblB4 As Double, dblB6 As Double
    Dim dblB1 As Double, dblB3 As Double, dblB5 As Double, dblB7 As Double
    Dim dblL As Double, dblLatr As Double, dblLonr As Double
    Dim dblRad As Double, dblCm As Double, dblRo As Double
    Dim dblFi As Double, dblLam As Double, dblS As Double
    Dim dlbL As Double, dblLs As Double
    Dim dblA1 As Double, dblA3 As Double, dblA5 As Double, dblA7 As Double
    Dim dblA2 As Double, dblA4 As Double, dblA6 As Double
    Dim dblX1 As Double, dblX2 As Double, dblX3 As Double, dblX4 As Double
    
'variable for batch files
    Public blnBatch As Boolean
    
'NAD83 constants for transverse mercator projection in New Jersey

'conversion for meters to feet
    Const conFt As Double = 3.28083333
    
'false easting in meters
    Const conFe As Double = 150000
'false northing in meters
    Const conFn As Double = 0
    
'central merdian
    Const conCm As Double = 74.5
'scale factor along the central merdian
    Const conSf As Double = 1 - 1 / 10000
    
'equatorial radius of the elipsoid in meters
    Const conEr As Double = 6378137
    
'reciprocal of the flattening of the elipsoid
    Const conRf As Double = 298.25722210088
'flattening of the elipsoid
    Const conF As Double = 1 / conRf
    
'semi-minor axis conPr = (1 - conF) * conEr
    Const conPr As Double = 63556752.3141403
    
'square of the 1st eccentricity conEsq = (conF + conF - (conF^2))
    Const conEsq As Double = 6.6943800229034E-03
    
'square of the 2nd eccentricity
    Const conEps As Double = conEsq / (1 - conEsq)
    
    Const conEn = (conEr - conPr) / (conEr + conPr)
    Const conEn2 = conEn ^ 2
    Const conEn3 = conEn ^ 3
    Const conEn4 = conEn ^ 4
    
'conA = -1.5 * conEn + (9 / 16) * conEn3
'conB = 0.9375 * conEn2 - (15 / 32) * conEn4
'conC = -(35 / 48) * conEn3
'conD = (135 / 512) * conEn4

'conU0 = 2 * (conA - (2 * conB) + (3 * conC) - (4 * conD))
    Const conU0 As Double = -0.005048250776
'conU2 = 8 * (conB - (4 * conC) + (10 * conD))
    Const conU2 As Double = 0.000021259204
'conU4 = 32 * (conC - (6 * conD))
    Const conU4 As Double = -0.000000111423
'conU6 = 128 * conD
    Const conU6 As Double = 0.000000000626
    
'conU = 1.5 * conEn - (27 / 32) * conEn3
'conV = 1.3125 * conEn2 - (55 / 32) * conEn4
'conW = (151 / 96) * conEn3
'conX = (1097 / 512) * conEn4

'conV0 = 2 * (conU - (2 * conV) + (3 * conW) - (4 * conX))
    Const conV0 As Double = 0.005022893948
'conV2 = 8 * (conV - (4 * conW) + (10 * conX))
    Const conV2 As Double = 0.000029370625
'conV4 = 32 * (conW - (6 * conX))
    Const conV4 As Double = 0.000000235059
'conV6 = 128 * conX
    Const conV6 As Double = 0.000000002181
    
'radius of the rectifying sphere
'conR = conEr * (1 - conEn) * (1 - conEn2) *
'(1 + 2.25 * conEn2 + ( 255 / 64) * conEn4)
    Const conR As Double = 6367449.14577

'southern most parallel of latitude (in radians) for which the northing
'coordinate is zero along the central merdian.
'conRo = 38.83333 / ATN(1) * 4

'rectifying latitude of grid origin
'conOmo = conRo + SIN(conRo) * COS(conRo) *
'(conU0 + conU2 * COS(conRo) ^ 2 +
'conU4 * COS(conRo) ^ 4 + conU6 * COS(conRo) ^ 6

'meridional distance from the equator to the latitude
'of grid origin, multiplied by the central meridian scale factor.
'conSo = conSf * conR * conOmo
    Const conSo As Double = 4299571.6693
    
Public Sub xytoll83()

dblPi = Atn(1) * 4
dblRad = 180 / dblPi
dblCm = conCm / dblRad

'convert to meters
    dblNorth = dblNorthing / conFt
    dblEast = dblEasting / conFt

dblOm = (dblNorth - conFn + conSo) / (conR * conSf)
'footprint latitude
dblX1 = Sin(dblOm) * Cos(dblOm)
dblX2 = conV0 + conV2 * Cos(dblOm) ^ 2
dblX3 = conV4 * Cos(dblOm) ^ 4
dblX4 = conV6 * Cos(dblOm) ^ 6
dblFoot = dblOm + dblX1 * (dblX2 + dblX3 + dblX4)
        
dblSinf = Sin(dblFoot)
dblCosf = Cos(dblFoot)
dblTn = dblSinf / dblCosf
dblTs = dblTn ^ 2
dblEts = conEps * dblCosf ^ 2
dblRn = conEr * conSf / (1 - conEsq * dblSinf ^ 2) ^ 0.5
dblQ = (dblEast - conFe) / dblRn
dblQs = dblQ ^ 2

'calculate the latitude
dblB2 = -dblTn * (1 + dblEts) / 2
dblB4 = -(5 + 3 * dblTs + dblEts * (1 - 9 * dblTs) - 4 * (dblEts) ^ 4) / 12
dblB6 = (61 + 90 * dblTs + 45 * (dblTs ^ 2) + dblEts * (46 - 252 * dblTs - _
        90 * (dblTs ^ 2))) / 360
dblLatr = dblFoot + dblB2 * dblQs * (1 + dblQs * (dblB4 + dblB6 * dblQs))

'calculate the longitude
dblB1 = 1#
dblB3 = -(1 + 2 * dblTs + dblEts) / 6
dblB5 = (5 + 28 * dblTs + 24 * (dblTs ^ 2) + dblEts * (6 + 8 * dblEts)) / 120
dblB7 = -(61 + 662 * dblTs + 1320 * (dblTs ^ 2) + 720 * (dblTs ^ 3)) / 5040

dblL = dblB1 * dblQ * (1 + dblQs * (dblB3 + dblQs * (dblB5 + dblB7 * dblQs)))

dblLonr = dblCm - dblL / dblCosf

'convert radians to decimal degrees
dblLatdecimal = dblLatr * dblRad
dblLondecimal = dblLonr * dblRad

End Sub

Public Sub lltoxy83()

dblPi = Atn(1) * 4
dblRad = 180 / dblPi
dblCm = conCm / dblRad

'conversion to radians
dblFi = dblLatdecimal / dblRad
dblLam = dblLondecimal / dblRad

'rectifying latitude
dblX1 = Sin(dblFi) * Cos(dblFi)
dblX2 = conU0 + conU2 * Cos(dblFi) ^ 2
dblX3 = conU4 * Cos(dblFi) ^ 4
dblX4 = conU6 * Cos(dblFi) ^ 6
dblOm = dblFi + dblX1 * (dblX2 + dblX3 + dblX4)

'meridional distance
dblS = conR * dblOm * conSf
dblTn = Tan(dblFi)
dblTs = dblTn ^ 2
dblEts = conEps * Cos(dblFi) ^ 2
dblL = (dblLam - dblCm) * Cos(dblFi)
dblLs = dblL ^ 2
dblRn = conSf * conEr / (1 - conEsq * Sin(dblFi) ^ 2) ^ 0.5

'calculate the x coordinate (easting)
dblA1 = -dblRn
dblA3 = (1 - dblTs + dblEts) / 6
dblA5 = (5 + dblTs * (dblTs - 18) + dblEts * (14 - 58 * dblTs)) / 120
dblA7 = (61 - 479 * dblTs + 179 * (dblTs ^ 2) - (dblTs ^ 3)) / 5040

dblEast = (dblA1 * dblL * (1 + dblLs * (dblA3 + dblLs * (dblA5 + dblA7 * dblLs))) _
          + conFe) * conFt

'calculate the y coordinate (northing)
dblA2 = dblRn * dblTn / 2
dblA4 = (5 - dblTs + dblEts * (9 + 4 * dblEts)) / 12
dblA6 = (61 - 58 * dblTs + dblTs ^ 2 + dblEts * (270 - 330 * dblTs)) / 360

dblNorth = (dblS - conSo + dblA2 * (dblLs * (1 + dblLs * (dblA4 + dblA6 * dblLs))) _
           + conFn) * conFt

dblNorthing = Round(dblNorth, 3)
dblEasting = Round(dblEast, 3)

End Sub

Public Sub Ddmmss()

'Convert degrees,minutes and seconds to decimal degrees

'Latitude
    dblLatdecimal = dblLatDeg + (dblLatMin / 60) + (dblLatSec / 3600)
    
'Longitude
    dblLondecimal = dblLonDeg + (dblLonMin / 60) + (dblLonSec / 3600)
    
End Sub

Public Sub DecimalDegrees()

'Convert decimal degrees to degrees, minutes and decimal seconds

'Latitude
    dblLatDeg = Int(dblLatdecimal)
    dblLatMin = Int((dblLatdecimal - dblLatDeg) * 60)
    dblLatSec = Round((((dblLatdecimal - dblLatDeg) * 60) - dblLatMin) * 60, 3)
      
    If dblLatSec <= -0.01 Then
        dblLatMin = dblLatMin - 1
        dblLatSec = dblLatSec + 60
    End If
    
    If dblLatSec >= 60 Then
        dblLatMin = dblLatMin + 1
        dblLatSec = dblLatSec - 60
    End If
    
    If dblLatMin >= 60 Then
        dblLatDeg = dblLatDeg + 1
        dblLatMin = dblLatMin - 60
    End If
       
       
'Format minutes and seconds to display properly when < 10
    If dblLatMin < 10 Then
        strLatMin = 0 & dblLatMin
    Else
        strLatMin = dblLatMin
    End If

    If dblLatSec < 10 Then
        strLatSec = 0 & dblLatSec
    Else
        strLatSec = dblLatSec
    End If

    
'Longitude
    dblLonDeg = Int(dblLondecimal)
    dblLonMin = Int((dblLondecimal - dblLonDeg) * 60)
    dblLonSec = Round((((dblLondecimal - dblLonDeg) * 60) - dblLonMin) * 60, 3)
      
    If dblLonSec <= -0.01 Then
        dblLonMin = dblLonMin - 1
        dblLonSec = dblLonSec + 60
    End If
    
    If dblLonSec >= 60 Then
        dblLonMin = dblLonMin + 1
        dblLonSec = dblLonSec - 60
    End If
    
    If dblLonMin >= 60 Then
        dblLonDeg = dblLonDeg + 1
        dblLonMin = dblLonMin - 60
    End If
    
'Format minutes and seconds to display properly when < 10
    If dblLonMin < 10 Then
        strLonMin = 0 & dblLonMin
    Else
        strLonMin = dblLonMin
    End If

    If dblLonSec < 10 Then
        strLonSec = 0 & dblLonSec
    Else
        strLonSec = dblLonSec
    End If
                    
dblLatdd = dblLatDeg & strLatMin & strLatSec
dblLondd = dblLonDeg & strLonMin & strLonSec

End Sub

Public Sub AtToLat()

'Disassemble the atlas sheet coordinate string into sheet, block, and rectangle.

'atlas sheet number
intSheet = Left(strAscCoor, 2)
'1st block digit - increases to south
intBlkSouth = Mid(strAscCoor, 4, 1)
'2nd block digit - increases to east
intBlkEast = Mid(strAscCoor, 5, 1)
'1st digit 3x3 rectangle
intR1 = Mid(strAscCoor, 7, 1)
'2nd digit 3x3 rectangle
intR2 = Mid(strAscCoor, 8, 1)
'3rd digit 3x3 rectangle
intR3 = Mid(strAscCoor, 9, 1)

'check for valid block number
If intBlkSouth > 4 Then
    GoTo Err1_cmdCalculate_Click
    Exit Sub
End If

If intSheet = 36 Then
    If intBlkEast > 6 Then
        GoTo Err1_cmdCalculate_Click
        Exit Sub
    End If
Else
    If intBlkEast > 5 Then
        GoTo Err1_cmdCalculate_Click
        Exit Sub
    End If
End If

'check for valid rectangle number
If intR2 = 0 Then
    GoTo Err1_cmdCalculate_Click
    Exit Sub
End If

If intR3 = 0 Then
    GoTo Err1_cmdCalculate_Click
    Exit Sub
End If

'Latitude and longitude calculation:
'sngSouth = southerning in minutes and sngEast = easting in minutes
'The southerning and easting are determined by successive additions
'of increments as indicated by the given digit of the atlas sheet.
'All increments are measured in minutes of arc. The increment size
'is determined by the given digit. All sngSouth and sngEast increments
'are negative or zero because the reference corner location is the
'northwest corner of the atlas sheet and hence the locations referenced
'to this point are in directions of decreasing latitude and longitude.

'calculate southerning and easting of block coordinates.
sngSouth = intBlkSouth * (-6)
sngEast = (intBlkEast - 1) * (-6)

'calculate southerning and easting addition using 1st digit, 3x3 rectangle.
'increment is 2 minutes
Select Case intR1
    Case Is = 1
        sngSouth = sngSouth - (0 * 2)
        sngEast = sngEast - (0 * 2)
    Case Is = 2
        sngSouth = sngSouth - (0 * 2)
        sngEast = sngEast - (1 * 2)
    Case Is = 3
        sngSouth = sngSouth - (0 * 2)
        sngEast = sngEast - (2 * 2)
    Case Is = 4
        sngSouth = sngSouth - (1 * 2)
        sngEast = sngEast - (0 * 2)
    Case Is = 5
        sngSouth = sngSouth - (1 * 2)
        sngEast = sngEast - (1 * 2)
    Case Is = 6
        sngSouth = sngSouth - (1 * 2)
        sngEast = sngEast - (2 * 2)
    Case Is = 7
        sngSouth = sngSouth - (2 * 2)
        sngEast = sngEast - (0 * 2)
    Case Is = 8
        sngSouth = sngSouth - (2 * 2)
        sngEast = sngEast - (1 * 2)
    Case Is = 9
        sngSouth = sngSouth - (2 * 2)
        sngEast = sngEast - (2 * 2)
    Case Else
        GoTo Err1_cmdCalculate_Click
        Exit Sub
End Select

'calculate southerning and easting addition using 2nd digit, 3x3 rectangle.
'increment is 2/3 minutes
Select Case intR2
    Case Is = 1
        sngSouth = sngSouth - (0 * 2 / 3)
        sngEast = sngEast - (0 * 2 / 3)
    Case Is = 2
        sngSouth = sngSouth - (0 * 2 / 3)
        sngEast = sngEast - (1 * 2 / 3)
    Case Is = 3
        sngSouth = sngSouth - (0 * 2 / 3)
        sngEast = sngEast - (2 * 2 / 3)
    Case Is = 4
        sngSouth = sngSouth - (1 * 2 / 3)
        sngEast = sngEast - (0 * 2 / 3)
    Case Is = 5
        sngSouth = sngSouth - (1 * 2 / 3)
        sngEast = sngEast - (1 * 2 / 3)
    Case Is = 6
        sngSouth = sngSouth - (1 * 2 / 3)
        sngEast = sngEast - (2 * 2 / 3)
    Case Is = 7
        sngSouth = sngSouth - (2 * 2 / 3)
        sngEast = sngEast - (0 * 2 / 3)
    Case Is = 8
        sngSouth = sngSouth - (2 * 2 / 3)
        sngEast = sngEast - (1 * 2 / 3)
    Case Is = 9
        sngSouth = sngSouth - (2 * 2 / 3)
        sngEast = sngEast - (2 * 2 / 3)
    Case Is = 0
        sngSouth = sngSouth
        sngEast = sngEast
    Case Else
        GoTo Err1_cmdCalculate_Click
        Exit Sub
End Select
    
'calculate southerning and easting addition using 3rd digit, 3x3 rectangle.
'increment is 2/9 minutes
Select Case intR3
    Case Is = 1
        sngSouth = sngSouth - (0 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (0 * 2 / 9) - (1 / 9)
    Case Is = 2
        sngSouth = sngSouth - (0 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (1 * 2 / 9) - (1 / 9)
    Case Is = 3
        sngSouth = sngSouth - (0 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (2 * 2 / 9) - (1 / 9)
    Case Is = 4
        sngSouth = sngSouth - (1 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (0 * 2 / 9) - (1 / 9)
    Case Is = 5
        sngSouth = sngSouth - (1 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (1 * 2 / 9) - (1 / 9)
    Case Is = 6
        sngSouth = sngSouth - (1 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (2 * 2 / 9) - (1 / 9)
    Case Is = 7
        sngSouth = sngSouth - (2 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (0 * 2 / 9) - (1 / 9)
    Case Is = 8
        sngSouth = sngSouth - (2 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (1 * 2 / 9) - (1 / 9)
    Case Is = 9
        sngSouth = sngSouth - (2 * 2 / 9) - (1 / 9)
        sngEast = sngEast - (2 * 2 / 9) - (1 / 9)
    Case Is = 0
        sngSouth = sngSouth
        sngEast = sngEast
    Case Else
        GoTo Err1_cmdCalculate_Click
        Exit Sub
End Select
    
'Add easting and southerning to the northwest corner of the appropriate
'arlas sheet. The northwest corner latitude and longitude is in minutes.
Select Case intSheet
    Case Is = 21
        sngSouth = sngSouth + 2484
        sngEast = sngEast + 4512
    Case Is = 22
        sngSouth = sngSouth + 2484
        sngEast = sngEast + 4486
    Case Is = 23
        sngSouth = sngSouth + 2484
        sngEast = sngEast + 4460
    Case Is = 24
        sngSouth = sngSouth + 2456
        sngEast = sngEast + 4512
    Case Is = 25
        sngSouth = sngSouth + 2456
        sngEast = sngEast + 4486
    Case Is = 26
        sngSouth = sngSouth + 2456
        sngEast = sngEast + 4460
    Case Is = 27
        sngSouth = sngSouth + 2428
        sngEast = sngEast + 4512
    Case Is = 28
        sngSouth = sngSouth + 2428
        sngEast = sngEast + 4486
    Case Is = 29
        sngSouth = sngSouth + 2428
        sngEast = sngEast + 4460
    Case Is = 30
        sngSouth = sngSouth + 2400
        sngEast = sngEast + 4538
    Case Is = 31
        sngSouth = sngSouth + 2400
        sngEast = sngEast + 4512
    Case Is = 32
        sngSouth = sngSouth + 2400
        sngEast = sngEast + 4486
    Case Is = 33
        sngSouth = sngSouth + 2400
        sngEast = sngEast + 4460
    Case Is = 34
        sngSouth = sngSouth + 2372
        sngEast = sngEast + 4538
    Case Is = 35
        sngSouth = sngSouth + 2372
        sngEast = sngEast + 4512
    Case Is = 36
        sngSouth = sngSouth + 2372
        sngEast = sngEast + 4486
    Case Is = 37
        sngSouth = sngSouth + 2344
        sngEast = sngEast + 4500
    Case Else
        GoTo Err1_cmdCalculate_Click
        Exit Sub
End Select

'Convert the latitude and longitude from minutes to degrees, minutes,
'and seconds.

'latitude
dblLatDeg = Int(sngSouth / 60)
sngLatMin = (sngSouth / 60 - dblLatDeg) * 60
dblLatMin = Int(sngLatMin)
If dblLatMin < 10 Then
    strLatMin = 0 & dblLatMin
Else
    strLatMin = dblLatMin
End If
dblLatSec = (sngLatMin - dblLatMin) * 60
If dblLatSec < 10 Then
    strLatSec = 0 & dblLatSec
Else
    strLatSec = dblLatSec
End If

'longitude
dblLonDeg = Int(sngEast / 60)
sngLonMin = (sngEast / 60 - dblLonDeg) * 60
dblLonMin = Int(sngLonMin)
If dblLonMin < 10 Then
    strLonMin = 0 & dblLonMin
Else
    strLonMin = dblLonMin
End If
dblLonSec = (sngLonMin - dblLonMin) * 60
If dblLonSec < 10 Then
    strLonSec = 0 & dblLonSec
Else
    strLonSec = dblLonSec
End If

'results in ddmmss.ss
dblLatdd = dblLatDeg & strLatMin & strLatSec
dblLondd = dblLonDeg & strLonMin & strLonSec

Exit Sub

Err1_cmdCalculate_Click:
    dblLatdd = ""
    dblLondd = ""
    Exit Sub
    
End Sub

Public Sub LatToAt()

'Calculate the southerning and easting in minutes
    sngSouth = (dblLatDeg * 60) + dblLatMin + (dblLatSec / 60)
    sngEast = (dblLonDeg * 60) + dblLonMin + (dblLonSec / 60)

'The southerning and easting are compared to the range of latitude and
'longitude covered by each of the atlas sheets. For the appropriate atlas
'sheet, the southerning and easting are subtracted from the latitude and
'longitude of the northwest corner of the atlas sheet. The new values of
'southerning and easting are the number of minutes south and east of the
'northwest corner of the atlas sheet.

strSheet = ""

If sngSouth <= 2484 And sngSouth > 2456 And sngEast <= 4512 And sngEast > 4486 Then
    strSheet = 21
    sngSouth = 2484 - sngSouth
    sngEast = 4512 - sngEast
End If
If sngSouth <= 2484 And sngSouth > 2456 And sngEast <= 4486 And sngEast > 4460 Then
    strSheet = 22
    sngSouth = 2484 - sngSouth
    sngEast = 4486 - sngEast
End If
If sngSouth <= 2484 And sngSouth > 2456 And sngEast <= 4460 And sngEast > 4434 Then
    strSheet = 23
    sngSouth = 2484 - sngSouth
    sngEast = 4460 - sngEast
End If
If sngSouth <= 2456 And sngSouth > 2428 And sngEast <= 4512.3 And sngEast > 4486 Then
    strSheet = 24
    sngSouth = 2456 - sngSouth
    sngEast = 4512 - sngEast
        If sngEast < 0 Then
            sngEast = 0
        End If
End If
If sngSouth <= 2456 And sngSouth > 2428 And sngEast <= 4486 And sngEast > 4460 Then
    strSheet = 25
    sngSouth = 2456 - sngSouth
    sngEast = 4486 - sngEast
End If
If sngSouth <= 2456 And sngSouth > 2428 And sngEast <= 4460 And sngEast > 4434 Then
    strSheet = 26
    sngSouth = 2456 - sngSouth
    sngEast = 4460 - sngEast
End If
If sngSouth <= 2428 And sngSouth > 2400 And sngEast <= 4512 And sngEast > 4486 Then
    strSheet = 27
    sngSouth = 2428 - sngSouth
    sngEast = 4512 - sngEast
End If
If sngSouth <= 2428 And sngSouth > 2400 And sngEast <= 4486 And sngEast > 4460 Then
    strSheet = 28
    sngSouth = 2428 - sngSouth
    sngEast = 4486 - sngEast
End If
If sngSouth <= 2428 And sngSouth > 2400 And sngEast <= 4460 And sngEast > 4434 Then
    strSheet = 29
    sngSouth = 2428 - sngSouth
    sngEast = 4460 - sngEast
End If
If sngSouth <= 2400 And sngSouth > 2372 And sngEast <= 4538 And sngEast > 4512 Then
    strSheet = 30
    sngSouth = 2400 - sngSouth
    sngEast = 4538 - sngEast
End If
If sngSouth <= 2400 And sngSouth > 2372 And sngEast <= 4512 And sngEast > 4486 Then
    strSheet = 31
    sngSouth = 2400 - sngSouth
    sngEast = 4512 - sngEast
End If
If sngSouth <= 2400 And sngSouth > 2372 And sngEast <= 4486 And sngEast > 4460 Then
    strSheet = 32
    sngSouth = 2400 - sngSouth
    sngEast = 4486 - sngEast
End If
If sngSouth <= 2400 And sngSouth > 2372 And sngEast <= 4460 And sngEast > 4434 Then
    strSheet = 33
    sngSouth = 2400 - sngSouth
    sngEast = 4460 - sngEast
End If
If sngSouth <= 2372 And sngSouth > 2344 And sngEast <= 4538 And sngEast > 4512 Then
    strSheet = 34
    sngSouth = 2372 - sngSouth
    sngEast = 4538 - sngEast
End If
If sngSouth <= 2372 And sngSouth > 2344 And sngEast <= 4512 And sngEast > 4486 Then
    strSheet = 35
    sngSouth = 2372 - sngSouth
    sngEast = 4512 - sngEast
End If
If sngSouth <= 2372 And sngSouth > 2344 And sngEast <= 4486 And sngEast > 4452 Then
    strSheet = 36
    sngSouth = 2372 - sngSouth
    sngEast = 4486 - sngEast
End If
If sngSouth <= 2344 And sngSouth > 2332 And sngEast <= 4500 And sngEast > 4474 Then
    strSheet = 37
    sngSouth = 2344 - sngSouth
    sngEast = 4500 - sngEast
End If
If strSheet = "" Then
    GoTo Err1_cmdCalculate_Click
End If

'Block coordinates are found by dividing the southerning and easting by
'6-minutes (size of blocks) and recording integer value of the quotient.
'The decimal portion of the quotient is the new southerning and easting.
'These are the number of minutes south and east of the northwest corner
'of the block.

'block coordinate calculation
intBlkSouth = Int(sngSouth / 6)
intBlkEast = Int(sngEast / 6) + 1

'minutes south of the block's NW corner
sngSouth = ((sngSouth / 6) - Int(sngSouth / 6)) * 6
'minutes east of the block's NW corner
sngEast = ((sngEast / 6) - Int(sngEast / 6)) * 6

'The first 3x3 rectangle coordinate is found by determining what
'2-minute by 2-minute interval the southerning and easting lie.

If sngSouth < 2 And sngSouth >= 0 And sngEast < 2 And sngEast >= 0 Then
    intR1 = 1
End If
If sngSouth < 2 And sngSouth >= 0 And sngEast < 4 And sngEast >= 2 Then
    intR1 = 2
End If
If sngSouth < 2 And sngSouth >= 0 And sngEast < 6 And sngEast >= 4 Then
    intR1 = 3
End If
If sngSouth < 4 And sngSouth >= 2 And sngEast < 2 And sngEast >= 0 Then
    intR1 = 4
End If
If sngSouth < 4 And sngSouth >= 2 And sngEast < 4 And sngEast >= 2 Then
    intR1 = 5
End If
If sngSouth < 4 And sngSouth >= 2 And sngEast < 6 And sngEast >= 4 Then
    intR1 = 6
End If
If sngSouth < 6 And sngSouth >= 4 And sngEast < 2 And sngEast >= 0 Then
    intR1 = 7
End If
If sngSouth < 6 And sngSouth >= 4 And sngEast < 4 And sngEast >= 2 Then
    intR1 = 8
End If
If sngSouth < 6 And sngSouth >= 4 And sngEast < 6 And sngEast >= 4 Then
    intR1 = 9
End If
        
'The southerning and easting are divided by 2 minutes and the decimal
'portion of the quotient is used to calculate a new southerning and
'easting. These are now the number of minutes from the northwest corner
'of the given 2-minute by 2-minute rectangle.

'minutes south of rectangle NW corner
sngSouth = ((sngSouth / 2) - Int(sngSouth / 2)) * 2
'minutes east of rectangle NW corner
sngEast = ((sngEast / 2) - Int(sngEast / 2)) * 2

'the second 3x3 rectangle coordinate is found by determining what
'two-thirds minute by two-thirds minute interval the southerning
'and easting lie

If sngSouth < 2 / 3 And sngSouth >= 0 And sngEast < 2 / 3 And sngEast >= 0 Then
    intR2 = 1
End If
If sngSouth < 2 / 3 And sngSouth >= 0 And sngEast < 4 / 3 And sngEast >= 2 / 3 Then
    intR2 = 2
End If
If sngSouth < 2 / 3 And sngSouth >= 0 And sngEast < 6 / 3 And sngEast >= 4 / 3 Then
    intR2 = 3
End If
If sngSouth < 4 / 3 And sngSouth >= 2 / 3 And sngEast < 2 / 3 And sngEast >= 0 Then
    intR2 = 4
End If
If sngSouth < 4 / 3 And sngSouth >= 2 / 3 And sngEast < 4 / 3 And sngEast >= 2 / 3 Then
    intR2 = 5
End If
If sngSouth < 4 / 3 And sngSouth >= 2 / 3 And sngEast < 6 / 3 And sngEast >= 4 / 3 Then
    intR2 = 6
End If
If sngSouth < 6 / 3 And sngSouth >= 4 / 3 And sngEast < 2 / 3 And sngEast >= 0 Then
    intR2 = 7
End If
If sngSouth < 6 / 3 And sngSouth >= 4 / 3 And sngEast < 4 / 3 And sngEast >= 2 / 3 Then
    intR2 = 8
End If
If sngSouth < 6 / 3 And sngSouth >= 4 / 3 And sngEast < 6 / 3 And sngEast >= 4 / 3 Then
    intR2 = 9
End If

'The southerning and easting are divided by two-thirds and the decimal
'portion of the quotient is used to calculate a new southerning and
'easting. These are now the number of minutes from the northwest corner
'of the given two-thirds minute by two-thirds minute rectangle.

sngSouth = ((sngSouth / (2 / 3)) - Int(sngSouth / (2 / 3))) * (2 / 3)
sngEast = ((sngEast / (2 / 3)) - Int(sngEast / (2 / 3))) * (2 / 3)

'The third 3x3 rectangle coordinate is found by determining what
'two-ninths minute by two-ninths minute interval the southerning
'and easting lie.

If sngSouth < 2 / 9 And sngSouth >= 0 And sngEast < 2 / 9 And sngEast >= 0 Then
    intR3 = 1
End If
If sngSouth < 2 / 9 And sngSouth >= 0 And sngEast < 4 / 9 And sngEast >= 2 / 9 Then
    intR3 = 2
End If
If sngSouth < 2 / 9 And sngSouth >= 0 And sngEast < 6 / 9 And sngEast >= 4 / 9 Then
    intR3 = 3
End If
If sngSouth < 4 / 9 And sngSouth >= 2 / 9 And sngEast < 2 / 9 And sngEast >= 0 Then
    intR3 = 4
End If
If sngSouth < 4 / 9 And sngSouth >= 2 / 9 And sngEast < 4 / 9 And sngEast >= 2 / 9 Then
    intR3 = 5
End If
If sngSouth < 4 / 9 And sngSouth >= 2 / 9 And sngEast < 6 / 9 And sngEast >= 4 / 9 Then
    intR3 = 6
End If
If sngSouth < 6 / 9 And sngSouth >= 4 / 9 And sngEast < 2 / 9 And sngEast >= 0 Then
    intR3 = 7
End If
If sngSouth < 6 / 9 And sngSouth >= 4 / 9 And sngEast < 4 / 9 And sngEast >= 2 / 9 Then
    intR3 = 8
End If
If sngSouth < 6 / 9 And sngSouth >= 4 / 9 And sngEast < 6 / 9 And sngEast >= 4 / 9 Then
    intR3 = 9
End If

'Display results
strBlk = intBlkSouth & intBlkEast
strRect = intR1 & intR2 & intR3
strAscCoor = strSheet & ":" & strBlk & ":" & strRect
Exit Sub

Err1_cmdCalculate_Click:
    If blnBatch = True Then
    strAscCoor = "00:00:000"
    Exit Sub
    Else
    MsgBox prompt:="Location not on an Atlas Sheet!"
    strAscCoor = "00:00:000"
    Exit Sub
    End If
    
End Sub

Public Sub Datum()

If Right(App.Path, 1) = "\" Then
    strPath = App.Path & "NAD27.dat"
    Open strPath For Input As #1
Else
    strPath = App.Path & "\NAD27.dat"
    Open strPath For Input As #1
End If

Do While Not EOF(1)
    Input #1, strQnum, strQnam, dblNwlat, dblNwlon, dblNw, dblS3, dblNelat, dblNelon, _
              dblNe, dblS4, dblSwlat, dblSwlon, dblSw, dblS1, dblSelat, dblSelon, _
              dblSe, dblS2
        If dblLatdd <= dblNwlat And dblLatdd >= dblSelat And _
        dblLondd <= dblNwlon And dblLondd >= dblSelon Then
            Exit Do
        End If
Loop
Close #1

'convert the latitudes and longitudes to seconds
dblAlatd = Left(dblLatdd, 2)
dblAlatm = Mid(dblLatdd, 3, 2)
dblAlats = Mid(dblLatdd, 5, 6)
dblAslat = (dblAlatd * 3600) + (dblAlatm * 60) + dblAlats

dblAlond = Left(dblLondd, 2)
dblAlonm = Mid(dblLondd, 3, 2)
dblAlons = Mid(dblLondd, 5, 6)
dblAslon = (dblAlond * 3600) + (dblAlonm * 60) + dblAlons

dblNwlatd = Left(dblNwlat, 2)
dblNwlatm = Mid(dblNwlat, 3, 2)
dblNwlats = Mid(dblNwlat, 5, 6)
dblSnwlat = (dblNwlatd * 3600) + (dblNwlatm * 60) + dblNwlats

dblNwlond = Left(dblNwlon, 2)
dblNwlonm = Mid(dblNwlon, 3, 2)
dblNwlons = Mid(dblNwlon, 5, 6)
dblSnwlon = (dblNwlond * 3600) + (dblNwlonm * 60) + dblNwlons

dblSelatd = Left(dblSelat, 2)
dblSelatm = Mid(dblSelat, 3, 2)
dblSelats = Mid(dblSelat, 5, 6)
dblSselat = (dblSelatd * 3600) + (dblSelatm * 60) + dblSelats

dblSelond = Left(dblSelon, 2)
dblSelonm = Mid(dblSelon, 3, 2)
dblSelons = Mid(dblSelon, 5, 6)
dblSselon = (dblSelond * 3600) + (dblSelonm * 60) + dblSelons

'compute the distance weights for the point to be converted in arc seconds
dblA = dblSnwlon - dblAslon
dblB = dblAslon - dblSselon
dblC = dblAslat - dblSselat
dblD = dblSnwlat - dblAslat

'compute the latitude and longitude shifts
dblSplat = ((dblSw * dblB * dblD) + (dblSe * dblA * dblD) + _
            (dblNw * dblB * dblC) + (dblNe * dblA * dblC)) / _
            ((dblA + dblB) * (dblC + dblD))
            
dblSplon = ((dblS1 * dblB * dblD) + (dblS2 * dblA * dblD) + _
            (dblS3 * dblB * dblC) + (dblS4 * dblA * dblC)) / _
            ((dblA + dblB) * (dblC + dblD))

If blnNAD83 = True Then
    'datum shift NAD27 to NAD83
    dblLatDatum = (dblAslat + dblSplat) / 3600
    dblLonDatum = (dblAslon + dblSplon) / 3600
Else
    'datum shift NAD83 to NAD27
    dblLatDatum = (dblAslat - dblSplat) / 3600
    dblLonDatum = (dblAslon - dblSplon) / 3600
End If


End Sub

