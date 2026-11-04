VBA MACRO modXLSMapCalcs.bas 
in file: x/xl/vbaProject.bin - OLE stream: 'VBA/modXLSMapCalcs'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Public Sub UpdateMapGPS(Optional iPage As Integer = 0, Optional bDelExist As Boolean = True, _
        Optional sPicName As String = "Picture 1", _
        Optional bConfirm As Boolean = True)
On Error GoTo Error_Func
    Dim sDegS As Single
    Dim sDegE As Single
    Dim vCoord As Variant
    Dim Lwin As Single
    Dim Twin As Single
    Dim Wwin As Single
    Dim Hwin As Single
    Dim sErr As String
    Dim Xanc1Scale As Single
    Dim Yanc1Scale As Single
    Dim Xanc2Scale As Single
    Dim Yanc2Scale As Single
    Dim XRat As Single
    Dim Xmin As Single
    Dim Xmax As Single
    Dim Xfrom As Single
    Dim Xto As Single
    Dim Xstep As Single
    Dim YRat As Single
    Dim Xanc1 As Single
    Dim Xanc2 As Single
    Dim Xanc1Win As Single
    Dim Xanc2Win As Single
    Dim Ymin As Single
    Dim Ymax As Single
    Dim Yfrom As Single
    Dim Yto As Single
    Dim Ystep As Single
    Dim sSheet As String
    Dim Yanc1 As Single
    Dim Yanc2 As Single
    Dim Yanc1Win As Single
    Dim Yanc2Win As Single
    Dim i As Integer
    Dim nGraph As Integer
    If Not IsArray(vWsNames) Then
        Call LoadStartUpVariables
    End If
    sErr = "O.K."
    sSheet = vWsNames(iPage)
    'graph image constants
    If iPage = 36 And sPicName = "Picture 1" Then 'seismic
        Xanc1Scale = 0.318
        Yanc1Scale = 0.296
        Xanc2Scale = 0.382
        Yanc2Scale = 0.585
        vCoord = GetObjectCoords(sSheet, sPicName)
        If Not IsNumeric(vCoord(0)) Then
            sErr = vCoord(0)
        End If
        sDegS = Range("celGPSs")
        sDegE = Range("celGPSw")
        Xmin = 13
        Xmax = 37
        Ymin = 20
        Ymax = 36
        Xfrom = 15
        Xto = 35
        Xstep = 5
        Xanc1 = 20
        Xanc2 = 30
        Yfrom = 25
        Yto = 35
        Ystep = 5
        Yanc1 = 25
        Yanc2 = 35
    ElseIf (iPage = 44 And sPicName = "Picture 1") Or _
            (iPage = 53 And sPicName = "Picture 21") Then  'wind map 2 and 1
        If iPage = 44 Then
            Xanc1Scale = 0.221
            Yanc1Scale = 0.075
            Xanc2Scale = 0.615
            Yanc2Scale = 0.765
            sDegS = Range("celGPSsWind2")
            sDegE = Range("celGPSwWind2")
            Xanc1 = 20
            Xanc2 = 32
            Yanc1 = 22
            Yanc2 = 34
        ElseIf iPage = 53 Then
            Xanc1Scale = 0.228
            Yanc1Scale = 0.16
            Xanc2Scale = 0.668
            Yanc2Scale = 0.685
            sDegS = Range("celGPSsWind1")
            sDegE = Range("celGPSwWind1")
            Xanc1 = 20
            Xanc2 = 32
            Yanc1 = 24
            Yanc2 = 34
        End If
        vCoord = GetObjectCoords(sSheet, sPicName)
        If Not IsNumeric(vCoord(0)) Then
            sErr = vCoord(0)
        End If
        Xmin = 16.4
        Xmax = 33
        Ymin = 22
        Ymax = 34.9
        Xfrom = 18
        Xto = 32
        Xstep = 2
        Yfrom = 22
        Yto = 34
        Ystep = 2
    ElseIf iPage = 46 And (sPicName = "Picture 2" Or sPicName = "Picture 3") Then 'thermal
        sDegS = Range("celGPSThermals")
        sDegE = Range("celGPSThermale")
        vCoord = GetObjectCoords(sSheet, sPicName)
        Xmin = 15
        Xmax = 33
        Ymin = 21.666
        Ymax = 35
        Xfrom = 16
        Xto = 32
        Xstep = 4
        Xanc1 = 18.423333 'Cape Town
        Xanc2 = 28.187111 'Pretoria
        Yfrom = 24
        Yto = 32
        Ystep = 4
        Yanc1 = 25.746056 'Pretoria
        Yanc2 = 33.923778 'Cape Town
        If sPicName = "Picture 2" Then
            Xanc1Scale = 0.225
            Yanc1Scale = 0.315
            Xanc2Scale = 0.52
            Yanc2Scale = 0.572
        ElseIf sPicName = "Picture 3" Then
            Xanc1Scale = 0.218
            Yanc1Scale = 0.315
            Xanc2Scale = 0.52
            Yanc2Scale = 0.6
        End If
    Else
        sErr = "Error: Sheet not programmed yet"
    End If
    If sErr = "O.K." Then
        If sDegS < Ymin Or sDegS > Ymax Or sDegE < Xmin Or sDegE > Xmax Then
            sErr = "Error: The requested position is not within the map boundaries"
        End If
    End If
    If sErr = "O.K." Then
        'coordinates
        Lwin = vCoord(0)
        Twin = vCoord(1)
        Wwin = vCoord(2)
        Hwin = vCoord(3)
        Xanc1Win = Lwin + Xanc1Scale * Wwin
        Xanc2Win = Xanc1Win + Xanc2Scale * Wwin
        XRat = (Xanc2Win - Xanc1Win) / (Xanc2 - Xanc1)
        Yanc1Win = Twin + Yanc1Scale * Hwin
        Yanc2Win = Yanc1Win + Yanc2Scale * Hwin
        YRat = (Yanc2Win - Yanc1Win) / (Yanc2 - Yanc1)
        'delete exist
        If bDelExist Then
            sErr = DeleteObjects(sSheet, , 9, 0)
        End If
        'draw grid
        For i = Xfrom To Xto Step Xstep
            Call DrawLine(sSheet, Array(Xanc1Win + (i - Xanc1) * XRat, Twin, _
                Xanc1Win + (i - Xanc1) * XRat, Twin + Hwin), , 1, Array(255, 0, 0))
        Next
        For i = Yfrom To Yto Step Ystep
            Call DrawLine(sSheet, Array(Lwin, Yanc1Win + (i - Yanc1) * YRat, Lwin + Wwin, _
                Yanc1Win + (i - Yanc1) * YRat), , 1, Array(255, 0, 0))
        Next
        'draw position
        Call DrawLine(sSheet, Array(Lwin, Yanc1Win + (sDegS - Yanc1) * YRat, Lwin + Wwin, _
            Yanc1Win + (sDegS - Yanc1) * YRat), , 2, Array(0, 0, 255))
        Call DrawLine(sSheet, Array(Xanc1Win + (sDegE - Xanc1) * XRat, Twin, _
            Xanc1Win + (sDegE - Xanc1) * XRat, Twin + Hwin), , 2, Array(0, 0, 255))
    End If
Exit_Func:
    If bConfirm And sErr <> "O.K." Then
        Call PopMes(sErr)
    End If
    Exit Sub
Error_Func:
    sErr = ErrMes(99, Err.Description & vbCrLf & "Module: UpdateMapGPS")
    Resume Exit_Func
End Sub
-------------------------------------------------------------------------------
VBA MACRO SheetID044.cls 
in file: x/xl/vbaProject.bin - OLE stream: 'VBA/SheetID044'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo Error_Sub
    If Target.Cells.Count > 1 Or IsEmpty(Target) Then Exit Sub
    If Not IsArray(vWsNames) Then
        Call LoadStartUpVariables
    End If
    bSkipSheet = True
    If Not aSkipTarget(44) Then
        If Target.Address = Range("celGPSsWindd2").Address Or _
                Target.Address = Range("celGPSsWindm2").Address Or _
                Target.Address = Range("celGPSsWinds2").Address Or _
                Target.Address = Range("celGPSwWindd2").Address Or _
                Target.Address = Range("celGPSwWindm2").Address Or _
                Target.Address = Range("celGPSwWinds2").Address Then
            bSkipSheet = False
        End If
    End If
    If Not bSkipSheet Then
        aSkipTarget(44) = True
        Worksheets(vWsNames(44)).Calculate
        Application.ScreenUpdating = False
        Application.Calculation = xlManual
        Call LockWorkSheets(False, 3)
        DoEvents
        Call UpdateMapGPS(44, , "Picture 1")
        Application.Calculate
    End If
Exit_Sub:
    aSkipTarget(44) = False
    If Not bSkipSheet Then
        If Not bProgMode Then Call LockWorkSheets(True, 3)
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
    End If
    Exit Sub
Error_Sub:
    MsgBox "Error in module Sheet44: " & Err.Description, vbCritical, "Error"
    Resume Exit_Sub
End Sub

-------------------------------------------------------------------------------
Public Sub CrossRefWindspeed(Optional sProvinceID As Integer = 0, _
        Optional sWindMapInput As String = "G")
On Error GoTo Error_Sub
    Dim sRa As Range
    Dim sErr As String
    Dim sDef As Double
    sErr = "O.K."
    aSkipTarget(10) = True
    Set sRa = Range("celWindVb0")
    If sWindMapInput = "G" Then
        sDef = 36
    Else
        sDef = 28
    End If
    If sProvinceID <> 0 Then
        Call LockCellEntry(sRa, "=VLOOKUP(celMunicipalityID," & _
            "ranMunicipalities,3,FALSE)")
    Else
        Call LockCellEntry(sRa, , CStr(sDef))
    End If
Exit_Sub:
    aSkipTarget(10) = False
    If sErr <> "O.K." Then
        Call PopMes(sErr)
    End If
    Exit Sub
Error_Sub:
    sErr = ErrMes(99, Err.Description)
    Resume Exit_Sub
End Sub

