Attribute VB_Name = "Module1"
Option Explicit
' Z�KLADN� PARAMETRY
Const PRINTORIENTATION As String = "Landscape" ' Portrait nebo Landscape
Public TIMEFROM As Date
Public TIMETO As Date
Public MINKM As Double
Public SPEEDCITY As Double
Public SPEEDOUT As Double
Public MAXTRIPSPERDAY As Long
Public TANKCAPACITY As Double
Sub Generovat()
' Optimalizace v�konu
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Application.EnableEvents = False
' On Error GoTo ErrorHandler
 
' Deklarace prom�nn�ch
Dim wsConfig As Worksheet, wsP As Worksheet, wsL As Worksheet
Dim wsS As Worksheet, wsK As Worksheet
Dim spotreba As Double, tacho As Double, domov As String
Dim mapa As Object
Dim posledniP As Long, typCol As Long, domovCol As Long
Dim monthlySheets As Object, monthData As Object
Dim tankPeriods As Object
Dim key As Variant

' Inicializace list�
Set wsConfig = Sheets("Konfigurace a start")
Set wsP = Sheets("N�kup PHM")
Set wsL = Sheets("Lokality")
Set wsS = Sheets("�erpac� stanice")
Set monthlySheets = CreateObject("Scripting.Dictionary")
Set monthData = CreateObject("Scripting.Dictionary")
Set tankPeriods = CreateObject("Scripting.Dictionary")

With Sheets("Konfigurace a start")
    TIMEFROM = CDate(.Range("H7").Value)
    TIMETO = CDate(.Range("H9").Value)
    MINKM = CDbl(.Range("H13").Value)
    SPEEDCITY = CDbl(.Range("H15").Value)
    SPEEDOUT = CDbl(.Range("H17").Value)
    MAXTRIPSPERDAY = CLng(.Range("H11").Value)
    TANKCAPACITY = CDbl(.Range("C11").Value)
End With

' Z�kladn� parametry auta
spotreba = CDbl(wsConfig.Range("C9").Value)
tacho = CDbl(wsConfig.Range("C13").Value)
domov = CStr(wsConfig.Range("C21").Value)

' Slovn�k �erpac�ch stanic
Set mapa = CreateObject("Scripting.Dictionary")
Dim rSt As Long
For rSt = 2 To wsS.Cells(wsS.Rows.count, 1).End(xlUp).Row
    If wsS.Cells(rSt, 1).Value <> "" Then
        mapa(CStr(wsS.Cells(rSt, 1).Value)) = CStr(wsS.Cells(rSt, 2).Value)
    End If
Next rSt

' Naj�t sloupec Typ a domovskou lokalitu
Call FindColumns(wsL, typCol, domovCol, domov)

' Vy�i�t�n� star�ch list� knihy j�zd
Call DeleteExistingKnihaJizdSheets

' Inicializace gener�toru n�hodn�ch ��sel
Randomize

' F�ZE 1: ANAL�ZA - Rozd�len� tankovac�ch obdob�
Call AnalyzeTankPeriods(wsP, tankPeriods, mapa)

' F�ZE 2: GENEROV�N� - Vytvo�en� j�zd pro ka�d� obdob�
Call GenerateTripsForPeriods(tankPeriods, wsL, wsConfig, wsP, monthlySheets, monthData, _
                            spotreba, tacho, domov, domovCol, typCol, mapa)

' Dokon�en� list�
Dim sheetCount As Long
sheetCount = monthlySheets.count
For Each key In monthlySheets.Keys
    Call FinalizeMonthlySheet(monthlySheets(key), monthData(key) - 1)
Next key

' Vytvo�en� souhrnn�ho listu
Call CreateSummarySheet(monthlySheets, monthData)

CleanUp:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
Set wsConfig = Nothing
Set wsP = Nothing
Set wsL = Nothing
Set wsS = Nothing
Set wsK = Nothing
Set mapa = Nothing
Set monthlySheets = Nothing
Set monthData = Nothing
Set tankPeriods = Nothing

MsgBox "Hotovo! Vygenerov�no " & sheetCount & " m�s��n�ch list�.", vbInformation
Exit Sub

ErrorHandler:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
MsgBox "Chyba: " & Err.Description & " (��slo " & Err.Number & ")", vbCritical
End Sub
' F�ZE 1: ANAL�ZA - Rozd�len� tankovac�ch obdob�
Sub AnalyzeTankPeriods(wsP As Worksheet, tankPeriods As Object, mapa As Object)
Dim rTank As Long, posledniP As Long
Dim periodKey As String, startDate As Date, endDate As Date
Dim fuelAmount As Double, tankStation As String
posledniP = wsP.Cells(wsP.Rows.count, 1).End(xlUp).Row

For rTank = 2 To posledniP
    If IsNumeric(wsP.Cells(rTank, 3).Value) Then
        fuelAmount = CDbl(wsP.Cells(rTank, 3).Value)
        If fuelAmount > 0 Then
            startDate = wsP.Cells(rTank, 1).Value
            tankStation = CStr(wsP.Cells(rTank, 2).Value)
            
            If rTank < posledniP Then
                endDate = wsP.Cells(rTank + 1, 1).Value
            Else
                endDate = DateSerial(9999, 12, 31)
            End If
            
            periodKey = "Period_" & rTank
            Dim periodInfo As Object
            Set periodInfo = CreateObject("Scripting.Dictionary")
            periodInfo("StartDate") = startDate
            periodInfo("EndDate") = endDate
            periodInfo("FuelAmount") = fuelAmount
            periodInfo("TankStation") = tankStation
            periodInfo("FuelPrice") = wsP.Cells(rTank, 5).Value
            periodInfo("TankRow") = rTank
            
            If mapa.Exists(tankStation) Then
                periodInfo("TankLocation") = mapa(tankStation)
            Else
                periodInfo("TankLocation") = ""
            End If
            
            Set tankPeriods(periodKey) = periodInfo
        End If
    End If
Next rTank

End Sub
' F�ZE 2: GENEROV�N� - Vytvo�en� j�zd pro ka�d� obdob�
Sub GenerateTripsForPeriods(tankPeriods As Object, wsL As Worksheet, wsA As Worksheet, wsP As Worksheet, _
monthlySheets As Object, monthData As Object, spotreba As Double, _
tacho As Double, domov As String, domovCol As Long, typCol As Long, mapa As Object)
Dim periodKey As Variant, periodInfo As Object
Dim currentFuel As Double, currentLocation As Long
Dim currentDate As Date
Dim key As Variant

currentFuel = 5 ' Za��n�me s t�m�� pr�zdnou n�dr��
currentLocation = domovCol

For Each periodKey In tankPeriods.Keys
    Set periodInfo = tankPeriods(periodKey)
       
    ' Kontrola p�ebytku - bere v �vahu spot�ebu na cest� k pump�
       Dim effectiveFuelBeforeTank As Double
       effectiveFuelBeforeTank = currentFuel
       Dim fuelConsumedToPump As Double
    
       If fuelConsumedToPump > 0 Then
           effectiveFuelBeforeTank = currentFuel - fuelConsumedToPump
       End If
    
       If effectiveFuelBeforeTank + periodInfo("FuelAmount") > TANKCAPACITY Then
           Dim overflow As Double
           overflow = effectiveFuelBeforeTank + periodInfo("FuelAmount") - TANKCAPACITY
           Dim neededKm As Double
           neededKm = overflow / (spotreba / 100)
    
           Call MarkProblem(monthlySheets, monthData, periodInfo("StartDate"), _
               "POZOR: Dal�� tankov�n� p�epln� " & TANKCAPACITY & " l n�dr� o " & Format(overflow, "0.0") & "l. Je pot�eba najezdit " & Format(neededKm, "0") & " km.", _
               spotreba, wsP, periodInfo("TankRow"))
           currentFuel = TANKCAPACITY
       Else
           currentFuel = currentFuel + periodInfo("FuelAmount")
       End If
    
         
    currentDate = periodInfo("StartDate")
    
    Call GenerateTripsForOnePeriod(tankPeriods, periodInfo, wsL, wsA, wsP, currentDate, _
                                   currentFuel, currentLocation, spotreba, tacho, _
                                   domovCol, typCol, monthlySheets, monthData, mapa)
    

    If currentFuel < 0 Then currentFuel = 0
Next periodKey

End Sub
' Generov�n� j�zd pro jedno tankovac� obdob�
Sub GenerateTripsForOnePeriod(tankPeriods As Object, periodInfo As Object, wsL As Worksheet, wsA As Worksheet, wsP As Worksheet, _
ByRef currentDate As Date, ByRef currentFuel As Double, _
ByRef currentLocation As Long, spotreba As Double, ByRef tacho As Double, _
domovCol As Long, typCol As Long, monthlySheets As Object, monthData As Object, _
mapa As Object)

Dim currentTime As Date
Dim tankingWritten As Boolean
Dim periodKey      As Variant     ' iterace p�es tankPeriods
Dim randomMinutes  As Long        ' generov�n� n�hodn�ch minut
Dim mimoPrahaCount As Long
Dim randomBreak    As Long        ' p�est�vka mezi j�zdami
Dim nextPumpDist As Variant       ' vzd�lenost k dal�� pump�
Dim loopCounter As Long
Dim dayHasMimoPraha As Boolean
Dim fuelReserveNeeded As Double   ' pot�ebn� rezerva paliva
Dim bestDestination As Long       ' pro v�b�r c�le
Dim tripDistance   As Double      ' vzd�lenost cesty
Dim isPrivateTrip  As Boolean     ' ozna�en� typu cesty
Dim endDate As Date, dailyTrips As Long
Dim homeStartTime As Date, homeEndTime As Date
Dim startTime      As Date        ' po��te�n� �as cesty
Dim endTime        As Date        ' koncov� �as cesty
Dim homeDistance   As Double      ' vzd�lenost dom�
Dim homeTripDuration As Double    ' doba j�zdy dom�
Dim visitedToday   As Object      ' slovn�k nav�t�ven�ch lokalit

periodInfo("Km") = 0

' Spr�vn� upozorn�n� na p�ebytek paliva p�i tankov�n�
Dim fuelConsumedToPump As Double


' V�po�et pot�ebn� rezervy na konci obdob�
Dim neededReserve As Double
neededReserve = CalculateNeededFuelReserve("Period_" & periodInfo("TankRow"), tankPeriods, wsL, domovCol)
Dim maxAllowedFuel As Double
maxAllowedFuel = currentFuel - neededReserve

endDate = periodInfo("EndDate")
mimoPrahaCount = 0
loopCounter = 0
tankingWritten = False

Do While currentDate < endDate

    ' spr�vn� upozorn�n� na p�ebytek paliva p�i tankov�n�
    fuelConsumedToPump = 0
    
    loopCounter = loopCounter + 1
    dailyTrips = 0
    dayHasMimoPraha = False
    randomMinutes = Int(Rnd * 120)
    currentTime = TimeValue(TIMEFROM) + TimeSerial(0, randomMinutes, 0)

    Set visitedToday = CreateObject("Scripting.Dictionary")
    
    ' Jednoduch� kontrola, zda m�me v n�dr�i alespo� minim�ln� rezervu pro j�zdy.
    Const MIN_FUEL_TO_START_DAY As Double = 5 ' Minim�ln� 5 litr� pro zah�jen� dne
    If currentFuel <= MIN_FUEL_TO_START_DAY Then
        currentDate = DateAdd("d", 1, currentDate)
        GoTo NextDay
    End If

    'Aby byl respektov�n limit na den
    Do While dailyTrips < MAXTRIPSPERDAY And currentFuel > 5 And currentFuel > neededReserve
                 
        bestDestination = SelectBestDestination(wsL, currentLocation, currentFuel, _
                                              spotreba, mimoPrahaCount, dayHasMimoPraha, _
                                              dailyTrips, domovCol, typCol, isPrivateTrip, _
                                              visitedToday, periodInfo, mapa)
        
        If bestDestination = -1 Or bestDestination = 0 Then
            Exit Do
        End If
        
        tripDistance = GetDistance(wsL, currentLocation, bestDestination)
        
        If tripDistance * spotreba / 100 > currentFuel Then
            Exit Do
        End If
        
        ' Ozna�it c�l jako nav�t�ven�
        visitedToday(bestDestination) = True
        
        If GetLocationType(wsL, currentLocation, typCol) = "Praha" And _
           GetLocationType(wsL, bestDestination, typCol) = "MimoPraha" Then
            mimoPrahaCount = mimoPrahaCount + 1
            dayHasMimoPraha = True
        End If
        
        ' Ozna�it jako soukromou, pokud je to posledn� j�zda dne
        If dailyTrips = MAXTRIPSPERDAY - 1 Then
            isPrivateTrip = True
        End If
        
        ' V�po�et �as� podle vzd�lenosti a rychlosti
        Dim tripDuration As Double
        If GetLocationType(wsL, currentLocation, typCol) = "Praha" And _
           GetLocationType(wsL, bestDestination, typCol) = "Praha" Then
            tripDuration = tripDistance / SPEEDCITY
        Else
            tripDuration = tripDistance / SPEEDOUT
        End If
        
        startTime = currentTime
        ' Spr�vn� p�id�n� minut pomoc� DateAdd
        endTime = DateAdd("n", Int(tripDuration * 60), currentTime)
        
        ' Z�pis j�zdy s tankov�n�m na prvn� j�zdu dne
        Dim fuelAmount As Double, fuelPrice As Double, fuelStation As String
        If currentDate = periodInfo("StartDate") And dailyTrips = 0 Then
            fuelAmount = periodInfo("FuelAmount")
            fuelPrice = periodInfo("FuelPrice")
            fuelStation = periodInfo("TankStation")
        Else
            fuelAmount = 0
            fuelPrice = 0
            fuelStation = ""
        End If
              
      
        ' Nejd��ve ode�teme palivo
        currentFuel = currentFuel - (tripDistance * spotreba / 100)
        
        Call WriteTrip(monthlySheets, monthData, currentDate, wsA, _
                      GetLocationName(wsL, currentLocation), _
                      GetLocationName(wsL, bestDestination), _
                      tripDistance, tacho, fuelAmount, fuelPrice, fuelStation, _
                      startTime, endTime, isPrivateTrip, currentFuel)
    
        tacho = tacho + tripDistance
        currentLocation = bestDestination
        dailyTrips = dailyTrips + 1
        periodInfo("Km") = periodInfo("Km") + tripDistance
        
        ' spr�vn� upozorn�n� na p�ebytek paliva p�i tankov�n�
        If fuelAmount > 0 Then
           fuelConsumedToPump = tripDistance * spotreba / 100
        End If
        
        ' spr�vn� upozorn�n� na p�ebytek paliva p�i tankov�n�
        If fuelAmount > 0 Then
            fuelConsumedToPump = tripDistance * spotreba / 100
        End If

        ' N�hodn� p�est�vka mezi j�zdami 60-180 minut
        randomBreak = 60 + Int(Rnd * 121) ' Generuje 60-180 minut
        currentTime = DateAdd("n", randomBreak, endTime)
    Loop
    
        ' Vynucen� a spolehliv� n�vrat dom� na konci dne
        homeDistance = GetDistance(wsL, currentLocation, domovCol)
        
        ' Pokud aktu�ln� lokalita nen� domov a existuje cesta dom� (vzd�lenost > 0)
        If currentLocation <> domovCol And homeDistance > 0 Then
            
            ' Ov���me pro jistotu, zda m�me dost paliva. D�ky vylep�en� logice
            ' ve funkci SelectBestDestination by zde v�dy m�lo b�t paliva dostatek.
            Dim fuelNeededHome As Double
            fuelNeededHome = (homeDistance * spotreba / 100)
            
            If currentFuel >= fuelNeededHome Then
                ' V�po�et �asu cesty dom�
                If GetLocationType(wsL, currentLocation, typCol) = "Praha" Then
                    homeTripDuration = homeDistance / SPEEDCITY
                Else
                    homeTripDuration = homeDistance / SPEEDOUT
                End If
                
                homeStartTime = currentTime
                homeEndTime = DateAdd("n", Int(homeTripDuration * 60), homeStartTime)

                ' Nejd��ve ode�teme palivo pro cestu dom�
                currentFuel = currentFuel - fuelNeededHome
                
                Call WriteTrip(monthlySheets, monthData, currentDate, wsA, _
                              GetLocationName(wsL, currentLocation), _
                              GetLocationName(wsL, domovCol), _
                              homeDistance, tacho, 0, 0, "", _
                              homeStartTime, homeEndTime, True, _
                              currentFuel) ' Posledn� cesta dne je soukrom�
                
                tacho = tacho + homeDistance
                currentLocation = domovCol
                periodInfo("Km") = periodInfo("Km") + homeDistance
            Else
                ' Tato situace by po oprav�ch nem�la nastat.
                ' Zap�eme pozn�mku pro p��pad, �e by se tak p�esto stalo.
                Call MarkProblem(monthlySheets, monthData, currentDate, _
                                "CHYBA: Nedostatek paliva pro n�vrat dom� z lokality " & _
                                GetLocationName(wsL, currentLocation), spotreba, wsP, 0)
            End If
        End If
        
NextDay:
currentDate = DateAdd("d", 1, currentDate)
Loop

'If loopCounter >= MAXLOOPITERATIONS Then
'    Call MarkProblem(monthlySheets, monthData, periodInfo("StartDate"), _
'                    "VAROV�N�: Dosa�en bezpe�nostn� limit iterac� v obdob� " & _
'                       Format(periodInfo("StartDate"), "dd.mm.yyyy"), spotreba, wsP, 0)
'End If

End Sub
' Inteligentn� v�b�r nejlep��ho c�le s vyu�it�m mapy pump
Function SelectBestDestination(wsL As Worksheet, currentLocation As Long, currentFuel As Double, _
spotreba As Double, mimoPrahaCount As Long, dayHasMimoPraha As Boolean, _
dailyTrips As Long, domovCol As Long, typCol As Long, _
ByRef isPrivateTrip As Boolean, visitedToday As Object, _
periodInfo As Object, mapa As Object) As Long
Dim bestCol As Long, maxDistance As Double, cCol As Long
Dim distance As Double, currentType As String, targetType As String
Dim posledniCol As Long

bestCol = 0
maxDistance = 0
isPrivateTrip = False
posledniCol = wsL.Cells(1, wsL.Columns.count).End(xlToLeft).Column
currentType = GetLocationType(wsL, currentLocation, typCol)

' Pokud je to den tankov�n� mimo Prahu, mus�me jet do spr�vn� lokality
If currentLocation = domovCol And dailyTrips = 0 And _
   periodInfo("TankLocation") <> "" And periodInfo("TankLocation") <> "Praha" Then
    
    For cCol = 2 To posledniCol
        If cCol <> domovCol And cCol <> typCol Then
            If GetLocationName(wsL, cCol) = periodInfo("TankLocation") Then
                distance = GetDistance(wsL, currentLocation, cCol)
                If distance >= 3 And (distance * spotreba / 100 + 5) <= currentFuel Then
                    SelectBestDestination = cCol
                    Exit Function
                End If
            End If
        End If
    Next cCol
End If

' Validace pra�sk�ch pump - start nebo c�l mus� b�t v Praze
If currentLocation = domovCol And dailyTrips = 0 And _
   periodInfo("TankLocation") = "Praha" Then
   ' Pro pra�sk� pumpy kontrolujeme, �e alespo� jedna strana cesty je v Praze
   ' Zat�m jen nastav�me flag pro dal�� kontrolu
End If

' Pokud je to posledn� j�zda dne, n�vrat dom�
If dailyTrips = MAXTRIPSPERDAY - 1 Then
    If currentLocation <> domovCol Then
        distance = GetDistance(wsL, currentLocation, domovCol)
        If distance >= 3 And (distance * spotreba / 100) <= currentFuel Then
            bestCol = domovCol
            isPrivateTrip = True
        End If
    End If
    SelectBestDestination = bestCol
    Exit Function
End If

' Proch�zet mo�n� c�le, vynechat u� nav�t�ven�
For cCol = 2 To posledniCol
    If cCol <> currentLocation And cCol <> domovCol And cCol <> typCol Then
        ' P�esko�it ji� nav�t�ven� lokality dnes
        If visitedToday.Exists(cCol) Then GoTo SkipDestination
        
        distance = GetDistance(wsL, currentLocation, cCol)
        If distance >= 3 Then
            targetType = GetLocationType(wsL, cCol, typCol)
            
            ' Kontrola omezen� mimo Prahu
            If currentType = "Praha" And targetType = "MimoPraha" Then
                If mimoPrahaCount >= 2 Then GoTo SkipDestination
                If dayHasMimoPraha Then GoTo SkipDestination
            End If
            
                ' Kontrola paliva - Z�SADN�
                ' V�dy mus�me zajistit, �e po dokon�en� napl�novan� cesty zbude dostatek
                ' paliva na n�vrat dom� z c�lov� lokality.
                Dim fuelNeeded As Double
                Dim homeDistance As Double
                
                ' Vypo��t�me vzd�lenost dom� z C�LOV� lokality (cCol) zva�ovan� cesty
                homeDistance = GetDistance(wsL, cCol, domovCol)
                
                ' Pot�ebn� palivo = palivo na tuto cestu + palivo na cestu dom� + rezerva 5 litr�
                fuelNeeded = (distance + homeDistance) * spotreba / 100 + 5
            
            ' Pro pra�sk� tankov�n� mus� cesta v�dy za��nat z Prahy
            If periodInfo("TankLocation") = "Praha" And dailyTrips = 0 Then
                If currentType <> "Praha" Then
                    GoTo SkipDestination
                End If
            End If

            If fuelNeeded <= currentFuel And distance > maxDistance Then
                bestCol = cCol
                maxDistance = distance
            End If
        End If
    End If

SkipDestination:
Next cCol
If bestCol = 0 Then
    SelectBestDestination = -1
Else
    SelectBestDestination = bestCol
End If

End Function
' Vr�t� sloupec lokality p��t�ho tankov�n� podle data
Function FindNextPumpColumn(currDate As Date, tankPeriods As Object, wsL As Worksheet, domovCol As Long) As Long
Dim periodKey As Variant, startDate As Date
Dim pumpLoc As String, colIndex As Variant
For Each periodKey In tankPeriods.Keys
startDate = tankPeriods(periodKey)("StartDate")
If startDate > currDate Then
pumpLoc = tankPeriods(periodKey)("TankLocation")
colIndex = Application.Match(pumpLoc, wsL.Rows(1), 0)
If Not IsError(colIndex) Then
FindNextPumpColumn = CLng(colIndex)
Exit Function
End If
End If
Next
FindNextPumpColumn = domovCol ' kdy� ��dn� dal�� pumpa, vra� domov
End Function


' Z�pis j�zdy s �asem a typem cesty verze bez sloupce s palivem v n�dr�i
' Sub WriteTrip(monthlySheets As Object, monthData As Object, _
' tripDate As Date, wsA As Worksheet, _
' fromLocation As String, toLocation As String, _
' distance As Double, ByRef tacho As Double, _
' fuelAmount As Double, fuelPrice As Double, fuelStation As String, _
' startTime As Date, endTime As Date, isPrivateTrip As Boolean)

' Z�pis j�zdy s �asem a typem cesty
' Do�asn� �prava pro sloupec s aktu�ln�m stavem paliva
Sub WriteTrip(monthlySheets As Object, monthData As Object, _
              tripDate As Date, wsA As Worksheet, _
              fromLocation As String, toLocation As String, _
              distance As Double, ByRef tacho As Double, _
              fuelAmount As Double, fuelPrice As Double, fuelStation As String, _
              startTime As Date, endTime As Date, isPrivateTrip As Boolean, _
              ByVal currentFuelState As Double)


Dim shName As String, wsK As Worksheet, r As Long
shName = "Kniha j�zd " & Format(tripDate, "MM-YYYY")

If Not monthlySheets.Exists(shName) Then
    Set wsK = CreateMonthlySheet(shName, tripDate)
    Set monthlySheets(shName) = wsK
    monthData(shName) = 6
End If

Set wsK = monthlySheets(shName)
r = monthData(shName)

Dim wsConfig As Worksheet
Set wsConfig = Sheets("Konfigurace a start")

With wsK
    .Cells(r, 1).Value = wsConfig.Range("C19").Value ' Jm�no �idi�e
    .Cells(r, 2).Value = wsConfig.Range("C7").Value  ' Zna�ka auta
    .Cells(r, 3).Value = wsConfig.Range("C15").Value ' RZ
    .Cells(r, 4).Value = wsConfig.Range("C17").Value ' Typ vozidla
    .Cells(r, 5).Value = tripDate
    
    ' Kontrola �asov�ch limit� - nesm� p�es p�lnoc
    Dim adjustedEndTime As Date
    adjustedEndTime = endTime
    
    ' Pokud �as konce p�esahuje 23:59, omez ho
    If Hour(adjustedEndTime) = 0 And adjustedEndTime > DateValue(tripDate) Then
        adjustedEndTime = tripDate + TimeValue("23:59:00")
    End If
    
    .Cells(r, 6).Value = startTime
    .Cells(r, 7).Value = tripDate
    .Cells(r, 8).Value = adjustedEndTime
    .Cells(r, 9).Value = fromLocation & " � " & toLocation
    .Cells(r, 10).Value = IIf(isPrivateTrip, "soukrom�", "slu�ebn�")
    .Cells(r, 11).Value = distance
    .Cells(r, 12).Value = tacho + distance
    .Cells(r, 13).Value = fuelAmount
    .Cells(r, 14).Value = fuelPrice
    .Cells(r, 15).Value = fuelStation
    ' Sloupec s aktu�ln�m stavem paliva
    .Cells(r, 17).Value = currentFuelState
    
    If r Mod 2 = 0 Then
        .Range("A" & r & ":O" & r).Interior.Color = RGB(242, 242, 242)
    End If
End With

tacho = tacho + distance
monthData(shName) = r + 1

End Sub
' Ozna�en� probl�m� na samostatn�m ��dku
Sub MarkProblem(monthlySheets As Object, monthData As Object, problemDate As Date, _
problemText As String, spotreba As Double, wsP As Worksheet, tankRow As Long)
Dim currentMonth As String, currentYear As String, sheetName As String
Dim wsK As Worksheet
currentMonth = Format(problemDate, "MM")
currentYear = Format(problemDate, "YYYY")
sheetName = "Kniha j�zd " & currentMonth & "-" & currentYear

If Not monthlySheets.Exists(sheetName) Then
    Set wsK = CreateMonthlySheet(sheetName, problemDate)
    Set monthlySheets(sheetName) = wsK
    monthData(sheetName) = 6
End If
Set wsK = monthlySheets(sheetName)

Dim rowOut As Long
rowOut = monthData(sheetName)

wsK.Range("A" & rowOut & ":O" & rowOut).Interior.Color = RGB(255, 200, 200)
wsK.Cells(rowOut, 16).Value = problemText
wsK.Cells(rowOut, 16).Font.Color = RGB(255, 0, 0)
wsK.Cells(rowOut, 16).Font.Bold = True

monthData(sheetName) = rowOut + 1

End Sub
' Vytvo�en� m�s��n�ho listu s kompletn�m form�tov�n�m
Function CreateMonthlySheet(sheetName As String, sampleDate As Date) As Worksheet
Dim ws As Worksheet
Set ws = Worksheets.Add
On Error Resume Next
ws.Name = sheetName
On Error GoTo 0

With ws
    .Cells(1, 1).Value = "Kniha j�zd"
    .Cells(1, 1).Font.Size = 16
    .Cells(1, 1).Font.Bold = True
    .Cells(1, 1).HorizontalAlignment = xlCenter
    .Range("A1:O1").Merge
    .Range("A1:O1").Interior.Color = RGB(68, 114, 196)
    .Range("A1:O1").Font.Color = RGB(255, 255, 255)
    .Range("A1:O1").Borders.Weight = xlThick
    
    .Cells(2, 1).Value = "za obdob�"
    .Cells(2, 1).Font.Size = 12
    .Cells(2, 1).HorizontalAlignment = xlCenter
    .Range("A2:O2").Merge
    .Range("A2:O4").Interior.Color = RGB(217, 225, 242)
    
    Dim firstDay As Date, lastDay As Date
    firstDay = DateSerial(Year(sampleDate), Month(sampleDate), 1)
    lastDay = DateSerial(Year(sampleDate), Month(sampleDate) + 1, 0)
    
    .Cells(3, 1).Value = Format(firstDay, "dd.mm.yyyy") & " - " & Format(lastDay, "dd.mm.yyyy")
    .Cells(3, 1).Font.Size = 12
    .Cells(3, 1).Font.Bold = True
    .Cells(3, 1).HorizontalAlignment = xlCenter
    .Range("A3:O3").Merge
    '.Range("A3:O3").Interior.Color = RGB(217, 225, 242)
    .Range("A1:O3").HorizontalAlignment = xlCenter
    .Rows(4).RowHeight = 10
    
    .Cells(5, 1).Value = "�idi�"
    .Cells(5, 2).Value = "Vozidlo"
    .Cells(5, 3).Value = "RZ vozidla"
    .Cells(5, 4).Value = "Typ vozidla"
    .Cells(5, 5).Value = "Datum po��tku"
    .Cells(5, 6).Value = "�as po��tku"
    .Cells(5, 7).Value = "Datum konce"
    .Cells(5, 8).Value = "�as konce"
    .Cells(5, 9).Value = "Cesta"
    .Cells(5, 10).Value = "Typ cesty"
    .Cells(5, 11).Value = "Ujet� km"
    .Cells(5, 12).Value = "Stav tachometru"
    .Cells(5, 13).Value = "�erp�no litr� PHM"
    .Cells(5, 14).Value = "Cena za PHM"
    .Cells(5, 15).Value = "M�sto �erp�n�"
    
    .Cells(5, 16).Value = "Pozn�mky"
    .Cells(5, 16).Font.Bold = True
    .Cells(5, 16).HorizontalAlignment = xlCenter
    .Cells(5, 16).Interior.Color = RGB(91, 155, 213)
    .Cells(5, 16).Font.Color = RGB(255, 255, 255)
    
    .Cells(5, 17).Value = "Stav n�dr�e"
    .Cells(5, 17).Font.Bold = True
    .Cells(5, 17).HorizontalAlignment = xlCenter
    
    
    
    .Range("A5:P5").Font.Bold = True
    .Range("A5:P5").Interior.Color = RGB(91, 155, 213)
    .Range("A5:P5").Font.Color = RGB(255, 255, 255)
    .Range("A5:P5").HorizontalAlignment = xlCenter
    .Range("A5:P5").VerticalAlignment = xlCenter
    
    .Range("A6:O5").Borders.LineStyle = xlContinuous
    .Range("A6:O5").Borders.Weight = xlThin
    .Range("A5:O5").Borders.Weight = xlMedium
    
    .Activate
    ActiveWindow.DisplayGridlines = False
    .Columns("A:Q").AutoFit
    .Range("A5:P5").AutoFilter
    
    With .PageSetup
        If PRINTORIENTATION = "Landscape" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .PaperSize = xlPaperA4
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .PrintGridlines = False
    End With
End With

Set CreateMonthlySheet = ws

End Function
' Dokon�en� m�s��n�ho listu
Sub FinalizeMonthlySheet(ws As Worksheet, lastRow As Long)
With ws
If lastRow >= 6 Then
.Range("E6:E" & lastRow).NumberFormat = "dd.mm.yyyy"
.Range("G6:G" & lastRow).NumberFormat = "dd.mm.yyyy"
.Range("F6:F" & lastRow).NumberFormat = "HH:MM"
.Range("H6:H" & lastRow).NumberFormat = "HH:MM"
.Range("K6:K" & lastRow).NumberFormat = "#,##0.00"
.Range("L6:L" & lastRow).NumberFormat = "#,##0.00"
.Range("M6:M" & lastRow).NumberFormat = "#,##0.00"
.Range("N6:N" & lastRow).NumberFormat = "#,##0.00"
' Do�asn� �prava (sloupec s aktu�ln�m stavem paliva)
.Range("Q6:Q" & lastRow).NumberFormat = "#,##0.00"



        .Range("A6:O" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A6:O" & lastRow).Borders.Weight = xlThin
        .Range("A6:O" & lastRow).Borders.Color = RGB(166, 166, 166)
        .Range("P6:P" & lastRow).Font.Size = 10
    End If
    
    On Error Resume Next
    With .PageSetup
        .PrintArea = "$A$1:$O$" & lastRow
        If PRINTORIENTATION = "Landscape" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo 0
    
    .Columns("A:P").AutoFit
End With

End Sub
' Vytvo�en� souhrnn�ho listu
Sub CreateSummarySheet(monthlySheets As Object, monthData As Object)
Dim wsSummary As Worksheet
Dim wsP As Worksheet, wsA As Worksheet
Dim prvniDatum As Date, posledniDatum As Date
Dim pocatecniTacho As Double, konecnyTacho As Double
Dim celkemLitry As Double, celkemCena As Double
Dim celkemKm As Double, sluzebniKm As Double, soukromeKm As Double
Dim key As Variant
Dim firstSheet As Boolean
Set wsP = Sheets("N�kup PHM")
Set wsA = Sheets("Auta")

On Error Resume Next
Application.DisplayAlerts = False
Sheets("Kniha j�zd - souhrn").Delete
Application.DisplayAlerts = True
On Error GoTo 0

Set wsSummary = Worksheets.Add
wsSummary.Name = "Kniha j�zd - souhrn"

firstSheet = True
celkemLitry = 0
celkemCena = 0
celkemKm = 0
sluzebniKm = 0
soukromeKm = 0
pocatecniTacho = CDbl(wsA.Cells(2, 3).Value)
konecnyTacho = pocatecniTacho

For Each key In monthlySheets.Keys
    Dim wsMonth As Worksheet
    Set wsMonth = monthlySheets(key)
    Dim lastRowMonth As Long
    lastRowMonth = monthData(key) - 1
    
    If lastRowMonth >= 6 Then
        If firstSheet Then
            prvniDatum = wsMonth.Cells(6, 5).Value
            firstSheet = False
        End If
        posledniDatum = wsMonth.Cells(lastRowMonth, 5).Value
        konecnyTacho = wsMonth.Cells(lastRowMonth, 12).Value
        
        Dim r As Long
        For r = 6 To lastRowMonth
            Dim tripKm As Double, tripType As String
            tripKm = wsMonth.Cells(r, 11).Value
            tripType = wsMonth.Cells(r, 10).Value
            
            celkemKm = celkemKm + tripKm
            celkemLitry = celkemLitry + wsMonth.Cells(r, 13).Value
            celkemCena = celkemCena + wsMonth.Cells(r, 14).Value
            
            If tripType = "soukrom�" Then
                soukromeKm = soukromeKm + tripKm
            Else
                sluzebniKm = sluzebniKm + tripKm
            End If
        Next r
    End If
Next key

With wsSummary

    ' Nadpis
    .Cells(1, 1).Value = "KNIHA J�ZD - SOUHRN"
    .Cells(1, 1).Font.Size = 18
    .Cells(1, 1).Font.Bold = True
    .Cells(1, 1).HorizontalAlignment = xlCenter
    .Range("A1:C1").Merge
    .Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    .Range("A1:C1").Font.Color = RGB(255, 255, 255)
    .Range("A1:C1").RowHeight = 30
    .Range("A1:C1").Borders.Weight = xlThick

    ' Pr�zdn� ��dek
    .Rows(2).RowHeight = 45
    .Range("A2:C2").Merge

    ' Zarovn�n�, p�smo, v��ka ��dk�
    .Range("A3:A15").HorizontalAlignment = xlLeft
    .Range("C3:C15").HorizontalAlignment = xlRight
    .Range("A3:A7").RowHeight = 25
    .Range("A9:A15").RowHeight = 25
    .Range("A3:A15").Font.Bold = True
    .Range("A3:C15").Font.Size = 12
    
    ' Or�mov�n�
    .Range("A3:C7").Borders.LineStyle = xlContinuous
    .Range("A9:C15").Borders.LineStyle = xlContinuous
    .Range("A3:C7").Borders.Weight = xlThin
    .Range("A9:C15").Borders.Weight = xlThin
    
    ' Obdob�
    .Cells(3, 1).Value = "Obdob�:"
    .Cells(3, 3).Value = Format(prvniDatum, "dd.mm.yyyy") & " - " & Format(posledniDatum, "dd.mm.yyyy")
    .Range("A3:C3").Interior.Color = RGB(217, 225, 242)
    .Range("B3:C3").Merge
    
    ' Jm�no a p��jmen�
    .Cells(4, 1).Value = "Jm�no a p��jmen�:"
    .Cells(4, 1).Font.Bold = True
    .Cells(4, 3).Value = wsA.Cells(2, 6).Value
    .Range("B4:C4").Merge
            
    ' Vozidlo
    .Cells(5, 1).Value = "Vozidlo:"
    .Cells(5, 1).Font.Bold = True
    .Cells(5, 1).Font.Size = 12
    .Cells(5, 3).Value = wsA.Cells(2, 1).Value
    .Range("A5:C5").Interior.Color = RGB(217, 225, 242)
    .Range("B5:C5").Merge
    
    ' RZ
    .Cells(6, 1).Value = "RZ:"
    .Cells(6, 1).Font.Bold = True
    .Cells(6, 3).Value = wsA.Cells(2, 4).Value
    .Range("B6:C6").Merge
    
    ' Spot�eba
    .Cells(7, 1).Value = "Spot�eba PHM dle TP (l/100km):"
    .Cells(7, 1).Font.Bold = True
    .Cells(7, 3).Value = wsA.Cells(2, 2).Value
    .Range("A7:C7").Interior.Color = RGB(217, 225, 242)
    .Range("B7:C7").Merge
    
    ' Pr�zdn� ��dek
    .Rows(8).RowHeight = 25
    .Range("A8:C8").Merge
    
    ' Tachometr po��tek
    .Cells(9, 1).Value = "Po��te�n� stav tachometru:"
    .Cells(9, 1).Font.Bold = True
    .Cells(9, 3).Value = pocatecniTacho
    .Cells(9, 3).NumberFormat = "#,##0"
    .Range("B9:C9").Merge
    
    ' Tachometr konec
    .Cells(10, 1).Value = "Kone�n� stav tachometru:"
    .Cells(10, 1).Font.Bold = True
    .Cells(10, 3).Value = konecnyTacho
    .Cells(10, 3).NumberFormat = "#,##0"
    .Range("A10:C10").Interior.Color = RGB(217, 225, 242)
    .Range("B10:C10").Merge
    
    ' Celkem najeto
    .Cells(11, 1).Value = "Celkem najet�ch km:"
    .Cells(11, 1).Font.Bold = True
    .Cells(11, 3).Value = celkemKm
    .Cells(11, 3).NumberFormat = "#,##0.00"
    .Range("B11:C11").Merge
    
    ' Slu�ebn� km
    .Cells(12, 1).Value = "Slu�ebn� km:"
    .Cells(12, 1).Font.Bold = True
    .Cells(12, 3).Value = sluzebniKm
    .Cells(12, 3).NumberFormat = "#,##0.00"
    .Range("A12:C12").Interior.Color = RGB(217, 225, 242)
    .Range("B12:C12").Merge
    
    ' Soukrom� km
    .Cells(13, 1).Value = "Soukrom� km:"
    .Cells(13, 1).Font.Bold = True
    .Cells(13, 3).Value = soukromeKm
    .Cells(13, 3).NumberFormat = "#,##0.00"
    .Range("B13:C13").Merge
    
    ' Na�erp�no paliva
    .Cells(14, 1).Value = "Celkem �erp�n� PHM (l):"
    .Cells(14, 1).Font.Bold = True
    .Cells(14, 3).Value = celkemLitry
    .Cells(14, 3).NumberFormat = "#,##0.00"
    .Range("A14:C14").Interior.Color = RGB(217, 225, 242)
    .Range("B14:C14").Merge
    
    ' Celkem cena za palivo
    .Cells(15, 1).Value = "Celkem cena PHM (K�):"
    .Cells(15, 1).Font.Bold = True
    .Cells(15, 3).Value = celkemCena
    .Cells(15, 3).NumberFormat = "#,##0.00"
    .Range("B15:C15").Merge
    
    ' Pr�zdn� ��dek
    .Cells(16, 1).Value = ""
    .Cells(16, 1).Font.Bold = True
    .Range("A16:C16").RowHeight = 25
    
         
    With .PageSetup
        .PrintArea = "$A$1:$C$15"
        If PRINTORIENTATION = "Landscape" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    .Activate
    ActiveWindow.DisplayGridlines = False
    .Columns("A:C").AutoFit
    .Columns("C").ColumnWidth = 25
End With

End Sub
' Pomocn� funkce
Function GetDistance(ws As Worksheet, fromCol As Long, toCol As Long) As Double
On Error Resume Next
Dim fromRow As Long, rDist As Long
fromRow = 0
For rDist = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
If CStr(ws.Cells(rDist, 1).Value) = CStr(ws.Cells(1, fromCol).Value) Then
fromRow = rDist
Exit For
End If
Next rDist
If fromRow > 0 Then
    GetDistance = CDbl(ws.Cells(fromRow, toCol).Value)
Else
    GetDistance = 0
End If
On Error GoTo 0

End Function
Function GetLocationType(ws As Worksheet, locCol As Long, typCol As Long) As String
On Error Resume Next
Dim locationName As String, rFind As Long
locationName = CStr(ws.Cells(1, locCol).Value)
For rFind = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If CStr(ws.Cells(rFind, 1).Value) = locationName Then
        GetLocationType = CStr(ws.Cells(rFind, typCol).Value)
        Exit Function
    End If
Next rFind
GetLocationType = "Praha"
On Error GoTo 0

End Function
Function GetLocationName(ws As Worksheet, locCol As Long) As String
GetLocationName = CStr(ws.Cells(1, locCol).Value)
End Function
Sub FindColumns(wsL As Worksheet, ByRef typCol As Long, ByRef domovCol As Long, domov As String)
Dim posledniCol As Long, cCol As Long
posledniCol = wsL.Cells(1, wsL.Columns.count).End(xlToLeft).Column

For cCol = 1 To posledniCol
    If UCase(Trim(CStr(wsL.Cells(1, cCol).Value))) = "TYP" Then
        typCol = cCol
        Exit For
    End If
Next cCol

For cCol = 2 To posledniCol
    If cCol <> typCol And CStr(wsL.Cells(1, cCol).Value) = domov Then
        domovCol = cCol
        Exit For
    End If
Next cCol

End Sub
Sub DeleteExistingKnihaJizdSheets()
Dim ws As Worksheet, i As Long
Application.DisplayAlerts = False
For i = Worksheets.count To 1 Step -1
Set ws = Worksheets(i)
If InStr(ws.Name, "Kniha j�zd") > 0 Then
ws.Delete
End If
Next i
Application.DisplayAlerts = True
End Sub
' Vypo�te spot�ebu podle ujet�ch km a spot�eby
Function CalculateFuelUsedInPeriod(actualKm As Double, consumption As Double) As Double
' actualKm = po�et najet�ch kilometr� za obdob�
' consumption = spot�eba l/100 km (glob�ln� prom�nn� Spotreba)
CalculateFuelUsedInPeriod = actualKm * consumption / 100
End Function

' Spo��t� pot�ebnou rezervu paliva na konci obdob� pro p��t� tankov�n�
Function CalculateNeededFuelReserve(currentPeriodKey As String, tankPeriods As Object, wsL As Worksheet, domovCol As Long) As Double
    Dim nextPeriodKey As String
    Dim nextTankLocation As String
    Dim distanceToNextPump As Double
    
    ' Najdi dal�� tankov�n�
    Dim periodKeys() As String
    Dim keyCount As Long
    keyCount = 0
    
    ' Vytvo� pole kl���
    ReDim periodKeys(0 To tankPeriods.count - 1)
    Dim key As Variant
    For Each key In tankPeriods.Keys
        periodKeys(keyCount) = key
        keyCount = keyCount + 1
    Next key
    
    ' Najdi sou�asn� kl�� a dal�� za n�m
    Dim i As Long
    For i = 0 To keyCount - 1
        If periodKeys(i) = currentPeriodKey Then
            If i < keyCount - 1 Then
                nextPeriodKey = periodKeys(i + 1)
                nextTankLocation = tankPeriods(nextPeriodKey)("TankLocation")
                
                ' Spo��tej vzd�lenost k p��t� pump�
                If nextTankLocation = "Praha" Then
                    distanceToNextPump = 0 ' V Praze nemus�me nikam jet
                Else
                    ' Najdi sloupec lokality
                    Dim colIndex As Variant
                    colIndex = Application.Match(nextTankLocation, wsL.Rows(1), 0)
                    If Not IsError(colIndex) Then
                        distanceToNextPump = GetDistance(wsL, domovCol, CLng(colIndex))
                    End If
                End If
                Exit For
            End If
        End If
    Next i
    
    ' Rezerva = palivo na cestu k pump� + 5 litr� bezpe�nostn� rezerva
    CalculateNeededFuelReserve = (distanceToNextPump * 15 / 100) + 5
End Function

