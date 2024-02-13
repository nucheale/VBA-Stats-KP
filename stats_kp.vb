'i.zabolotny@spb-neo.ru

Function getMaxTwoDimArrayValue(arr) As Double
    maxValue = arr(LBound(arr), 1)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) > maxValue Then maxValue = arr(i, 1)
    Next i
    getMaxTwoDimArrayValue = maxValue
End Function

Function getMinTwoDimArrayValue(arr) As Double
    minValue = arr(LBound(arr), 1)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) < minValue Then minValue = arr(i, 1)
    Next i
    getMinTwoDimArrayValue = minValue
End Function

Function removeDublicatesFromTwoDimArray(arr)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not dict.Exists(arr(i, 1)) Then dict.Add arr(i, 1), arr(i, 1)
    Next i
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To dict.Count)
    i = 1
    For Each Key In dict.keys
        uniqueArr(i) = Key
        i = i + 1
    Next Key
    removeDublicatesFromTwoDimArray = uniqueArr
End Function

Function removeDublicatesFromOneDimArray(arr)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr) To UBound(arr)
        If Not dict.Exists(arr(i)) Then dict.Add arr(i), arr(i)
    Next i
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To dict.Count)
    i = 1
    For Each Key In dict.keys
        uniqueArr(i) = Key
        i = i + 1
    Next Key
    removeDublicatesFromOneDimArray = uniqueArr
End Function

Sub Stats()

    Dim e, element, i, j, fileIndex, listKpRow As Long
    
    Set macroWb = ActiveWorkbook
    
    filesToOpen = Application.GetOpenFilename(FileFilter:="All files (*.*), *.*", MultiSelect:=True, Title:="Выберите файлы")
    If TypeName(filesToOpen) = "Boolean" Then Exit Sub
    
    With Application
        .AskToUpdateLinks = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    
    With macroWb.Worksheets("Справочник")
        Dim districts, carriers, files As Variant
        'Dim carriers, files As Variant
        districts = .ListObjects("Районы").DataBodyRange.Value
        carriers = .ListObjects("Перевозчики").DataBodyRange.Value
        files = .ListObjects("Файлы").DataBodyRange.Value
    End With
    
    Set listKpWb = Application.Workbooks.Add

    listKpRow = 0
    fileIndex = 1
    For Each file In filesToOpen
        Set currentWb = Application.Workbooks.Open(Filename:=filesToOpen(fileIndex))
        Select Case True
            Case currentWb.Name Like "*Статистика за*"
                Set statsKpWb = currentWb
            Case currentWb.Name Like "*Отчет по срывам*"
                Set failuresKpWb = currentWb
            Case currentWb.Name Like "*Список КП по участкам*"
                lastRow = currentWb.Sheets(1).Cells.SpecialCells(xlLastCell).Row
                lastColumn = currentWb.Sheets(1).Cells.SpecialCells(xlLastCell).Column
                Set listData = currentWb.Worksheets(1).Range(currentWb.Worksheets(1).Cells(4, 1), currentWb.Worksheets(1).Cells(lastRow, lastColumn))
                listData.Copy Destination:=listKpWb.Sheets(1).Cells(listKpRow + 1, 1)
                listKpRow = listKpWb.Sheets(1).Cells.SpecialCells(xlLastCell).Row
                currentWb.Close SaveChanges:=False
            Case Else
                MsgBox "Неопознанный файл: " & currentWb.Name
                GoTo errorExit
        End Select
        fileIndex = fileIndex + 1
    Next file

    reportDate = CDate(Left(Right(statsKpWb.Name, 15), 10))
    
    With listKpWb.Sheets(1)
        lastRowListKp = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnListKp = .Cells.SpecialCells(xlLastCell).Column
        Set findIDCell = .Range(.Cells(1, 1), .Cells(1, lastColumnListKp)).Find(What:="Код КП", LookAt:=xlWhole)
        Set findDistrictCell = .Range(.Cells(1, 1), .Cells(1, lastColumnListKp)).Find(What:="Район", LookAt:=xlWhole)
        Dim listKpIDList, listKpDistrictsList As Variant
        listKpIDList = .Range(.Cells(findIDCell.Row + 1, findIDCell.Column), .Cells(lastRowListKp, findIDCell.Column))
        listKpDistrictsList = .Range(.Cells(findDistrictCell.Row + 1, findDistrictCell.Column), .Cells(lastRowListKp, findDistrictCell.Column))
        Debug.Print "listKpIDList: " & UBound(listKpIDList)
        Debug.Print "listKpDistrictsList: " & UBound(listKpDistrictsList)
        listKpWb.Close SaveChanges:=False
    End With
    
    With failuresKpWb.Sheets("report")
        lastRowFailures = .Cells(Rows.Count, 3).End(xlUp).Row
        lastColumnFailures = .Cells.SpecialCells(xlLastCell).Column
        Set findIDCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Код КП", LookAt:=xlWhole)
        Set findDistrictCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Участок", LookAt:=xlWhole)
        Set findCarrierCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Перевозчик", LookAt:=xlWhole)
        Set findProblemCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Проблема", LookAt:=xlWhole)
        Dim failuresIDList, failuresDistrictsList, failuresProblemsList As Variant
        failuresIDList = .Range(.Cells(findIDCell.Row + 1, findIDCell.Column), .Cells(lastRowFailures, findIDCell.Column))
        failuresDistrictsList = .Range(.Cells(findDistrictCell.Row + 1, findDistrictCell.Column), .Cells(lastRowFailures, findDistrictCell.Column))
        failuresProblemsList = .Range(.Cells(findProblemCell.Row + 1, findProblemCell.Column), .Cells(lastRowFailures, findProblemCell.Column))
        Dim failuresCarriersList, failuresDistrictsAndCarriersList As Variant
        failuresCarriersList = .Range(.Cells(findCarrierCell.Row + 1, findCarrierCell.Column), .Cells(lastRowFailures, findCarrierCell.Column))
        ReDim failuresDistrictsAndCarriersList(1 To UBound(failuresCarriersList)) As String
        For i = LBound(failuresCarriersList, 1) To UBound(failuresCarriersList, 1)
            failuresDistrictsAndCarriersList(i) = failuresDistrictsList(i, 1) & " | " & failuresCarriersList(i, 1)
        Next i
        Debug.Print "failuresIDList: " & UBound(failuresIDList)
        Debug.Print "failuresDistrictsList: " & UBound(failuresDistrictsList)
        Debug.Print "failuresDistrictsAndCarriersList: " & UBound(failuresDistrictsAndCarriersList)
        Debug.Print "failuresDistrictsAndCarriersList(1): " & failuresDistrictsAndCarriersList(1)
        If Not UBound(failuresIDList) = UBound(failuresDistrictsList) Or Not UBound(failuresDistrictsList) = UBound(failuresDistrictsAndCarriersList) Then
            MsgBox ("Массивы failuresIDList, failuresDistrictsList, failuresDistrictsAndCarriersList не равны по длине")
            GoTo errorExit
        End If
    End With
    
    findedDistricts = 0
    For e = LBound(failuresIDList) To UBound(failuresIDList) 'заполнение района из реестра кп по коду кп
        For n = LBound(listKpIDList) To UBound(listKpIDList)
            If listKpIDList(n, 1) = failuresIDList(e, 1) Then
                failuresDistrictsList(e, 1) = listKpDistrictsList(n, 1)
                findedDistricts = findedDistricts + 1
                Exit For
            Else
                failuresDistrictsList(e, 1) = "КП не найдена"
            End If
        Next n
    Next e
    Debug.Print "findedDistricts: " & findedDistricts
        
    With failuresKpWb.Sheets("report")
        ' Debug.Print "findDistrictCell.Row + 1: " & findDistrictCell.Row + 1
        ' Debug.Print "findDistrictCell.Column: " & findDistrictCell.Column
        .Cells(findDistrictCell.Row + 1, findDistrictCell.Column).Resize(UBound(failuresDistrictsList), UBound(failuresDistrictsList, 2)).Value = failuresDistrictsList 'заполнение района
    End With
    
    findDistrictCell = Empty
    findCarrierCell = Empty
    With statsKpWb.Sheets("Вывоз КП")
        lastRowStatsKp = .Cells(Rows.Count, 2).End(xlUp).Row
        lastColumnStatsKp = .Cells.SpecialCells(xlLastCell).Column
        Set findDistrictCell = .Range(.Cells(1, 1), .Cells(6, lastColumnFailures)).Find(What:="Район", LookAt:=xlWhole)
        Set findCarrierCell = .Range(.Cells(1, 1), .Cells(6, lastColumnFailures)).Find(What:="Перевозчик, в чьи маршруты попали КП из данного МО", LookAt:=xlWhole)
        Dim statsKpDistrictsList, statsKpCarriersList As Variant
        statsKpDistrictsList = .Range(.Cells(findDistrictCell.Row + 2, findDistrictCell.Column), .Cells(lastRowStatsKp, findDistrictCell.Column))
        statsKpCarriersList = .Range(.Cells(findCarrierCell.Row + 2, findCarrierCell.Column), .Cells(lastRowStatsKp, findCarrierCell.Column))
        Debug.Print "statsKpDistrictsList: " & UBound(statsKpDistrictsList)
        Debug.Print "statsKpCarriersList: " & UBound(statsKpCarriersList)
        If Not UBound(statsKpDistrictsList) = UBound(statsKpCarriersList) Then
            MsgBox ("Массивы statsKpDistrictsList, statsKpCarriersList не равны по длине. " & "statsKpDistrictsList: " & UBound(statsKpDistrictsList) & ". statsKpCarriersList: " & UBound(statsKpCarriersList))
            GoTo errorExit
        End If
        Dim statsKpDistrictsAndCarriersList As Variant
        ReDim statsKpDistrictsAndCarriersList(1 To UBound(statsKpDistrictsList)) As String
        For i = LBound(statsKpDistrictsList, 1) To UBound(statsKpDistrictsList, 1)
            statsKpDistrictsAndCarriersList(i) = statsKpDistrictsList(i, 1) & " | " & statsKpCarriersList(i, 1) '299
        Next i
        
        ' Debug.Print "statsKpDistrictsAndCarriersList До: " & UBound(statsKpDistrictsAndCarriersList)
        statsKpDistrictsAndCarriersListClear = removeDublicatesFromOneDimArray(statsKpDistrictsAndCarriersList) '77
        ' Debug.Print "statsKpDistrictsAndCarriersList После: " & UBound(statsKpDistrictsAndCarriersList)
        
        Dim statsWbKpPlan, statsWbKpFact As Variant
        statsWbKpPlan = .Range(.Cells(findDistrictCell.Row + 2, 8), .Cells(lastRowStatsKp, 8))
        statsWbKpFact = .Range(.Cells(findDistrictCell.Row + 2, 10), .Cells(lastRowStatsKp, 10))
    End With
    
    macroWb.Sheets("Шаблон").Copy After:=macroWb.Sheets(macroWb.Sheets.Count - 1)
    Set newWs = ActiveSheet
    newWs.Name = Date & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now)
    With newWs
        Dim macroKpPlan, macroKpFact As Variant
        ReDim macroKpPlan(1 To UBound(statsKpDistrictsAndCarriersListClear))
        ReDim macroKpFact(1 To UBound(statsKpDistrictsAndCarriersListClear))
        For e = LBound(statsKpDistrictsAndCarriersListClear) To UBound(statsKpDistrictsAndCarriersListClear) 'КП из маршрутов без учёта графика, которые попали в план задание и Количество КП из этого МО, попавшее в план задания и есть отчёт
            kpPlan = 0
            kpFact = 0
            For n = LBound(statsKpDistrictsAndCarriersList) To UBound(statsKpDistrictsAndCarriersList)
                If statsKpDistrictsAndCarriersListClear(e) = statsKpDistrictsAndCarriersList(n) Then
                    kpPlan = kpPlan + statsWbKpPlan(n, 1)
                    kpFact = kpFact + statsWbKpFact(n, 1)
                End If
            Next n
            macroKpPlan(e) = kpPlan
            macroKpFact(e) = kpFact
        Next e

        Dim effiency As Variant
        ReDim effiency(1 To UBound(macroKpFact)) As Double
        For i = LBound(effiency) To UBound(effiency)
            If macroKpPlan(i) = 0 And macroKpFact(i) > 0 Then
                effiency(i) = 1
            ElseIf macroKpPlan(i) = 0 And macroKpFact(i) = 0 Then
                effiency(i) = 0
            Else
                effiency(i) = macroKpFact(i) / macroKpPlan(i)
            End If
        Next i

        Dim macroWbDistricts, macroWbCarriers As Variant
        ReDim macroWbDistricts(1 To UBound(statsKpDistrictsAndCarriersListClear))
        ReDim macroWbCarriers(1 To UBound(statsKpDistrictsAndCarriersListClear))
        For i = LBound(statsKpDistrictsAndCarriersListClear) To UBound(statsKpDistrictsAndCarriersListClear)
            splitIndex = InStr(statsKpDistrictsAndCarriersListClear(i), " | ")
            macroWbDistricts(i) = Left(statsKpDistrictsAndCarriersListClear(i), splitIndex - 1)
            macroWbCarriers(i) = Right(statsKpDistrictsAndCarriersListClear(i), Len(statsKpDistrictsAndCarriersListClear(i)) - splitIndex - 2)
        Next i

        Dim macroWbGeneralCarriers As Variant
        ReDim macroWbGeneralCarriers(1 To UBound(macroWbCarriers)) As String
        For i = LBound(macroWbDistricts) To UBound(macroWbDistricts)
            For e = LBound(districts, 1) To UBound(districts, 1)
                If macroWbDistricts(i) = districts(e, 1) Then macroWbGeneralCarriers(i) = districts(e, 2)
            Next e
        Next i
        
        Debug.Print carriers(2, 2)
        

        Debug.Print UBound(failuresProblemsList)
        Dim problems As Variant
        ReDim problems(1 To UBound(statsKpDistrictsAndCarriersListClear))
        For e = LBound(statsKpDistrictsAndCarriersListClear) To UBound(statsKpDistrictsAndCarriersListClear)
            problem = ""
            counter = 1
            For n = LBound(failuresDistrictsAndCarriersList) To UBound(failuresDistrictsAndCarriersList)
                If statsKpDistrictsAndCarriersListClear(e) = failuresDistrictsAndCarriersList(n) Then
                    If InStr(problem, failuresProblemsList(n, 1)) = 0 Then
                        problem = problem & counter & ". " & failuresProblemsList(n, 1) & vbLf
                        counter = counter + 1
                    End If
                End If
            Next n
            If Right(problem, 1) = vbLf Then problem = Left(problem, Len(problem) - 1) Else If problem = "" Then problem = "–"
            problems(e) = problem
        Next e
        
        .Cells(1, 1) = "Отчет за " & reportDate
        ' .Cells(3, 1).Resize(UBound(statsKpDistrictsAndCarriersListClear)).Value = Application.Transpose(statsKpDistrictsAndCarriersListClear)
        .Cells(3, 1).Resize(UBound(macroWbDistricts)).Value = Application.Transpose(macroWbDistricts)
        .Cells(3, 2).Resize(UBound(macroWbGeneralCarriers)).Value = Application.Transpose(macroWbGeneralCarriers)
        .Cells(3, 3).Resize(UBound(macroWbCarriers)).Value = Application.Transpose(macroWbCarriers)
        .Cells(3, 4).Resize(UBound(macroKpPlan)).Value = Application.Transpose(macroKpPlan)
        .Cells(3, 5).Resize(UBound(macroKpFact)).Value = Application.Transpose(macroKpFact)
        .Cells(3, 6).Resize(UBound(effiency)).Value = Application.Transpose(effiency)
        .Cells(3, 7).Resize(UBound(problems)).Value = Application.Transpose(problems)
        lastRowMacroWb = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnMacroWb = .Cells.SpecialCells(xlLastCell).Column
        .Range(.Cells(2, 1), .Cells(lastRowMacroWb, lastColumnMacroWb)).Borders.LineStyle = xlContinuous
    End With

    statsKpWb.Close SaveChanges:=False
    failuresKpWb.Close SaveChanges:=False
    
errorExit:
    With Application
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


