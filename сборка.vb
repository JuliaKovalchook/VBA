Sub Consolidated_Range_of_Books_and_Sheets()
    Dim iBeginRange As Object, lCalc As Long, lCol As Long
    Dim oAwb        As String, sCopyAddress As String, sSheetName As String
    Dim lLastrow    As Long, lLastRowMyBook As Long, li As Long, iLastColumn As Integer
    Dim wsSh        As Object, wsDataSheet As Object, bPolyBooks As Boolean, avFiles
    Dim wbAct       As Workbook
    Dim bPasteValues As Boolean
    On Error Resume Next
    'Выбираем диапазон выборки с книг
    Set iBeginRange = Application.InputBox("Выберите диапазон сбора данных." & vbCrLf & _
    "1. При выборе только одной ячейки данные будут собраны со всех листов начиная с этой ячейки. " & _
    vbCrLf & "2. При выделении нескольких ячеек данные будут собраны только с указанного диапазона всех листов.", Type:=8)
    'для указания диапазона без диалогового окна:
    'Set iBeginRange = Range("A1:A10") 'диапазон указывается нужный
    'Если диапазон не выбран - завершаем процедуру
    If iBeginRange Is Nothing Then Exit Sub
    'Указываем имя листа
    'Допустимо указывать в имени листа символы подставки ? и *.
    'Если указать только * то данные будут собираться со всех листов
    sSheetName = InputBox("Введите имя листа, с которого собирать данные(если не указан, то данные собираются со всех листов)", "Параметр")
    'Если имя листа не указано - данные будут собраны со вех листов
    If sSheetName = "" Then sSheetName = "*"
    On Error GoTo 0
    'Запрос - вставлять на результирующий лист все данные
    'или только значения ячеек (без формул и форматов)
    bPasteValues = (MsgBox("Вставлять только значения?", vbQuestion + vbYesNo, "Excel-VBA") = vbYes)
    'Запрос сбора данных с книг(если Нет - то сбор идет с активной книги)
    If MsgBox("Собрать данные с нескольких книг?", vbInformation + vbYesNo, "Excel-VBA.ru") = vbYes Then
        avFiles = Application.GetOpenFilename("Excel files(*.xls*),*.xls*", , "Выбор файлов", , True)
        If VarType(avFiles) = vbBoolean Then Exit Sub
        bPolyBooks = TRUE
        lCol = 1
    Else
        avFiles = Array(ThisWorkbook.FullName)
    End If
    'отключаем обновление экрана, автопересчет формул и отслеживание событий
    'для скорости выполнения кода и для избежания ошибок, если в книгах есть иные коды
    With Application
        lCalc = .Calculation
        .ScreenUpdating = False: .EnableEvents = False: .Calculation = xlManual
    End With
    'создаем новый лист в книге для сбора
    Set wsDataSheet = ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    'если нужно сделать сбор данных на новый лист книги с кодом
    'Set wsDataSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    'цикл по книгам
    For li = LBound(avFiles) To UBound(avFiles)
        If bPolyBooks Then
            Set wbAct = Workbooks.Open(Filename:=avFiles(li))
        Else
            Set wbAct = ThisWorkbook
        End If
        oAwb = wbAct.Name
        'цикл по листам
        For Each wsSh In wbAct.Sheets
            If wsSh.Name Like sSheetName Then
                'Если имя листа совпадает с именем листа, в который собираем данные
                'и сбор идет только с активной книги - то переходим к следующему листу
                If wsSh.Name = wsDataSheet.Name And bPolyBooks = FALSE Then GoTo NEXT_
                With wsSh
                    Select Case iBeginRange.Count
                        Case 1        'собираем данные начиная с указанной ячейки и до конца данных
                            lLastrow = .Cells(1, 1).SpecialCells(xlLastCell).Row
                            iLastColumn = .Cells.SpecialCells(xlLastCell).Column
                            sCopyAddress = .Range(.Cells(iBeginRange.Row, iBeginRange.Column), .Cells(lLastrow, iLastColumn)).Address
                        Case Else        'собираем данные с фиксированного диапазона
                            sCopyAddress = iBeginRange.Address
                    End Select
                    lLastRowMyBook = wsDataSheet.Cells.SpecialCells(xlLastCell).Row + 1
                    'вставляем имя книги, с которой собраны данные
                    If lCol Then wsDataSheet.Cells(lLastRowMyBook, 1).Resize(Range(sCopyAddress).Rows.Count).Value = oAwb
                    If bPasteValues Then        'если вставляем только значения
                    .Range(sCopyAddress).Copy
                    wsDataSheet.Cells(lLastRowMyBook, 1).Offset(, lCol).PasteSpecial xlPasteValues
                Else
                    .Range(sCopyAddress).Copy wsDataSheet.Cells(lLastRowMyBook, 1).Offset(, lCol)
                End If
            End With
        End If
        NEXT_:
    Next wsSh
    If bPolyBooks Then wbAct.Close FALSE
Next li

With Application
    .ScreenUpdating = True: .EnableEvents = True: .Calculation = lCalc
End With
End Sub