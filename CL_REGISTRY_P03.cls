VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_REGISTRY_P03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'          Класс создания приложения к акту
'          освидетельствования скрытых работ №3
'=============================================================


'Столбцы
Private Const L_PP$ = "A"
Private Const L_Mat$ = "B"
Private Const L_Doc$ = "C"
Private Const L_Akt$ = "D"

'Настройка таблицы
Private Const start_row As Byte = 5

'Переменные класса
Private mass_data()
Private dict_data As Object
Private dict_ic As Object
Private SH As Excel.Worksheet
Private OtherFunction As New CL_OTHERFUN
Private coll_function As New Collection   'Коллекция функциональных обязанностей
Private coll_emploee As New Collection    'Коллекция Фамилия И.О. подписантов


'Конструктор и деструктор класса
Private Sub Class_Initialize()
    
    coll_function.Add "Представитель застройщика, технического заказчика, лица, ответственного за эксплуатацию здания, сооружения, или регионального оператора по вопросам строительного контроля"
    coll_function.Add "Представитель лица, осуществляющего строительство, реконструкцию, капитальный ремонт"
    coll_function.Add "Представитель лица, осуществляющего строительство, реконструкцию, капитальный ремонт, по вопросам строительного контроля"
    
    coll_emploee.Add "Ефимчук А.А."
    'coll_emploee.Add "Жигарев П.В."
    coll_emploee.Add "Васильев О.А."
    'coll_emploee.Add "Перевощиков Д.А."
    coll_emploee.Add "Середа А.А."
    
    'Массив данных таблицы
    Set dict_data = CreateObject("Scripting.Dictionary")
    Dim THWB As New CL_THISWORKBOOK
    mass_data = THWB.MassData
    Set THWB = Nothing
    
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
    Set dict_ic = Nothing
    Set SH = Nothing
    Set OtherFunction = Nothing
    Set coll_function = Nothing
    Set coll_emploee = Nothing
End Sub

'Номера столбцов таблицы
Private Property Get cNumPP%()
    cNumPP = Columns(L_PP).Column
End Property

'Установка данных входного контроля
Private Property Set DictDataIC(ByRef dict As Object)
    Set dict_ic = dict
End Property

'Последняя строка данных в таблице
Private Property Get cnt_row&()
    cnt_row = SH.Cells(Rows.Count, cNumPP).End(xlUp).Row
End Property

'Подготовка списка для загрузки входного контроля и список материалов
Public Function GetDownLoadedList(ByRef sFilter$) As Object
    Dim q&, sKey$, dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add key:="ListTitle", Item:=CreateObject("Scripting.Dictionary")
    dict.Add key:="ListDP", Item:=CreateObject("Scripting.Dictionary")
    
    For q = LBound(mass_data) To UBound(mass_data)
        
        If IsError(mass_data(q, idx_aosr_num)) Then GoTo continue
        
        sKey = Trim(mass_data(q, idx_aosr_num))
        If sKey <> sFilter Then GoTo continue
        
        'Список титулов
        sKey = mass_data(q, idx_title)
        If Not dict("ListTitle").exists(sKey) Then dict("ListTitle").Add key:=sKey, Item:=Empty
        
        'Список dp
        sKey = mass_data(q, idx_dp)
        If Not dict("ListDP").exists(sKey) Then dict("ListDP").Add key:=sKey, Item:=Empty
        
        'Список элементов для реестра
        sKey = GetKeyElement(q)
        If Not dict_data.exists(sKey) Then dict_data.Add key:=sKey, Item:=tmpdict(q)
        
continue:
    Next
    
    dict("ListTitle") = dict("ListTitle").keys
    dict("ListDP") = dict("ListDP").keys
    
    Set GetDownLoadedList = dict
    Set dict = Nothing
    
End Function
Private Function tmpdict(q) As Object
    Set tmpdict = CreateObject("Scripting.Dictionary")
    With tmpdict
        .Add key:=.Count, Item:=mass_data(q, idx_name)    'Наименование элемента
        .Add key:=.Count, Item:=mass_data(q, idx_mark)    'Марка элемента
        .Add key:=.Count, Item:=mass_data(q, idx_pos)     'Счетчик элемента
        .Add key:=.Count, Item:=mass_data(q, idx_wght)    'Вес элемента
    End With
End Function

'Сопоставление данных ведомости с данными входного контроля
Public Sub SearchInputControl(ByRef dict_ic As Object)
    Dim q&, iKey, dict As Object
    
    For Each iKey In dict_data
        
        If Not dict_ic.exists(iKey) Then GoTo continue
            
        Set dict = dict_ic(iKey)
        With dict_data(iKey)
            .Add key:=.Count, Item:=dict(idx_ic_aorpi_name)
            .Add key:=.Count, Item:=dict(idx_ic_aorpi_date)
            .Add key:=.Count, Item:=dict(idx_ic_sert_name)
            .Add key:=.Count, Item:=dict(idx_ic_sert_date)
        End With

continue:
    Next
    
End Sub

'Создание документа
Public Sub CreateDocument(ByRef sFilter$, ByRef mass(), ByRef sFile$)
    
    'Создание листа если его нет
    Set SH = OtherFunction.CreateSheet(sFilter & "_АОСР_п3")
    CreateHeader sFilter    'Создание шапки таблицы
    
    With SH.Range(L_PP & start_row).Resize(UBound(mass), UBound(mass, 2))
        OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignCenter, "Times New Roman", 14, Wrap:=True
        .Value = mass
        
        'Форматирование созданной таблицы
        .Columns(L_Mat).Font.Italic = True
        .Columns(L_Doc).Font.Italic = True
        .Columns(L_Akt).Font.Italic = True
        
        .FormatConditions.Add(xlTextString, _
        TextOperator:=xlDoesNotContain, String:="").Interior.Color = rgbRed
        
    End With
    
    CreateFooter    'Создание подвала подписантов
    
    OtherFunction.BoldersCells SH.Range(L_PP & start_row).CurrentRegion, xlThin
    PrinterConfiguration
    
    SH.Move
    ActiveWorkbook.SaveAs Filename:=sFile, FileFormat:=XlFileFormat.xlOpenXMLWorkbook
    
    
End Sub

'Массив вставляемой таблицы
Public Function GetMassRegistry()
    GetMassRegistry = GetMassData(dict_table)
End Function

'Формирование массива данных для вставки таблицы Реестр к п.3
Private Function dict_table() As Object
    Dim i, sKey$, dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each i In dict_data
        sKey = dict_data(i)(0)
        sKey = sKey & dict_data(i)(1)
        sKey = sKey & dict_data(i)(4)
        
        If Not dict.exists(sKey) Then dict.Add key:=sKey, Item:=CreateObject("Scripting.Dictionary") Else GoTo continue
        With dict(sKey)
            .Add key:=.Count, Item:=Empty
            .Add key:=.Count, Item:=dict_data(i)(0) & " " & dict_data(i)(1)
            .Add key:=.Count, Item:=GetNameDocQuality(dict_data(i)(6) & " от " & dict_data(i)(7) & "г.")
            .Add key:=.Count, Item:=dict_data(i)(4) & Chr(10) & "от " & dict_data(i)(5) & "г."
        End With
continue:
    Next
    
    Set dict_table = dict
    Set dict = Nothing
    
End Function

'Преобразование словаря в массив
Private Function GetMassData(ByRef dict As Object)
    Dim q&, w&, i
    ReDim mass(1 To dict.Count, 1 To 4)
    For Each i In dict.items
        q = q + 1
        mass(q, 1) = q
        For w = LBound(i.items) + 1 To UBound(i.items)
            mass(q, w + 1) = i(w)
        Next
    Next
    GetMassData = mass
End Function

'хрен знает поиск серии номера или б/н
Private Function GetNameDocQuality$(ByRef sTXT$)
    Dim q&
    q = InStr(1, sTXT, sSerial(sTXT), vbTextCompare)
    If q = 1 Then GetNameDocQuality = sTXT: Exit Function
    GetNameDocQuality = Trim(Left(sTXT, q)) & Chr(10) & Trim(Mid(sTXT, q))
End Function

'Регулярное выражение поиск серия или № и до конца строки
Private Function sSerial$(ByRef sTXT$)
    Dim objExec As Object
    With CreateObject("VBScript.RegExp")
        .Global = True 'Поиск по всему тексту
        .IgnoreCase = True
        .Pattern = "\s+(серия|№|б/н).*" 'Шаблон поиска
        Set objExec = .Execute(sTXT)
    End With
    If objExec.Count = 1 Then sSerial = objExec.Item(0).Value
End Function

'Подвал таблицы
Private Sub CreateFooter()
    Dim q&, addr$, iRow&
    
    iRow = cnt_row + 2
    
    With SH
        
        For q = 1 To coll_emploee.Count
            iRow = iRow + 1
            With .Range("A" & iRow & ":D" & iRow)
                OtherFunction.FormatCells .Offset, xlHAlignLeft, xlVAlignBottom, "Times New Roman", 12, Merge:=True, Bold:=True, Wrap:=True
                .RowHeight = 30
                .Value = coll_function.Item(q)
            End With
            
            iRow = iRow + 1
            With .Range("A" & iRow & ":D" & iRow)
                OtherFunction.FormatCells .Offset, xlHAlignLeft, xlVAlignBottom, "Times New Roman", 12, Merge:=True, Bold:=False
                .RowHeight = 20
                .Font.Italic = True
                .Value = coll_emploee.Item(q)
            End With
            rows_UnderLine iRow
        Next
        
    End With
End Sub

'Подпись под строкой
Private Sub rows_UnderLine(ByRef iRow&)
    Dim addr$
    
    With SH
        iRow = iRow + 1
        With .Range("A" & iRow & ":D" & iRow)
            OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignTop, "Times New Roman", 8, Merge:=True, Bold:=True, Wrap:=True
            .RowHeight = 21
            .Font.Italic = True
            .Borders(xlEdgeTop).Weight = xlThin
            .Value = "(фамилия, инициалы, подпись)"
        End With
    End With
    
End Sub

'Ключ для словаря с элементами титул, dp, марка, наименование, позиция, вес
Private Property Get GetKeyElement$(ByRef q&)
    Dim sKey$
    sKey = mass_data(q, idx_title) & vbTab
    sKey = sKey & mass_data(q, idx_dp) & vbTab
    sKey = sKey & mass_data(q, idx_mark) & vbTab
    sKey = sKey & mass_data(q, idx_name) & vbTab
    sKey = sKey & mass_data(q, idx_pos) & vbTab
    sKey = sKey & mass_data(q, idx_wght)
    GetKeyElement = sKey
End Property

'Формирование названий столбцов
Private Function dict_header() As Object
    Set dict_header = CreateObject("Scripting.Dictionary")
    With dict_header
        .Add key:=.Count + 1, Item:="№" & Chr(10) & "п/п"
        .Add key:=.Count + 1, Item:="Наименование материала"
        .Add key:=.Count + 1, Item:="Наименование документов, подтверждающих качество и" & Chr(10) & _
                                        "безопасность строительных материалов (изделий), номер и дата"
        .Add key:=.Count + 1, Item:="№ акта о результатах проверки изделий на" & Chr(10) & _
                                        "соответствие техдокументации, дата"
    End With
End Function

'Шапка таблицы Приложение к акту освидетельствования скрытых работ
Private Sub CreateHeader(ByRef sFilter$)
    Dim addr$, tmpstr$
    
    'Ширина столбцов
    OtherFunction.ColumnsWidth SH, 1, 8, 35, 85, 55
    
    With SH
    
        'Высоты строк
        .Rows("1:2").RowHeight = 20
        .Rows(3).RowHeight = 12
        .Rows(4).RowHeight = 40
        
        'Заголовоки документа
        With .Range("A1:D1")
            OtherFunction.FormatCells .Offset, xlCenter, xlBottom, "Times New Roman", 14, Merge:=True, Bold:=True
            .Value = "Реестр к п.4"
        End With
        With .Range("A2:D2")
            OtherFunction.FormatCells .Offset, xlCenter, xlBottom, "Times New Roman", 14, Merge:=True, Bold:=False
            tmpstr = "Приложение к акту освидетельствования ответственных конструкций " & sFilter
            tmpstr = tmpstr & " от «" & String(4, "_") & "»" & String(12, "_") & "20___г."
            .Value = tmpstr
        End With
        
        'Заголовки таблицы
        With .Range("A4:D4")
            OtherFunction.FormatCells .Offset, xlCenter, xlCenter, "Times New Roman", 14, Bold:=True
            .Value = dict_header.items
        End With
        
    End With
    
End Sub

'настройки принтинга листа
Private Sub PrinterConfiguration()

    Application.PrintCommunication = False
    
    With SH.PageSetup
        
        '.PrintTitleColumns = ""
        '.PrintArea = ""
        .PrintTitleRows = "4:4"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Страница &P из &N"
        .LeftMargin = Application.InchesToPoints(1)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.8)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    
    Application.PrintCommunication = True
    
End Sub
