VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_PIVOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================
'Класс для создания сводной таблицы по элементам
'Все настройки сводной таблицы находятся здесь
'======================================================


'Настройки таблицы
Private Const start_row As Byte = 4
Private Const cnt_columns% = 21

Private Const color_total& = XlRgbColor.rgbLightGray 'Цвет строки ИТОГО

'Индексы столбцов в словаре для сводной
Public Enum pivot_col
    idx_pv_title    'Титул
    idx_pv_all      'Всего
    idx_pv_dlvr     'Поставлено
    idx_pv_mount    'Вес, тонн/Смонт.
    idx_pv_remain   'Вес, тонн/Остаток
    
    idx_pv_geo_incomesmr    'Ведется СМР
    idx_pv_geo_readphis     'Готовая физика под геодезию (съемка)
    idx_pv_geo_survey       'Геодезическая съемка
    idx_pv_geo_isdrawn      'ИГС отрисована
    idx_pv_geo_underapp     'ИГС на согласовании
    idx_pv_geo_agreed       'ИГС согласованы
    
    idx_pv_insp_vk_nodvlr   'Отсутствует RFI ВК на поставленные м/к
    idx_pv_insp_vk_yesmount 'RFI ВК проведен на смонтиро-ванные м/к
    idx_pv_insp_vk_nomount  'Отсутствует RFI ВК на смонтиро-ванные м/к
    
    idx_pv_aorpi_yes    'ИГС на согласовании
    idx_pv_aorpi_no     'ИГС согласованы
    
    idx_pv_rfi_app  'Заявлено по RFI (монтаж МК)
    idx_pv_rfi_acc  'Принято по RFI (монтаж МК)
    
    idx_pv_id_reg   'ИД на оформлении
    idx_pv_id_app   'ИД на согласовании
    idx_pv_id_acc   'ИД согласовано
    
End Enum


'Переменные класса
Private wght#
Private sheet_row&
Private mass_tb()
Private mass_data()
Private mass_sum_total#(1 To 20)
Private dict_mf As Object
Private dict_ic As Object
Private OtherFunction As New CL_OTHERFUN

'Конструктор и деструктор класса
Private Sub Class_Terminate()
    Set dict_mf = Nothing
    Set dict_ic = Nothing
    Set OtherFunction = Nothing
    Erase mass_tb
End Sub


'Объект рабочего листа с таблицей
Public Property Get SH_PIVOT() As Excel.Worksheet
    Set SH_PIVOT = Лист2
End Property

'Словарь с данными для сводной таблицы
Public Property Set DictDataMF(ByRef dict As Object)
    Set dict_mf = dict
End Property

'Словарь с данными входного контроля
Public Property Set DictDataIC(ByRef dict As Object)
    Set dict_ic = dict
End Property

'Массив данных таблицы ведомости элементов
Public Property Let MassDataTB(ByRef mass)
    mass_tb = mass
End Property


'Основная процедура обновления таблицы
Public Sub UpdateTable()
    
    Application.StatusBar = "Формирование сводной"
    Debug.Print Application.StatusBar: DoEvents
    
    Application.ScreenUpdating = False
    
    MakeDictPivot   'Формирование словаря с данными
    SortData        'Сортировка данных
    
    CreateTabtle    'Создание заголовка
    InsertData      'Вставка собранных данных
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub

'Сортировка данных
Private Sub SortData()
    Dim SRT As New CL_SORT
    SRT.SetSortOrder = set_ASC
    Set dict_mf = SRT.Sort(dict_mf)
    Set SRT = Nothing
End Sub

'Вставка собранных данных
Private Sub InsertData()
    Dim iKey, cnt_dps&
    
    sheet_row = start_row
    
    With SH_PIVOT
        
        For Each iKey In dict_mf
            
            cnt_dps = dict_mf(iKey).Count + 1
            ReDim mass_data(1 To cnt_dps, 1 To 21)
            
            mass_data(1, 1) = iKey  'Имя листа из мастер файла
            
            ins_DpName dict_mf(iKey)
            
            With .Range("A" & sheet_row).Resize(cnt_dps, 21)
                .Value = mass_data  'Вставка собранных данных
                FormatTable .Offset 'Форматирование части таблицы
                OtherFunction.BoldersCells .Offset, xlThin
            End With
            
            'Группирование, скрытие строк
            With .Rows(sheet_row + 1 & ":" & sheet_row + cnt_dps - 1)
                .Group
                .EntireRow.Hidden = True
            End With
            
            sheet_row = sheet_row + cnt_dps
            
        Next
    
        'Итого
        .Range("A" & sheet_row).Value = "ИТОГО"
        .Range("B" & sheet_row).Resize(1, UBound(mass_sum_total)).Value = mass_sum_total 'Вставка данных ИТОГО
        With .Range("A" & sheet_row).Resize(1, UBound(mass_sum_total) + 1)
            .Range("A1").HorizontalAlignment = xlRight
            .Font.Italic = True
            .Font.Bold = True
            .Interior.Color = color_total
            OtherFunction.BoldersCells .Offset, xlThin
        End With
    
    End With
    
    MiniTable sheet_row     'Вставка мини-таблицы для тупых
    Signature sheet_row     'Вставка поля подписантов в сводной
    
End Sub

'Вставка имени dp
Private Sub ins_DpName(ByRef dict As Object)
    Dim iKey, idx, q&
    q = 1
    For Each iKey In dict
        q = q + 1
        mass_data(q, 1) = iKey
        ins_Value q, dict(iKey)
    Next
End Sub

'Заполнение массиов данными весов материалов
Private Sub ins_Value(ByRef q&, ByRef dict As Object)
    Dim idx
    For Each idx In dict
        wght = WorksheetFunction.Round(dict(idx) / 1000, 3) 'Вес материала, округ. до 3 знаков
        mass_data(q, idx + 1) = wght
        mass_data(1, idx + 1) = mass_data(1, idx + 1) + wght
        ins_TotalSheetValue idx
    Next
End Sub

'Расчет массива данных ИТОГО
Private Sub ins_TotalSheetValue(ByRef idx)
    mass_sum_total(idx) = mass_sum_total(idx) + wght
End Sub


'Формирование словаря с данными
Private Sub MakeDictPivot()
    Dim q&, sKey$, i_ic As Object, i_mf As Object
    
    For q = LBound(mass_tb) To UBound(mass_tb)
        
        'Словарь сводной таблицы
        '1. Титул (лист из мастер файла)
        sKey = Trim(mass_tb(q, idx_sheet))
        If dict_mf.exists(sKey) Then Set i_mf = dict_mf(sKey) Else GoTo continue
            
        '2. номер DP
        sKey = Trim(mass_tb(q, idx_dp))
        If sKey = "" Then sKey = "NO_DP"
        If i_mf.exists(sKey) Then Set i_mf = i_mf(sKey) Else GoTo continue
        
        wght = mass_tb(q, idx_wght)
        
        'Данные входного контроля
        sKey = GetKey(q)
        If Not dict_ic Is Nothing Then
            If dict_ic.exists(sKey) Then
                Set i_ic = dict_ic(sKey)
                calc_pv_ic q, i_ic, i_mf
            End If
        End If
        
        'Данные по исполнительной документации
        calc_pv_id q, i_mf
        
continue:
    Next
    
End Sub

'Расчеты сводной таблицы в части исполнительной документации
Private Sub calc_pv_id(ByRef q&, ByRef i_mf As Object)
    
    'Заявлено по RFI (монтаж МК)
    If isInspSmuEx(q) Then
        i_mf(idx_pv_rfi_app) = i_mf(idx_pv_rfi_app) + wght
    Else
        i_mf(idx_pv_rfi_app) = i_mf(idx_pv_rfi_app) + 0
    End If
    
    'Принято по RFI (монтаж МК)
    If isInspAccept(q) Then
        i_mf(idx_pv_rfi_acc) = i_mf(idx_pv_rfi_acc) + wght
    Else
        i_mf(idx_pv_rfi_acc) = i_mf(idx_pv_rfi_acc) + 0
    End If
    
    'ИД на оформлении
    If isIdRegistration(q) Then
        i_mf(idx_pv_id_reg) = i_mf(idx_pv_id_reg) + wght
    Else
        i_mf(idx_pv_id_reg) = i_mf(idx_pv_id_reg) + 0
    End If
    
    'ИД на согласовании
    If isIdApplication(q) Then
        i_mf(idx_pv_id_app) = i_mf(idx_pv_id_app) + wght
    Else
        i_mf(idx_pv_id_app) = i_mf(idx_pv_id_app) + 0
    End If
    
    'ИД согласовано
    If isIdAcception(q) Then
        i_mf(idx_pv_id_acc) = i_mf(idx_pv_id_acc) + wght
    Else
        i_mf(idx_pv_id_acc) = i_mf(idx_pv_id_acc) + 0
    End If
    
End Sub


'Расчеты сводной таблицы в части входного контроля
Private Sub calc_pv_ic(ByRef q&, ByRef i_ic As Object, ByRef i_mf As Object)
    
    'Отсутствует RFI ВК на поставленные м/к
    If Not isInspExists(i_ic) Then
        i_mf(idx_pv_insp_vk_nodvlr) = i_mf(idx_pv_insp_vk_nodvlr) + wght
    Else
        i_mf(idx_pv_insp_vk_nodvlr) = i_mf(idx_pv_insp_vk_nodvlr) + 0
    End If
    
    'RFI ВК проведен на смонтиро-ванные м/к
    If isInspExists(i_ic) And isMount(q) Then
        i_mf(idx_pv_insp_vk_yesmount) = i_mf(idx_pv_insp_vk_yesmount) + wght
    Else
        i_mf(idx_pv_insp_vk_yesmount) = i_mf(idx_pv_insp_vk_yesmount) + 0
    End If
    
    'Отсутствует RFI ВК на смонтиро-ванные м/к
    If Not isInspExists(i_ic) And isMount(q) Then
        i_mf(idx_pv_insp_vk_nomount) = i_mf(idx_pv_insp_vk_nomount) + wght
    Else
        i_mf(idx_pv_insp_vk_nomount) = i_mf(idx_pv_insp_vk_nomount) + 0
    End If
    
    'АоРПИ на смонтиро-ванные м/к
    If isAorpiExists(i_ic) And isMount(q) Then
        i_mf(idx_pv_aorpi_yes) = i_mf(idx_pv_aorpi_yes) + wght
    Else
        i_mf(idx_pv_aorpi_yes) = i_mf(idx_pv_aorpi_yes) + 0
    End If
    
    'Отсутствует АоРПИ на смонтиро-ванные м/к
    If Not isAorpiExists(i_ic) And isMount(q) Then
        i_mf(idx_pv_aorpi_no) = i_mf(idx_pv_aorpi_no) + wght
    Else
        i_mf(idx_pv_aorpi_no) = i_mf(idx_pv_aorpi_no) + 0
    End If
    
End Sub

'Булевые вычисления:
Private Property Get isMount(ByRef q&) As Boolean           'элемент смонтирован
    isMount = IsDate(mass_tb(q, idx_date))
End Property
Private Property Get isInspExists(ByRef i_ic) As Boolean    'есть наименование инспекции rfi
    isInspExists = i_ic(idx_ic_insp) <> Empty
End Property
Private Property Get isAorpiExists(ByRef i_ic) As Boolean   'есть наименование аорпи
    isAorpiExists = i_ic(idx_ic_aorpi_name) <> Empty
End Property
Private Property Get isInspSmuEx(ByRef q&) As Boolean      'Инспекция по монтажу заявлена
    isInspSmuEx = mass_tb(q, idx_insp_smu) <> Empty
End Property
Private Property Get isInspAccept(ByRef q&) As Boolean      'rfi монтажа принята
    isInspAccept = InStr(1, mass_tb(q, idx_insp_stat_sk), "принят", vbTextCompare) = 1
End Property
Private Property Get isIdRegistration(ByRef q&)             'ИД на оформлении
    isIdRegistration = InStr(1, mass_tb(q, idx_aosr_stat), "оформлен", vbTextCompare) = 1
End Property
Private Property Get isIdApplication(ByRef q&)              'ИД на согласовании
    isIdApplication = InStr(1, mass_tb(q, idx_aosr_stat), "на проверк", vbTextCompare) = 1
End Property
Private Property Get isIdAcception(ByRef q&)                'ИД согласовано
    isIdAcception = InStr(1, mass_tb(q, idx_aosr_stat), "согласован", vbTextCompare) = 1
End Property


'Ключ: Титул, DP, Марка, Имя, Позиция, Вес
Private Property Get GetKey$(ByRef q&)
    Dim sKey$
    sKey = mass_tb(q, idx_title) & vbTab
    sKey = sKey & mass_tb(q, idx_dp) & vbTab
    sKey = sKey & mass_tb(q, idx_mark) & vbTab
    sKey = sKey & mass_tb(q, idx_name) & vbTab
    sKey = sKey & mass_tb(q, idx_pos) & vbTab
    sKey = sKey & mass_tb(q, idx_wght)
    GetKey = sKey
End Property


'Создание заголовка
Private Sub CreateTabtle()
    
    SH_PIVOT.Cells.Delete 'УДАЛЕНИЕ ДАННЫХ СВОДНОЙ ТАБЛИЦЫ
    
    'Ширина столбцов
    OtherFunction.ColumnsWidth SH_PIVOT, 1, 40, 12, 12, 11, 11, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13
    CreateHeader SH_PIVOT 'Имена столбцов
        
    With SH_PIVOT
        
        'Форматирование заголовков
        With .Range("A2", .Cells(3, cnt_columns))
            OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignCenter, "Arial", 9, Bold:=True, Wrap:=True
            OtherFunction.BoldersCells .Offset, xlThin
        End With
        
        .Rows("2:3").RowHeight = 25 'Высота строк
        
        'Вставка примечаний
        With .Range("S2")
            FormatComment .AddComment("Статус АОСР" & Chr(10) & """Оформление""")
        End With
        
        With .Range("T2")
            FormatComment .AddComment("Статус АОСР" & Chr(10) & """На проверке""")
        End With
        
        With .Range("U2")
            FormatComment .AddComment("Статус АОСР" & Chr(10) & """Согласован""")
        End With
        
        'Сводка + вчерашняя дата в ячейке "A1"
        With .Range("A1")
            OtherFunction.FormatCells .Offset, xlLeft, xlCenter, "Arial", 12, True, False, False
            .Offset.Font.Italic = True
            .Value = "Сводка " & Format(Date - 1, "yyyy.mm.dd")
        End With
        
    End With

End Sub

'Словарь с наименованиями столбцов и адресов ячеек
Private Function dict_header() As Object
    Set dict_header = CreateObject("Scripting.Dictionary")
    With dict_header
        .Add key:="Титул", Item:="A2:A3"
        .Add key:="Всего", Item:="B2:B3"
        .Add key:="Поставлено", Item:="C2:C3"
        .Add key:="Вес, тонн", Item:="D2:E2"
        .Add key:="Смонт.", Item:="D3"
        .Add key:="Остаток", Item:="E3"
        
        .Add key:="Ведется СМР", Item:="F2:F3"
        .Add key:="Готовая физика под геодезию (съемка)", Item:="G2:G3"
        .Add key:="Геодезическая съемка", Item:="H2:H3"
        .Add key:="ИГС отрисована", Item:="I2:I3"
        .Add key:="ИГС на согласовании", Item:="J2:J3"
        .Add key:="ИГС согласованы", Item:="K2:K3"
        
        .Add key:="Отсутствует RFI ВК на поставленные м/к", Item:="L2:L3"
        .Add key:="RFI ВК проведен на смонтиро-ванные м/к", Item:="M2:M3"
        .Add key:="Отсутствует RFI ВК на смонтиро-ванные м/к", Item:="N2:N3"
        
        .Add key:="АоРПИ на смонтиро-ванные м/к", Item:="O2:O3"
        .Add key:="Отсутствует АоРПИ на смонтиро-ванные м/к", Item:="P2:P3"
        .Add key:="Заявлено по RFI (монтаж МК)", Item:="Q2:Q3"
        .Add key:="Принято по RFI (монтаж МК)", Item:="R2:R3"
        
        .Add key:="ИД на оформлении", Item:="S2:S3"
        .Add key:="ИД на согласовании", Item:="T2:T3"
        .Add key:="ИД согласовано", Item:="U2:U3"
    End With
End Function

'Вставка имен столбцов сводной по адресам из значений словаря
Private Sub CreateHeader(ByRef SH As Excel.Worksheet)
    Dim iKey, addr$, dict As Object
    Set dict = dict_header
    For Each iKey In dict
        addr = dict(iKey)
        With SH.Range(addr)
            If .Count > 1 Then .Merge
            .Value = iKey
        End With
    Next
End Sub

'Мини таблица с итогами для тупых которые не могут посмотреть в строку ИТОГО
Private Sub MiniTable(ByRef cnt_row&)

    Dim addr$, dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    cnt_row = cnt_row + 3
    
    With dict
        .Add key:=.Count, Item:="Всего смонтированно:"
        .Add key:=.Count, Item:="Комплекс смр (не завершен смр для съемки):"
        .Add key:=.Count, Item:="Готово под геодезию (съемка):"
        .Add key:=.Count, Item:="Выполнена геодезия (съемка):"
        .Add key:=.Count, Item:="Геодезия не требуется:"
        
        .Add key:=.Count, Item:="Отрисовано схем (накопительно):"
        .Add key:=.Count, Item:="Требуется согласование геодезии:"
        .Add key:=.Count, Item:="ИГС согласованы:"
        
        .Add key:=.Count, Item:="Заявлено RFI (монтаж):"
        .Add key:=.Count, Item:="Сдано под RFI (монтаж):"
        .Add key:=.Count, Item:="Отсутствует RFI ВК на смонтированные м/к:"
        .Add key:=.Count, Item:="Отсутствует ВК (АоРПИ):"
        
    End With
    
    With SH_PIVOT
        
        addr = "B" & cnt_row
        OtherFunction.FormatCells .Range(addr), xlCenter, xlCenter, "Times New Roman", 11, Bold:=False
        .Range(addr).Value = "объем в т."
        cnt_row = cnt_row + 1
        
        addr = "A" & cnt_row & ":A" & cnt_row + dict.Count - 1
        OtherFunction.FormatCells .Range(addr), xlLeft, xlCenter, "Times New Roman", 11, Bold:=False
        .Range(addr).Value = WorksheetFunction.Transpose(dict.items)
        
        addr = "B" & cnt_row & ":B" & cnt_row + dict.Count - 1
        OtherFunction.FormatCells .Range(addr), xlCenter, xlCenter, "Arial", 11, Bold:=True
        
        'Всего смонтированно:
        With .Range(addr).Cells(1)
            .Value = mass_sum_total(3) 'Порядковый номер итоговой строки (Итого=1,Всего=2, и т.д.)
            ColoRowsPivot .Offset, rgbAliceBlue 'Цвет выделения ячеек
        End With
        
        'Комплекс смр (не завершен смр для съемки):
        With .Range(addr).Cells(2)
            .Value = mass_sum_total(5)
            ColoRowsPivot .Offset, rgbAliceBlue
        End With
        
        'Готово под геодезию (съемка):
        With .Range(addr).Cells(3)
            .Value = mass_sum_total(6)
            ColoRowsPivot .Offset, rgbAliceBlue
        End With
        
        'Выполнена геодезия (съемка):
        With .Range(addr).Cells(4)
            .Value = mass_sum_total(7)
            ColoRowsPivot .Offset, rgbAliceBlue
        End With
        
        'Геодезия не требуется:
        With .Range(addr).Cells(5)
            .EntireRow.Hidden = True    'Скрыть строку
            .Value = 0
            ColoRowsPivot .Offset, rgbBeige
        End With
        
        'Отрисовано схем (накопительно):
        With .Range(addr).Cells(6)
            .Value = mass_sum_total(8)
            ColoRowsPivot .Offset, rgbBeige
        End With
        
        'Требуется согласование геодезии:
        With .Range(addr).Cells(7)
            .Value = mass_sum_total(9)
            ColoRowsPivot .Offset, rgbBeige
        End With
        
        'ИГС согласованы:
        With .Range(addr).Cells(8)
            .Value = mass_sum_total(10)
            ColoRowsPivot .Offset, rgbBeige
        End With
        
        'Сдано под RFI (монтаж):
        With .Range(addr).Cells(9)
            .Value = mass_sum_total(16)
            ColoRowsPivot .Offset, rgbLightGray
        End With
        
        'Заявлено RFI (монтаж):
        With .Range(addr).Cells(10)
            .Value = mass_sum_total(17)
            ColoRowsPivot .Offset, rgbLightGray
        End With
        
        
        'Отсутствует RFI ВК на смонтированные м/к:
        With .Range(addr).Cells(11)
            .Value = mass_sum_total(13)
            ColoRowsPivot .Offset, rgbLightGray
        End With
        
        'Отсутствует ВК (АоРПИ):
        With .Range(addr).Cells(12)
            .Value = mass_sum_total(15)
            ColoRowsPivot .Offset, rgbLightGray
        End With
        
        OtherFunction.BoldersCells .Range("A" & cnt_row).CurrentRegion, xlThin
    End With
    
End Sub

'Закрашивание строк в мини табличке под сводной
Private Sub ColoRowsPivot(ByRef RG_value As Range, ByRef i_color)
    With RG_value
        .Interior.Color = i_color
        .Offset(, -1).Interior.Color = i_color
    End With
End Sub
Private Sub Signature(ByRef cnt_row)
    Dim iRow&, addr$
    iRow = cnt_row + 11
    addr = "J" & iRow & ":" & "U" & iRow
    With SH_PIVOT.Range(addr)
        OtherFunction.FormatCells .Offset, Merge:=True
        .Value = "/"
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
End Sub

'Функция форматирование части таблицы выружаемой сводной
Private Sub FormatTable(ByRef RG As Range)
    With RG
        .Columns("A").HorizontalAlignment = xlHAlignRight
        .Range("A1").HorizontalAlignment = xlHAlignLeft
        .Rows(1).Font.Bold = True
    End With
End Sub

'Функция форматирования комментариев
Private Sub FormatComment(ByRef objComm As Excel.Comment)
    With objComm.Shape
        .Width = 100
        .Height = 40
        .TextFrame.Characters.Font.Size = 14
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
    End With
End Sub

'Настройки принтинга листа
Private Sub PrinterConfiguration(ByRef SH As Worksheet)

    Application.PrintCommunication = False
    
    With SH.PageSetup
        .PrintArea = ""
        .PrintTitleRows = "$2:$3"
        .PrintTitleColumns = ""
        .PrintArea = ""
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "Страница &P из &N"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0.2)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA3
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        '.FitToPagesWide = 1
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
    
    Const txt$ = "&""Yu Gothic UI Light,обычный"" &7 &K08+059 Отчет от &D_&T --> ©" & "Автор макроса: Антон Валерьевич Собянин, редактор: Саидахмедов Бехруз Баходурович email: behruz1229@gmail.com, whatsapp: +7 (996) 322 00 24"
    SH.PageSetup.RightHeader = txt
    On Error GoTo en
    Application.PrintCommunication = True
    
en:
End Sub
