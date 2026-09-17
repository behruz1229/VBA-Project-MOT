VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_GEODETIC_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Настройки таблицы геодезического контроля
Private Const mask_file$ = "реестр геодезия"
Private Const sFolder$ = "C:\Users\SobyaninAV\Desktop"

Private Const sheet_name$ = "Реестр"    'Имя листа на которой таблица
Private Const start_row As Byte = 38    '

'Столбцы таблицы геодезия сводная:
Private Const L_Title$ = "B" 'Титул
Private Const L_DP$ = "D"    'DP
Private Const L_Wght$ = "L"  'Масса



'Переменные класса:
Private sFileName$
Private mass_data()
Private dict_data As Object
Private WB As Excel.Workbook


'Конструктор и деструктор класса
Private Sub Class_Initialize()
    Set dict_data = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
End Sub


'Номера столбцов таблицы
Private Property Get cNumTitle%()
    cNumTitle = Columns(L_Title).Column
End Property
Private Property Get cNumDP%()
    cNumDP = Columns(L_DP).Column
End Property
Private Property Get cNumWght%()
    cNumWght = Columns(L_Wght).Column
End Property



'Словарь сводной таблицы
Public Property Get DictData() As Object
    Set DictData = dict_data
End Property

'Загрузка данных в словарь данных из таблицы мастер файла
Public Sub LoadData()
    
    Dim OPWB As New CL_OPENFILE
    OPWB.SetMaskFile = mask_file
    OPWB.ShowFileDialog = True
    OPWB.ListFolder.Add sFolder
    
    sFileName = OPWB.GetFileName
    If sFileName = "" Then Err.Raise 513, mask_file
    
    Set WB = OPWB.GetWorkBook
    
    Application.StatusBar = "Загрузка данных из таблицы по геодезии..."
    Debug.Print Application.StatusBar: DoEvents
    
    'LoadDataInMass
    LoadDataInDict
    'CalculateValueInDP
    
    
    Set WB = Nothing
    Set OPWB = Nothing
    Application.StatusBar = False
    
    
End Sub

'Чтение данных из таблицы
Private Sub LoadDataInMass()
    mass_data = WB.Worksheets(sheet_name).UsedRange.Value
End Sub


'Загрузка данных в словарь
Private Sub LoadDataInDict()
    Dim sTitle$, sDP$, wght#, RG As Range, dict As Object
    
    
    With WB.Worksheets(sheet_name)
                
        For Each RG In .Range(L_DP & start_row, .Cells(Rows.Count, cNumDP).End(xlUp)).Rows
            
            If Not .Range(L_Title & RG.Row).Value = Empty Then
                sTitle = GetTitle(.Range(L_Title & RG.Row).Value)
                GoTo continue
            End If
            
            If Not .Range(L_DP & RG.Row).Value = Empty Then
                
                If sTitle = "" Then GoTo continue
                
                sDP = Replace(.Range(L_DP & RG.Row).Value, Chr(32), "")
                wght = 0
                
                If IsNumeric(.Range(L_Wght & RG.Row).Value) Then
                    wght = .Range(L_Wght & RG.Row).Value
                End If
                
                If wght = 0 Then GoTo continue
                
                If Not dict_data.exists(sTitle) Then dict_data.Add key:=sTitle, Item:=CreateObject("Scripting.Dictionary")
                Set dict = dict_data(sTitle)
                
                If Not dict.exists(sDP) Then dict.Add key:=sDP, Item:=wght
                
                GoTo continue
            End If
            
continue:
        Next
    
    End With
    
End Sub
Private Sub LoadDpData(ByRef RG As Range)
    
End Sub

'Поиск даты в строке материала
Private Property Get GetTitle$(ByRef sTXT)
    Dim objExec As Object, tmpstr$
    Set objExec = RegExp(sTXT, "[A-Z]\d-\d{2}-?[\d\w]*")
    If objExec.Count = 1 Then
        tmpstr = objExec(0).Value
        GetTitle = tmpstr
    End If
    Set objExec = Nothing
End Property

'Свойство создания (настроек) объекта регулярных выражений
Private Property Get RegExp(ByRef sTXT, ByRef sMask$) As Object
    With CreateObject("VBScript.RegExp")
        .Global = True 'Поиск по всему тексту
        .Pattern = sMask 'Шаблон поиска
        .IgnoreCase = True
        Set RegExp = .Execute(sTXT)
    End With
End Property




'
''Расчет значений по геодезии
'Private Sub CalculateValueInDP()
'    Dim iTitle, iDP
'    For Each iTitle In dictDP.Items
'        For Each iDP In iTitle.Items
'            With iDP
'                .Item("G") = GetComletedPhisic(iDP)
'                .Item("F") = GetWorkIsUnderway(iDP)
'            End With
'        Next
'    Next
'End Sub
'
''Ведется СМР '"F" => "=D5-G5-ЕСЛИ(H5="н/п";D5;H5)"
'Private Property Get GetWorkIsUnderway#(ByRef dict)
'    Dim val#
'    With dict
'        val = .Item("D") - .Item("G") - .Item("H")
'        GetWorkIsUnderway = IIf(val < 0, 0, val)
'    End With
'End Property
'
''Готовая физика под геодезию '"G" => "=ЕСЛИ(D5/C5>0,8;D5-ЕСЛИ(H5="н/п";D5;H5);0)"
'Private Property Get isDelivDivMont(ByRef dict) As Boolean
'    With dict
'        If .Item("C") = 0 Then Exit Property
'        If .Item("D") / .Item("C") > 0.8 Then isDelivDivMont = True
'    End With
'End Property
'Private Property Get GetComletedPhisic#(ByRef dict)
'    Dim val#
'    If isDelivDivMont(dict) Then
'        With dict
'            val = .Item("D") - .Item("H")
'            GetComletedPhisic = IIf(val < 0, 0, val)
'        End With
'    Else
'        GetComletedPhisic = 0
'    End If
'End Property
'
''Данные весов из таблицы
'Private Property Get GetGeoSurvey#(ByRef mass, ByRef q&)
'    GetGeoSurvey = mass(q, NumGeoSurvey) * 1000
'End Property
'Private Property Get GetGeoPnt#(ByRef mass, ByRef q&)
'    GetGeoPnt = mass(q, NumGeoPnt) * 1000
'End Property
'Private Property Get GetGeoAgrd#(ByRef mass, ByRef q&)
'    GetGeoAgrd = mass(q, NumGeoAgrd) * 1000
'End Property
'Private Property Get GetGeoAgrmnt#(ByRef mass, ByRef q&)
'    GetGeoAgrmnt = mass(q, NumGeoAgrmnt) * 1000
'End Property
