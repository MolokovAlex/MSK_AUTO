'==================================================================================================================================
'================================================= Module =============================================================

' ---------------------- столбцы группы Спецификация --------------------------------------
Public Const Offset_Colunms_Group_specf             As Integer = 2 'столбцы группа Спецификации
Public Const Size_Offset_Colunms_Group_specf        As Integer = 3 'длина группы 
' ---------------------- столбцы группы Производителя --------------------------------------
Public Const Offset_Colunms_Group_producer          As Integer = 5 'столбцы группа Производитель
Public Const Size_Offset_Colunms_Group_producer     As Integer = 4 'длина группы 
' ---------------------- столбцы группы ГШВА --------------------------------------
Public Const Offset_Rows_GSHVA_Summ                 As Integer = 5 'строки группа ГШВА 
Public Const Size_Rows_GSHVA_Summ                   As Integer = 18 'длина группы 
' ---------------------- 1 квартал ---------------------------------------
Public Const Offset_Column_need_plan_1kvartal              As Integer = 10    ' потребность-план 1 квартала 2025г                   
Public Const Column_initial_warehouse_balance1kv    As Integer = 11    ' столбец начальный складской остаток на 1 квартал 2025г
Public Const Column_plan_1kvartal                   As Integer = 12    ' столбец план реализации 1 квартала 2025г
Public Const Column_need_1kvartal                   As Integer = 13    ' потребность 1 квартала 2025г              
Public Const Column_buy_1kvartal                    As Integer = 14    ' в закупку 1 квартал 2025г
Public Const Column_outgo_1kvartal                  As Integer = 15    ' расход 1 квартал 2025г
Public Const Offset_Column_final_warehouse_balance1kv      As Integer = 16    ' конечный складской остаток 1 квартал 2025г
Public Const Offset_Columns_Group_1kvartal     As Integer = 17    ' начало группы стобцов месяцев 1 кв.2025 + группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г
Public Const Size_Columns_Group_1kvartal     As Integer = 13    ' длина групп Columns_Group_1kvartal



'==================================================================================================================================
' Функция скрывает столбец или группу столбцов
' In:
' Start_num_column - начальный номер группы столбцов
' Size_column - количество столбцов для скрытия
Public Function Entire_Column(ByVal Start_num_column As Integer, ByVal Size_column As Integer) As Boolean 
    Dim I As Integer
    For I = Start_num_column To Start_num_column+Size_column-1
        Range("A1").Offset(rowOffset:=0, columnOffset:=I).EntireColumn.Hidden = True 
    Next I
    Entire_Column = True
End Function
'==================================================================================================================================
'==================================================================================================================================
' Функция скрывает строку  или группу строк
' In:
' Start_num_row - начальный номер группы стстрок
' Size_row - количество строк для скрытия
Public Function Entire_Row(ByVal Start_num_row As Integer, ByVal Size_row As Integer) As Boolean 
    Dim I As Integer
    For I = Start_num_row To Start_num_row+Size_row-1
        Range("A1").Offset(rowOffset:=I, columnOffset:=0).EntireRow.Hidden = True 
    Next I
    Entire_Row = True
End Function




'==================================================================================================================================
'================================================ End Module =============================================================
'==================================================================================================================================






'==================================================================================================================================
Private Sub ExpandAll()
'UpdatebyExtendoffice20181031
    Dim I As Integer
    Dim J As Integer
   
    On Error Resume Next
    For I = 1 To 100
        Worksheets("Детализация").Outline.ShowLevels rowLevels:=I
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next I
    For J = 1 To 100
        Worksheets("Детализация").Outline.ShowLevels columnLevels:=J
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next J
End Sub
'==================================================================================================================================

Private Function UngroupAll()
    Dim J As Integer
   
    On Error Resume Next
    For J = 1 To 6
        Range("A1:DM1").Ungroup
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next J

    Range("1:700").Ungroup
End Function

'==================================================================================================================================

Private Function InitTable()
' разворачивание всех группированных столбцов
' Worksheets("Детализация").Outline.ShowLevels 6, 6
' ActiveSheet.Cells.ClearOutline
ExpandAll

' Покажем все скрытые строки и столбцы
Range("A1:DM1").EntireColumn.Hidden = False
Range("1:700").EntireRow.Hidden = False 

' отменим группировку всех столбцов
' Range("A1:DM1").Ungroup
' Range("A1:DM1").Ungroup
UngroupAll

End Function
'==================================================================================================================================



















'==================================================================================================================================
'============================================================     Режим Большой таблицы     ========================================
'==================================================================================================================================
'Макрос ВКЛючения режима отображеия Большой таблицы - без группировок столбцов
' Форматирует таблицу 
'in:
' текущий лист книги
'==================================================================================================================================
Sub MSK_ON_BigTable()

Dim password As String
Dim good_password As String
Dim MODE_BIG_TABLE As String
MODE_BIG_TABLE = "Режим Большой таблицы"



' If Range("A1").Value = MODE_BIG_TABLE Then
'     MsgBox "Режим Большой таблицы уже применен. Макрос не запущен."
'     Exit Sub
' End If

' введем защиту от случайного запуска макроса - нужно ввести пароль
good_password = "987"
password = InputBox("Введите пароль на запуск макроса (987):")
If Not IsNumeric(password) Then
    MsgBox "Пароль не верен - не число. Макрос не запущен."
    Exit Sub
End If
If password <> good_password Then
    MsgBox "Пароль не верен - не _верное_ число. Макрос не запущен."
    Exit Sub
End If



InitTable


' ---------------------- столбцы группы Спецификация --------------------------------------
' set Colunms_Group_specf                 = Range("C1:D1") 'столбцы группа Спецификации
' ---------------------- столбцы группы Производителя --------------------------------------
' set Colunms_Group_producer              = Range("F1:H1") 'столбцы группа Производитель
' группируем группы Спецификации и Производитель
' Colunms_Group_specf.Group
' Colunms_Group_producer.Group

' закроем(свернем) все группы
' Worksheets("Детализация").Outline.ShowLevels 1, 1

Range("A1").Value = MODE_BIG_TABLE
MsgBox "Работа макроса закончена"
End Sub





















'==================================================================================================================================
'============================================================     Режим БДДС       ========================================
'==================================================================================================================================
'Макрос ВКЛючения режима отображеия для среза БДДС
' Форматирует таблицу для анализа БДДС
'in:
' текущий лист книги
'==================================================================================================================================

Sub MSK_ON_BDDS()

Dim password As String
Dim good_password As String

Dim MODE_BDDS As String
MODE_BDDS = "Режим БДДС"

If Range("A1").Value = MODE_BDDS Then
    MsgBox "Режим БДДС уже применен. Макрос не запущен."
    Exit Sub
End If

' введем защиту от случайного запуска макроса - нужно ввести пароль
good_password = "987"
password = InputBox("Введите пароль на запуск макроса (987):")
If Not IsNumeric(password) Then
    MsgBox "Пароль не верен - не число. Макрос не запущен."
    Exit Sub
End If
If password <> good_password Then
    MsgBox "Пароль не верен - не _верное_ число. Макрос не запущен."
    Exit Sub
End If

' ----------------------------------- Назначение столбцов ---------------------------------------------------------------
' TODO -  вывести эти диапазоны как константы и сделать функции скрытия столбцов, пердавая туда константы как архив
' и тогда не надо в каждой функции дублировать эти  SET-ы
' ---------------------- столбцы группы Спецификация --------------------------------------
set Colunms_Group_specf                 = Range("C1:D1") 'столбцы группа Спецификации
' ---------------------- столбцы группы Производителя --------------------------------------
set Colunms_Group_producer              = Range("F1:H1") 'столбцы группа Производитель

' ---------------------- все кварталы ---------------------------------------

set Columns_Group_2kvartal              = Range("AL1:AW1") 'столбцы группа Аванс/Ок.расчет/Примечание для 2 квартала 2025г
set Columns_Group_3kvartal              = Range("BF1:BQ1") 'столбцы группа Аванс/Ок.расчет/Примечание для 3 квартала 2025г
set Columns_Group_4kvartal              = Range("BZ1:CK1") 'столбцы группа Аванс/Ок.расчет/Примечание для 4 квартала 2025г


' ---------------------- 1 квартал ---------------------------------------
set Column_need_plan_1kvartal           = Range("K1")  ' потребность-план 1 квартала 2025г
set Column_initial_warehouse_balance1kv = Range("L1")  ' столбец начальный складской остаток на 1 квартал 2025г
set Column_plan_1kvartal                = Range("M1")  ' столбец план реализации 1 квартала 2025г
set Column_need_1kvartal                = Range("N1")  ' потребность 1 квартала 2025г
set Column_buy_1kvartal                 = Range("O1")  ' в закупку 1 квартал 2025г
set Column_outgo_1kvartal               = Range("P1")  ' расход 1 квартал 2025г
set Column_final_warehouse_balance1kv   = Range("Q1")  ' конечный складской остаток 1 квартал 2025г
set Columns_jan_1kvartal                = Range("R1:T1") 'группа столбцов январь
set Columns_feb_1kvartal                = Range("V1:X1") 'группа столбцов февраль
set Columns_march_1kvartal              = Range("Z1:AB1") 'группа столбцов март
set Columns_Group_1kvartal              = Range("R1:AC1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г

' ---------------------- 2 квартал ---------------------------------------
set Column_need_plan_2kvartal           = Range("AE1")  ' потребность-план 2 квартала 2025г
set Column_initial_warehouse_balance2kv = Range("AF1")  ' столбец начальный складской остаток на 2 квартал 2025г
set Column_plan_2kvartal                = Range("AG1")  ' столбец план реализации 2 квартала 2025г
set Column_need_2kvartal                = Range("AH1")  ' потребность 2 квартала 2025г
set Column_buy_2kvartal                 = Range("AI1")  ' в закупку 2 квартал 2025г
set Column_outgo_2kvartal               = Range("AJ1")  ' расход 2 квартал 2025г
set Column_final_warehouse_balance2kv   = Range("AK1")  ' конечный складской остаток 2 квартал 2025г
set Columns_aprl_2kvartal               = Range("AL1:AN1") 'группа столбцов апрель
set Columns_may_2kvartal                = Range("AP1:AR1") 'группа столбцов май
set Columns_june_2kvartal               = Range("AT1:AV1") 'группа столбцов июнь
set Columns_Group_2kvartal              = Range("AL1:AW1") 'столбцы группа Аванс/Ок.расчет/Примечание для 2 квартала 2025г

' ---------------------- 3 квартал ---------------------------------------
set Column_need_plan_3kvartal           = Range("AY1")  ' потребность-план 3 квартала 2025г
set Column_initial_warehouse_balance3kv = Range("AZ1")  ' столбец начальный складской остаток на 3 квартал 2025г
set Column_plan_3kvartal                = Range("BA1")  ' столбец план реализации 3 квартала 2025г
set Column_need_3kvartal                = Range("BB1")  ' потребность 3 квартала 2025г
set Column_buy_3kvartal                 = Range("BC1")  ' в закупку 3 квартал 2025г
set Column_outgo_3kvartal               = Range("BD1")  ' расход 3 квартал 2025г
set Column_final_warehouse_balance3kv   = Range("BE1")  ' конечный складской остаток 3 квартал 2025г
set Columns_jule_3kvartal               = Range("BF1:BH1") 'группа столбцов июль
set Columns_augst_3kvartal              = Range("BJ1:BL1") 'группа столбцов август
set Columns_sept_3kvartal               = Range("BN1:BP1") 'группа столбцов сентябрь
set Columns_Group_3kvartal              = Range("BF1:BQ1") 'столбцы группа Аванс/Ок.расчет/Примечание для 3 квартала 2025г

' ---------------------- 4 квартал ---------------------------------------
set Column_need_plan_4kvartal           = Range("BS1")  ' потребность-план 4 квартала 2025г
set Column_initial_warehouse_balance4kv = Range("BT1")  ' столбец начальный складской остаток на 4 квартал 2025г
set Column_plan_4kvartal                = Range("BU1")  ' столбец план реализации 4 квартала 2025г
set Column_need_4kvartal                = Range("BV1")  ' потребность 4 квартала 2025г
set Column_buy_4kvartal                 = Range("BW1")  ' в закупку 4 квартал 2025г
set Column_outgo_4kvartal               = Range("BX1")  ' расход 4 квартал 2025г
set Column_final_warehouse_balance4kv   = Range("BY1")  ' конечный складской остаток 4 квартал 2025г
set Columns_oktr_4kvartal               = Range("BZ1:CB1") 'группа столбцов октябрь
set Columns_nov_4kvartal                = Range("CD1:CF1") 'группа столбцов ноябрь
set Columns_dec_4kvartal                = Range("CH1:CJ1") 'группа столбцов декабрь
set Columns_Group_4kvartal              = Range("BZ1:CK1") 'столбцы группа Аванс/Ок.расчет/Примечание для 4 квартала 2025г

' ---------------------- 1 квартал следующего года ---------------------------------------
set Column_need_plan_next_kvartal           = Range("CM1")  ' потребность-план 1 квартала 2026г
set Column_initial_warehouse_balance_next_kv = Range("CN1")  ' столбец начальный складской остаток на 1 квартал 2026г
set Column_plan_next_kvartal                = Range("CO1")  ' столбец план реализации 1 квартала 2026г
set Column_need_next_kvartal                = Range("CP1")  ' потребность 1 квартала 2026г
set Column_buy_next_kvartal                 = Range("CQ1")  ' в закупку 1 квартал 2026г
set Column_outgo_next_kvartal               = Range("CR1")  ' расход 1 квартал 2026г
set Column_final_warehouse_balance_next_kv  = Range("CS1")  ' конечный складской остаток 1 квартал 2026г
set Columns_jan_next_kvartal                = Range("CT1:CV1") 'группа столбцов январь
set Columns_feb_next_kvartal                = Range("CX1:CZ1") 'группа столбцов февраль
set Columns_march_next_kvartal              = Range("DB1:DD1") 'группа столбцов март
set Columns_Group_next_kvartal              = Range("CT1:DE1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г

' -----------------------------------
' ----------------------------------- Назначение строк ---------------------------------------------------------------
set Rows_GSHVA                              = Range("6:23") ' все строки от ГШВА



' разворачивание всех группированных столбцов
' ActiveSheet.Outline.ShowLevels ColumnLevels:=1
' Worksheets("Детализация").Outline.ShowLevels 6, 6
ExpandAll

Range("A1:DM1").EntireColumn.Hidden = False

' отменим группировку всех столбцов
' Range("A1:DM1").Ungroup
' Range("A1:DM1").Ungroup
UngroupAll

' группируем группы Спецификации и Производитель
Colunms_Group_specf.Group
Colunms_Group_producer.Group

Range("J1").EntireColumn.Hidden = True
Range("DG1:DM1").EntireColumn.Hidden = True

' сгруппируем группы месяцев 1 квартала и группы месяцев 1 квартала
Columns_jan_1kvartal.Group
Columns_feb_1kvartal.Group
Columns_march_1kvartal.Group
Columns_Group_1kvartal.Group
' Скрываем ненужные для анализа БДДС столбцы
Column_need_plan_1kvartal.EntireColumn.Hidden = True          
Column_initial_warehouse_balance1kv.EntireColumn.Hidden = True
Column_plan_1kvartal.EntireColumn.Hidden = True               
Column_need_1kvartal.EntireColumn.Hidden = True               
Column_outgo_1kvartal.EntireColumn.Hidden = True              
Column_final_warehouse_balance1kv.EntireColumn.Hidden = True  


' сгруппируем группы месяцев 2 квартала и группы месяцев 2 квартала
Columns_aprl_2kvartal.Group 
Columns_may_2kvartal.Group  
Columns_june_2kvartal.Group 
Columns_Group_2kvartal.Group
' Скрываем ненужные для анализа БДДС столбцы
Column_need_plan_2kvartal.EntireColumn.Hidden = True           
Column_initial_warehouse_balance2kv.EntireColumn.Hidden = True 
Column_plan_2kvartal.EntireColumn.Hidden = True                
Column_need_2kvartal.EntireColumn.Hidden = True                
Column_outgo_2kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance2kv.EntireColumn.Hidden = True   


' сгруппируем группы месяцев 3 квартала и группы месяцев 3 квартала
Columns_jule_3kvartal.Group 
Columns_augst_3kvartal.Group
Columns_sept_3kvartal.Group 
Columns_Group_3kvartal.Group
' Скрываем ненужные для анализа БДДС столбцы
Column_need_plan_3kvartal.EntireColumn.Hidden = True           
Column_initial_warehouse_balance3kv.EntireColumn.Hidden = True 
Column_plan_3kvartal.EntireColumn.Hidden = True                
Column_need_3kvartal.EntireColumn.Hidden = True                
Column_outgo_3kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance3kv.EntireColumn.Hidden = True   


' сгруппируем группы месяцев 4 квартала и группы месяцев 4 квартала
Columns_oktr_4kvartal.Group 
Columns_nov_4kvartal.Group  
Columns_dec_4kvartal.Group  
Columns_Group_4kvartal.Group
' Скрываем ненужные для анализа БДДС столбцы
Column_need_plan_4kvartal.EntireColumn.Hidden = True           
Column_initial_warehouse_balance4kv.EntireColumn.Hidden = True 
Column_plan_4kvartal.EntireColumn.Hidden = True                
Column_need_4kvartal.EntireColumn.Hidden = True                
Column_outgo_4kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance4kv.EntireColumn.Hidden = True   


' сгруппируем группы месяцев 1 квартала след года и группы месяцев 1 квартала след года
Columns_jan_next_kvartal.Group               
Columns_feb_next_kvartal.Group                
Columns_march_next_kvartal.Group              
Columns_Group_next_kvartal.Group  
' Скрываем ненужные для анализа БДДС столбцы
Column_need_plan_next_kvartal.EntireColumn.Hidden = True            
Column_initial_warehouse_balance_next_kv.EntireColumn.Hidden = True 
Column_plan_next_kvartal.EntireColumn.Hidden = True                 
Column_need_next_kvartal.EntireColumn.Hidden = True                  
Column_outgo_next_kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance_next_kv.EntireColumn.Hidden = True   
            
'скрытие строк
Rows("6:10").Hidden = True
Rows("12:16").Hidden = True
Rows("25:28").Hidden = True
Rows("31:32").Hidden = True
Rows("38:46").Hidden = True
Rows("50:52").Hidden = True
Rows("54").Hidden = True
Rows("56:57").Hidden = True
Rows("60:68").Hidden = True
Rows("77").Hidden = True
Rows("86:89").Hidden = True
Rows("95").Hidden = True
Rows("98:99").Hidden = True
Rows("104:105").Hidden = True
Rows("108:109").Hidden = True
Rows("112").Hidden = True
Rows("116:118").Hidden = True
Rows("128").Hidden = True
Rows("130:132").Hidden = True
Rows("136:137").Hidden = True
Rows("139").Hidden = True
Rows("141").Hidden = True
Rows("147").Hidden = True
Rows("149").Hidden = True
Rows("155").Hidden = True
Rows("157:158").Hidden = True
Rows("164:166").Hidden = True
Rows("181:184").Hidden = True
Rows("189:190").Hidden = True
Rows("192").Hidden = True
Rows("196").Hidden = True
Rows("198:200").Hidden = True

' закроем(свернем) все группы
Worksheets("Детализация").Outline.ShowLevels 1, 1

' TODO - подумать как высвобождать память от SET-ов и нужно ли это вообще
' высвобождаем память на вякий случай
' set Columns_Group_1kvartal              = Nothing
' set Columns_Group_2kvartal              = Nothing
' set Columns_Group_3kvartal              = Nothing
' set Columns_Group_4kvartal              = Nothing
' set Column_need_plan_1kvartal           = Nothing
' set Column_initial_warehouse_balance    = Nothing
' set Column_plan_1kvartal                = Nothing
' set Column_need_1kvartal                = Nothing
' set Columns_01_1kvartal                 = Nothing
' set Column_buy_1kvartal                 = Nothing
' set Column_outgo_1kvartal               = Nothing
' set Column_final_warehouse_balance      = Nothing
' set Columns_02_1kvartal                 = Nothing

Range("A1").Value = MODE_BDDS
MsgBox "Работа макроса закончена"
End Sub






















'==================================================================================================================================
'============================================================    Режим Ввод Данных из Выпуска Продукции   ========================================
'==================================================================================================================================
'Макрос ВКЛючения режима Режим Ввод Данных из Выпуска Продукции
' Форматирует таблицу для режима отображения Режим Ввод Данных из Выпуска Продукции
'in:
' текущий лист книги
'==================================================================================================================================

Sub MSK_ON_INPUT_DATA()

Dim password As String
Dim good_password As String

Dim MODE_MODE_INPUT_DATA As String
MODE_MODE_INPUT_DATA = "Режим Ввод Данных из Выпуска Продукции"

If Range("A1").Value = MODE_MODE_INPUT_DATA Then
    MsgBox "Режим Ввод Данных из Выпуска Продукции уже применен. Макрос не запущен."
    Exit Sub
End If

' введем защиту от случайного запуска макроса - нужно ввести пароль
good_password = "987"
password = InputBox("Введите пароль на запуск макроса (987):")
If Not IsNumeric(password) Then
    MsgBox "Пароль не верен - не число. Макрос не запущен."
    Exit Sub
End If
If password <> good_password Then
    MsgBox "Пароль не верен - не _верное_ число. Макрос не запущен."
    Exit Sub
End If

' Назначение столбцов
' TODO -  вывести эти диапазоны как константы и сделать функции скрытия столбцов, пердавая туда константы как архив
' и тогда не надо в каждой функции дублировать эти  SET-ы
' ' ---------------------- столбцы группы Спецификация --------------------------------------
' set Colunms_Group_specf                 = Range("C1:D1") 'столбцы группа Спецификации
' ' ---------------------- столбцы группы Производителя --------------------------------------
' set Colunms_Group_producer              = Range("F1:H1") 'столбцы группа Производитель

' ---------------------- все кварталы ---------------------------------------

set Columns_Group_2kvartal              = Range("AL1:AW1") 'столбцы группа Аванс/Ок.расчет/Примечание для 2 квартала 2025г
set Columns_Group_3kvartal              = Range("BF1:BQ1") 'столбцы группа Аванс/Ок.расчет/Примечание для 3 квартала 2025г
set Columns_Group_4kvartal              = Range("BZ1:CK1") 'столбцы группа Аванс/Ок.расчет/Примечание для 4 квартала 2025г


' ---------------------- 1 квартал ---------------------------------------
' set Column_need_plan_1kvartal           = Range("K1")  ' потребность-план 1 квартала 2025г
' set Column_initial_warehouse_balance1kv = Range("L1")  ' столбец начальный складской остаток на 1 квартал 2025г
' set Column_plan_1kvartal                = Range("M1")  ' столбец план реализации 1 квартала 2025г
' set Column_need_1kvartal                = Range("N1")  ' потребность 1 квартала 2025г
' set Column_buy_1kvartal                 = Range("O1")  ' в закупку 1 квартал 2025г
' set Column_outgo_1kvartal               = Range("P1")  ' расход 1 квартал 2025г
' set Column_final_warehouse_balance1kv   = Range("Q1")  ' конечный складской остаток 1 квартал 2025г
' set Columns_jan_1kvartal                = Range("R1:T1") 'группа столбцов январь
' set Columns_feb_1kvartal                = Range("V1:X1") 'группа столбцов февраль
' set Columns_march_1kvartal              = Range("Z1:AB1") 'группа столбцов март
' set Columns_Group_1kvartal              = Range("R1:AC1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г

' ---------------------- 2 квартал ---------------------------------------
set Column_need_plan_2kvartal           = Range("AE1")  ' потребность-план 2 квартала 2025г
set Column_initial_warehouse_balance2kv = Range("AF1")  ' столбец начальный складской остаток на 2 квартал 2025г
set Column_plan_2kvartal                = Range("AG1")  ' столбец план реализации 2 квартала 2025г
set Column_need_2kvartal                = Range("AH1")  ' потребность 2 квартала 2025г
set Column_buy_2kvartal                 = Range("AI1")  ' в закупку 2 квартал 2025г
set Column_outgo_2kvartal               = Range("AJ1")  ' расход 2 квартал 2025г
set Column_final_warehouse_balance2kv   = Range("AK1")  ' конечный складской остаток 2 квартал 2025г
set Columns_aprl_2kvartal               = Range("AL1:AN1") 'группа столбцов апрель
set Columns_may_2kvartal                = Range("AP1:AR1") 'группа столбцов май
set Columns_june_2kvartal               = Range("AT1:AV1") 'группа столбцов июнь
set Columns_Group_2kvartal              = Range("AL1:AW1") 'столбцы группа Аванс/Ок.расчет/Примечание для 2 квартала 2025г

' ---------------------- 3 квартал ---------------------------------------
set Column_need_plan_3kvartal           = Range("AY1")  ' потребность-план 3 квартала 2025г
set Column_initial_warehouse_balance3kv = Range("AZ1")  ' столбец начальный складской остаток на 3 квартал 2025г
set Column_plan_3kvartal                = Range("BA1")  ' столбец план реализации 3 квартала 2025г
set Column_need_3kvartal                = Range("BB1")  ' потребность 3 квартала 2025г
set Column_buy_3kvartal                 = Range("BC1")  ' в закупку 3 квартал 2025г
set Column_outgo_3kvartal               = Range("BD1")  ' расход 3 квартал 2025г
set Column_final_warehouse_balance3kv   = Range("BE1")  ' конечный складской остаток 3 квартал 2025г
set Columns_jule_3kvartal               = Range("BF1:BH1") 'группа столбцов июль
set Columns_augst_3kvartal              = Range("BJ1:BL1") 'группа столбцов август
set Columns_sept_3kvartal               = Range("BN1:BP1") 'группа столбцов сентябрь
set Columns_Group_3kvartal              = Range("BF1:BQ1") 'столбцы группа Аванс/Ок.расчет/Примечание для 3 квартала 2025г

' ---------------------- 4 квартал ---------------------------------------
set Column_need_plan_4kvartal           = Range("BS1")  ' потребность-план 4 квартала 2025г
set Column_initial_warehouse_balance4kv = Range("BT1")  ' столбец начальный складской остаток на 4 квартал 2025г
set Column_plan_4kvartal                = Range("BU1")  ' столбец план реализации 4 квартала 2025г
set Column_need_4kvartal                = Range("BV1")  ' потребность 4 квартала 2025г
set Column_buy_4kvartal                 = Range("BW1")  ' в закупку 4 квартал 2025г
set Column_outgo_4kvartal               = Range("BX1")  ' расход 4 квартал 2025г
set Column_final_warehouse_balance4kv   = Range("BY1")  ' конечный складской остаток 4 квартал 2025г
set Columns_oktr_4kvartal               = Range("BZ1:CB1") 'группа столбцов октябрь
set Columns_nov_4kvartal                = Range("CD1:CF1") 'группа столбцов ноябрь
set Columns_dec_4kvartal                = Range("CH1:CJ1") 'группа столбцов декабрь
set Columns_Group_4kvartal              = Range("BZ1:CK1") 'столбцы группа Аванс/Ок.расчет/Примечание для 4 квартала 2025г

' ---------------------- 1 квартал следующего года ---------------------------------------
set Column_need_plan_next_kvartal           = Range("CM1")  ' потребность-план 1 квартала 2026г
set Column_initial_warehouse_balance_next_kv = Range("CN1")  ' столбец начальный складской остаток на 1 квартал 2026г
set Column_plan_next_kvartal                = Range("CO1")  ' столбец план реализации 1 квартала 2026г
set Column_need_next_kvartal                = Range("CP1")  ' потребность 1 квартала 2026г
set Column_buy_next_kvartal                 = Range("CQ1")  ' в закупку 1 квартал 2026г
set Column_outgo_next_kvartal               = Range("CR1")  ' расход 1 квартал 2026г
set Column_final_warehouse_balance_next_kv   = Range("CS1")  ' конечный складской остаток 1 квартал 2026г
set Columns_jan_next_kvartal                = Range("CT1:CV1") 'группа столбцов январь
set Columns_feb_next_kvartal                = Range("CX1:CZ1") 'группа столбцов февраль
set Columns_march_next_kvartal                = Range("DB1:DD1") 'группа столбцов март
set Columns_Group_next_kvartal              = Range("CT1:DE1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г

' -----------------------------------
' -----------------------------------
' ' ----------------------------------- Назначение строк ---------------------------------------------------------------
' set Rows_GSHVA                              = Range("6:23") ' все строки от ГШВА
' set Rows_BP14_5                             = Range("25:37") ' все строки от БП14-5
' set Rows_TAIS                               = Range("39:46") ' все строки от ТАИС
' set Rows_ACMicro                            = Range("51:52") ' все строки от АСМикро
' set Rows_OMS2000                            = Range("54:54") ' все строки от OMS-2000
' set Rows_OMS2000M                           = Range("56:63") ' все строки от OMS-2000M
' set Rows_US6                                = Range("65:68") ' все строки от УС6


' подготовим таблицу для дальнейших модификаций - покажем все скрытые строки/столбцы, откроем и разгруппируем группы
InitTable

' ' разворачивание всех группированных столбцов
' ' ActiveSheet.Outline.ShowLevels ColumnLevels:=1
' ' Worksheets("Детализация").Outline.ShowLevels 6, 6
' ExpandAll

' Range("A1:DM1").EntireColumn.Hidden = False

' ' отменим группировку всех столбцов
' ' Range("A1:DM1").Ungroup
' ' Range("A1:DM1").Ungroup
' UngroupAll


' группируем группы Спецификации и Производитель
' Colunms_Group_specf.Group
' Colunms_Group_producer.Group

' Скроем "Всего Потребности в год"
Range("J1").EntireColumn.Hidden = True
' Скроем крайнюю правую таблицу "2026"
Range("DG1:DM1").EntireColumn.Hidden = True



' Columns_jan_1kvartal                = Range("R1:T1") 'группа столбцов январь
' Columns_feb_1kvartal                = Range("V1:X1") 'группа столбцов февраль
' Columns_march_1kvartal              = Range("Z1:AB1") 'группа столбцов март
' Columns_Group_1kvartal              = Range("R1:AC1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г





Dim rez As Boolean
set Base_first_cell                         = Range("A1")





' скроем группы Спецификации и Производитель
Base_first_cell.Select
rez = Entire_Column (Offset_Colunms_Group_specf, Size_Colunms_Group_specf)
Base_first_cell.Select
rez = Entire_Column (Offset_Colunms_Group_producer, Size_Colunms_Group_producer)
' скроем все, что не нужно в 1 квартале
Base_first_cell.Select
rez = Entire_Column (Offset_Column_need_plan_1kvartal, 5)
Base_first_cell.Select
rez = Entire_Column (Offset_Column_final_warehouse_balance1kv, 1)
Base_first_cell.Select
rez = Entire_Column (Offset_Columns_Group_1kvartal, Size_Columns_Group_1kvartal)

' Скрываем ненужные для Режим Ввод Данных из Выпуска Продукции столбцы
' Columns_Group_1kvartal.EntireColumn.Hidden = True   
'Column_need_plan_1kvartal.EntireColumn.Hidden = True          
' Column_initial_warehouse_balance1kv.EntireColumn.Hidden = True
' Column_plan_1kvartal.EntireColumn.Hidden = True               
' Column_need_1kvartal.EntireColumn.Hidden = True            
' Column_buy_1kvartal.EntireColumn.Hidden = True    
' Column_outgo_1kvartal.EntireColumn.Hidden = True              
' Column_final_warehouse_balance1kv.EntireColumn.Hidden = True  
' Columns_Group_1kvartal.EntireColumn.Hidden = True  
' Range("AD1").EntireColumn.Hidden = True  


' Скрываем ненужные для Режим Ввод Данных из Выпуска Продукции столбцы
Columns_Group_2kvartal.EntireColumn.Hidden = True      
Column_need_plan_2kvartal.EntireColumn.Hidden = True           
Column_initial_warehouse_balance2kv.EntireColumn.Hidden = True 
Column_plan_2kvartal.EntireColumn.Hidden = True                
Column_need_2kvartal.EntireColumn.Hidden = True
Column_buy_2kvartal.EntireColumn.Hidden = True                
' Column_outgo_2kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance2kv.EntireColumn.Hidden = True   
Columns_Group_2kvartal.EntireColumn.Hidden = True
Range("AX1").EntireColumn.Hidden = True 

' Скрываем ненужные для Режим Ввод Данных из Выпуска Продукции столбцы
Columns_Group_3kvartal.EntireColumn.Hidden = True  
Column_need_plan_3kvartal.EntireColumn.Hidden = True           
Column_initial_warehouse_balance3kv.EntireColumn.Hidden = True 
Column_plan_3kvartal.EntireColumn.Hidden = True                
Column_need_3kvartal.EntireColumn.Hidden = True
Column_buy_3kvartal.EntireColumn.Hidden = True                
' Column_outgo_3kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance3kv.EntireColumn.Hidden = True   
Columns_Group_3kvartal.EntireColumn.Hidden = True
Range("BR1").EntireColumn.Hidden = True 

' Скрываем ненужные для Режим Ввод Данных из Выпуска Продукции столбцы
Columns_Group_4kvartal.EntireColumn.Hidden = True 
Column_need_plan_4kvartal.EntireColumn.Hidden = True           
Column_initial_warehouse_balance4kv.EntireColumn.Hidden = True 
Column_plan_4kvartal.EntireColumn.Hidden = True                
Column_need_4kvartal.EntireColumn.Hidden = True
Column_buy_4kvartal.EntireColumn.Hidden = True                
' Column_outgo_4kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance4kv.EntireColumn.Hidden = True   
Columns_Group_4kvartal.EntireColumn.Hidden = True
Range("CL1").EntireColumn.Hidden = True 


' Скрываем ненужные для Режим Ввод Данных из Выпуска Продукции столбцы
Columns_Group_next_kvartal.EntireColumn.Hidden = True   
Column_need_plan_next_kvartal.EntireColumn.Hidden = True            
Column_initial_warehouse_balance_next_kv.EntireColumn.Hidden = True 
Column_plan_next_kvartal.EntireColumn.Hidden = True                 
Column_need_next_kvartal.EntireColumn.Hidden = True
Column_buy_next_kvartal.EntireColumn.Hidden = True                  
' Column_outgo_next_kvartal.EntireColumn.Hidden = True               
Column_final_warehouse_balance_next_kv.EntireColumn.Hidden = True   
Columns_Group_next_kvartal.EntireColumn.Hidden = True   
Range("DF1").EntireColumn.Hidden = True             


' ----------------------------------- Назначение строк ---------------------------------------------------------------
' set Rows_GSHVA                              = Range("6:23") ' все строки от ГШВА
set Rows_BP14_5                             = Range("25:37") ' все строки от БП14-5
set Rows_TAIS                               = Range("39:46") ' все строки от ТАИС
set Rows_ACMicro                            = Range("51:52") ' все строки от АСМикро
set Rows_OMS2000                            = Range("54:54") ' все строки от OMS-2000
set Rows_OMS2000M                           = Range("56:63") ' все строки от OMS-2000M
set Rows_US6                                = Range("65:68") ' все строки от УС6


' скроем строки группы ГШВА
Base_first_cell.Select
rez = Entire_Row(Offset_Rows_GSHVA_Summ, Size_Rows_GSHVA_Summ)

' Rows_GSHVA.EntireRow.Hidden = True  
Rows_BP14_5.EntireRow.Hidden = True   
Rows_TAIS.EntireRow.Hidden = True    
Rows_ACMicro.EntireRow.Hidden = True  
Rows_OMS2000.EntireRow.Hidden = True  
Rows_OMS2000M.EntireRow.Hidden = True  
Rows_US6.EntireRow.Hidden = True  






' закроем(свернем) все группы
Worksheets("Детализация").Outline.ShowLevels 1, 1
Range("A1").Value = MODE_INPUT_DATA
MsgBox "Работа макроса закончена"
End Sub

    












