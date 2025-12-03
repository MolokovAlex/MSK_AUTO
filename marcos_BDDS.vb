Sub ExpandAll()
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

Sub UngroupAll()
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
End Sub






















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

' разворачивание всех группированных столбцов
' Worksheets("Детализация").Outline.ShowLevels 6, 6
' ActiveSheet.Cells.ClearOutline
ExpandAll

Range("A1:DM1").EntireColumn.Hidden = False


' отменим группировку всех столбцов
' Range("A1:DM1").Ungroup
' Range("A1:DM1").Ungroup
UngroupAll

' ---------------------- столбцы группы Спецификация --------------------------------------
set Colunms_Group_specf                 = Range("C1:D1") 'столбцы группа Спецификации
' ---------------------- столбцы группы Производителя --------------------------------------
set Colunms_Group_producer              = Range("F1:H1") 'столбцы группа Производитель
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

' Назначение столбцов
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
set Column_final_warehouse_balance_next_kv   = Range("CS1")  ' конечный складской остаток 1 квартал 2026г
set Columns_jan_next_kvartal                = Range("CT1:CV1") 'группа столбцов январь
set Columns_feb_next_kvartal                = Range("CX1:CZ1") 'группа столбцов февраль
set Columns_march_next_kvartal                = Range("DB1:DD1") 'группа столбцов март
set Columns_Group_next_kvartal              = Range("CT1:DE1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г

' -----------------------------------


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


' Скрываем ненужные для Режим Ввод Данных из Выпуска Продукции столбцы
Columns_Group_1kvartal.EntireColumn.Hidden = True   
Column_need_plan_1kvartal.EntireColumn.Hidden = True          
Column_initial_warehouse_balance1kv.EntireColumn.Hidden = True
Column_plan_1kvartal.EntireColumn.Hidden = True               
Column_need_1kvartal.EntireColumn.Hidden = True            
Column_buy_1kvartal.EntireColumn.Hidden = True    
' Column_outgo_1kvartal.EntireColumn.Hidden = True              
Column_final_warehouse_balance1kv.EntireColumn.Hidden = True  
Columns_Group_1kvartal.EntireColumn.Hidden = True  
Range("AD1").EntireColumn.Hidden = True  


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



' закроем(свернем) все группы
Worksheets("Детализация").Outline.ShowLevels 1, 1
Range("A1").Value = MODE_INPUT_DATA
MsgBox "Работа макроса закончена"
End Sub





















'==================================================================================================================================
'============================================================    Режим План платежей       ========================================
'==================================================================================================================================
'Макрос ВКЛючения режима отображеия Плана Платежей
' Форматирует таблицу для режима отображения Плана платежей
'in:
' текущий лист книги
'==================================================================================================================================


Sub MSK_ON_PP()

Dim constPathFile As Variant
Dim pathSearchFile As Variant
Dim listHiddenColumn As Variant
Dim element As Variant
Dim pathArch As String

Dim MODE_PP As String
MODE_PP = "Режим ПланПлатежей"

If Range("A1").Value = MODE_PP Then
    MsgBox "Режим ПланПлатежей уже применен. Макрос не запущен."
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
' ---------------------- все кварталы - исходное состояние --------------------------------------
set Colunms_Group_producer              = Range("F1:G1") 'столбцы группа Производитель, Входимость в специф..
set Columns_Group_1kvartal              = Range("P1:X1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г
set Columns_Group_2kvartal              = Range("AF1:AN1") 'столбцы группа Аванс/Ок.расчет/Примечание для 2 квартала 2025г
set Columns_Group_3kvartal              = Range("AV1:BD1") 'столбцы группа Аванс/Ок.расчет/Примечание для 3 квартала 2025г
set Columns_Group_4kvartal              = Range("BL1:BT1") 'столбцы группа Аванс/Ок.расчет/Примечание для 4 квартала 2025г
' ---------------------- все кварталы - состояние послед преобразования --------------------------------------
set Colunms_Group_producer_after        = Range("C1:G1") 'столбцы группа Производитель, Входимость в специф..

' ---------------------- 1 квартал ---------------------------------------
set Column_need_plan_1kvartal           = Range("I1")  ' потребность-план 1 квартала 2025г
set Column_initial_warehouse_balance1kv = Range("J1")  ' столбец начальный складской остаток на 1 квартал 2025г
set Column_plan_1kvartal                = Range("K1")  ' столбец план реализации 1 квартала 2025г
set Column_need_1kvartal                = Range("L1")  ' потребность 1 квартала 2025г
set Columns_01_1kvartal                 = Range("I1:L1")

set Column_buy_1kvartal                 = Range("M1")  ' в закупку 1 квартал 2025г
set Column_outgo_1kvartal               = Range("N1")  ' расход 1 квартал 2025г
set Column_final_warehouse_balance1kv   = Range("O1")  ' конечный складской остаток 1 квартал 2025г
set Columns_02_1kvartal                 = Range("N1:O1")
' ---------------------- 2 квартал ---------------------------------------
set Column_need_plan_2kvartal           = Range("Y1")  ' потребность-план 2 квартала 2025г
set Column_initial_warehouse_balance2kv = Range("Z1")  ' столбец начальный складской остаток на 2 квартал 2025г
set Column_plan_2kvartal                = Range("AA1")  ' столбец план реализации 2 квартала 2025г
set Column_need_2kvartal                = Range("AB1")  ' потребность 2 квартала 2025г
set Columns_01_2kvartal                 = Range("Y1:AB1")

set Column_buy_2kvartal                 = Range("AC1")  ' в закупку 2 квартал 2025г
set Column_outgo_2kvartal               = Range("AD1")  ' расход 2 квартал 2025г
set Column_final_warehouse_balance2kv   = Range("AE1")  ' конечный складской остаток 2 квартал 2025г
set Columns_02_2kvartal                 = Range("AD1:AE1")
' ---------------------- 3 квартал ---------------------------------------
set Column_need_plan_3kvartal           = Range("AO1")  ' потребность-план 3 квартала 2025г
set Column_initial_warehouse_balance3kv = Range("AP1")  ' столбец начальный складской остаток на 3 квартал 2025г
set Column_plan_3kvartal                = Range("AQ1")  ' столбец план реализации 3 квартала 2025г
set Column_need_3kvartal                = Range("AR1")  ' потребность 3 квартала 2025г
set Columns_01_3kvartal                 = Range("AO1:AR1")

set Column_buy_3kvartal                 = Range("AS1")  ' в закупку 3 квартал 2025г
set Column_outgo_3kvartal               = Range("AT1")  ' расход 3 квартал 2025г
set Column_final_warehouse_balance3kv   = Range("AU1")  ' конечный складской остаток 3 квартал 2025г
set Columns_02_3kvartal                 = Range("AT1:AU1")
' ---------------------- 4 квартал ---------------------------------------
set Column_need_plan_4kvartal           = Range("BE1")  ' потребность-план 4 квартала 2025г
set Column_initial_warehouse_balance4kv = Range("BF1")  ' столбец начальный складской остаток на 4 квартал 2025г
set Column_plan_4kvartal                = Range("BG1")  ' столбец план реализации 4 квартала 2025г
set Column_need_4kvartal                = Range("BH1")  ' потребность 4 квартала 2025г
set Columns_01_4kvartal                 = Range("BE1:BH1")

set Column_buy_4kvartal                 = Range("BI1")  ' в закупку 4 квартал 2025г
set Column_outgo_4kvartal               = Range("BJ1")  ' расход 4 квартал 2025г
set Column_final_warehouse_balance4kv   = Range("BK1")  ' конечный складской остаток 4 квартал 2025г
set Columns_02_4kvartal                 = Range("BJ1:BK1")
' -----------------------------------

' constPathFile = Array( _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Оборудование\Выпуск ПО\Заводские номера ПО.mdb", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Учет сертифицированной продукции\Оборудование\БазаШ5Л\БазаШ5Л.accdb", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Учет сертифицированной продукции\Оборудование\БазаШ3\БазаШ3.accdb", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Оборудование\Выпуск продукции\Журнал учета выпуска продукции\Журнал учёта выпуска продукции.xlsx", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Учет сертифицированной продукции\Оборудование\Сертификация_ФСТЭК_25.01.31.xls", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Учет сертифицированной продукции\Оборудование\Система Шорох-3_ФСБ_25.04.17.xls" _
' )

' pathSearchFileXLSX = Array( _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\Возможность выпуска Шорох-3\", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\Возможность выпуска Шорох-5Л\", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\Возможность выпуска Телефон-Н2\", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\Выпуск продукции\", _
' "S:\НТЦ\Заявки\Производство\2025\", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\План платежей\2025\", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\Возможность выпуска Старкад\", _
' "S:\НТЦ\ПРО\ВЫПУСК ПРОДУКЦИИ\Планирование\Возможность выпуска Систем\", _
' "P:\ДО_ПрО–ДКР_ЛабСП\График проведения СП\" _
' )

' listHiddenColumn = Array("I1:L1", _
' "N1:O1", _
' "Z1:AB1" _
' )


'Range("I1:L1").EntireColumn.Hidden = False
'Range("N1:O1").EntireColumn.Hidden = Fasle
'Range("Y1:AB1").EntireColumn.Hidden = False

' For Each element In listHiddenColumn
'     Range(element).EntireColumn.Hidden = False
'     'pathArch = "M:\ArchFolderMSKM\"
' Next
Colunms_Group_producer_after.Ungroup

' Показываем все скрытые до этого столбцы
Columns_01_1kvartal.EntireColumn.Hidden = False
Columns_02_1kvartal.EntireColumn.Hidden = False
Columns_01_2kvartal.EntireColumn.Hidden = False
Columns_02_2kvartal.EntireColumn.Hidden = False
Columns_01_3kvartal.EntireColumn.Hidden = False
Columns_02_3kvartal.EntireColumn.Hidden = False
Columns_01_4kvartal.EntireColumn.Hidden = False
Columns_02_4kvartal.EntireColumn.Hidden = False

' возвращаем группировку на столбцы Аванс/Ок.расчет/Примечание для 1,2,3,4 кварталов
' Range("P1", "X1").Group
Columns_Group_1kvartal.Group
Columns_Group_2kvartal.Group
Columns_Group_3kvartal.Group
Columns_Group_4kvartal.Group
Colunms_Group_producer.Group



' закроем(свернем) все группы
Worksheets("Детализация").Outline.ShowLevels 1, 1




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
Range("A1").Value = MODE_PP
MsgBox "Работа макроса закончена"
End Sub



