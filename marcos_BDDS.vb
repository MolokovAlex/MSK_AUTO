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

' Назначение столбцов
' ---------------------- столбцы группы Спецификация --------------------------------------
set Colunms_Group_specf                 = Range("C1:D1") 'столбцы группа Спецификации
' ---------------------- столбцы группы Производителя --------------------------------------
set Colunms_Group_producer              = Range("F1:H1") 'столбцы группа Производитель

If Range("A1").Value = MODE_BIG_TABLE Then
    MsgBox "Режим Большой таблицы уже применен. Макрос не запущен."
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

' разворачивание всех группированных столбцов
ActiveSheet.Outline.ShowLevels ColumnLevels:=8

' отменим группировку всех столбцов
Range("A1:DM1").Ungroup
Range("A1:DM1").Ungroup

' группируем группы Спецификации и Производитель
Colunms_Group_specf.Group
Colunms_Group_producer.Group

' закроем(свернем) все группы
Worksheets("Детализация").Outline.ShowLevels 1, 1

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

Dim constPathFile As Variant
Dim pathSearchFile As Variant
Dim listHiddenColumn As Variant
Dim element As Variant
Dim pathArch As String

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
' ---------------------- столбцы группы Спецификация --------------------------------------
set Colunms_Group_specf                 = Range("C1:D1") 'столбцы группа Спецификации
' ---------------------- столбцы группы Производителя --------------------------------------
set Colunms_Group_producer              = Range("F1:H1") 'столбцы группа Производитель

' ---------------------- все кварталы ---------------------------------------
set Columns_Group_1kvartal              = Range("R1:AC1") 'столбцы группа Аванс/Ок.расчет/Примечание для 1 квартала 2025г
set Columns_Group_2kvartal              = Range("AL1:AW1") 'столбцы группа Аванс/Ок.расчет/Примечание для 2 квартала 2025г
set Columns_Group_3kvartal              = Range("BF1:BQ1") 'столбцы группа Аванс/Ок.расчет/Примечание для 3 квартала 2025г
set Columns_Group_4kvartal              = Range("BZ1:CK1") 'столбцы группа Аванс/Ок.расчет/Примечание для 4 квартала 2025г


' ---------------------- 1 квартал ---------------------------------------
set Column_need_plan_1kvartal           = Range("K1")  ' потребность-план 1 квартала 2025г
set Column_initial_warehouse_balance1kv = Range("L1")  ' столбец начальный складской остаток на 1 квартал 2025г
set Column_plan_1kvartal                = Range("M1")  ' столбец план реализации 1 квартала 2025г
set Column_need_1kvartal                = Range("N1")  ' потребность 1 квартала 2025г
set Columns_jan_1kvartal                = Range("R1:T1") 'группа столбцов январь
set Columns_feb_1kvartal                = Range("V1:X1") 'группа столбцов февраль
set Columns_march_1kvartal              = Range("Z1:AB1") 'группа столбцов март

set Column_buy_1kvartal                 = Range("O1")  ' в закупку 1 квартал 2025г
set Column_outgo_1kvartal               = Range("P1")  ' расход 1 квартал 2025г
set Column_final_warehouse_balance1kv   = Range("Q1")  ' конечный складской остаток 1 квартал 2025г

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
'     Range(element).EntireColumn.Hidden = True
'     'pathArch = "M:\ArchFolderMSKM\"
' Next



' разворачивание всех группированных столбцов
ActiveSheet.Outline.ShowLevels ColumnLevels:=8

' отменим группировку всех столбцов
Range("A1:DM1").Ungroup


' группируем группы Спецификации и Производитель
Colunms_Group_specf.Group
Colunms_Group_producer.Group

' сгруппируем группы месяцев 1 квартала
Columns_jan_1kvartal.Group
Columns_feb_1kvartal.Group
Columns_march_1kvartal.Group

' сгруппируем группы 1 квартала
Columns_Group_1kvartal.Group

' закроем(свернем) все группы
Worksheets("Детализация").Outline.ShowLevels 1, 1










' снимаем группировку со столбцов Аванс/Ок.расчет/Примечание для 1,2,3,4 кварталов
' Columns_Group_1kvartal.Ungroup
' Columns_Group_2kvartal.Ungroup
' Columns_Group_3kvartal.Ungroup
' Columns_Group_4kvartal.Ungroup
' Скрываем ненужные для анализа БДДС столбцы
' Columns_01_1kvartal.EntireColumn.Hidden = True
' Columns_02_1kvartal.EntireColumn.Hidden = True
' Columns_01_2kvartal.EntireColumn.Hidden = True
' Columns_02_2kvartal.EntireColumn.Hidden = True
' Columns_01_3kvartal.EntireColumn.Hidden = True
' Columns_02_3kvartal.EntireColumn.Hidden = True
' Columns_01_4kvartal.EntireColumn.Hidden = True
' Columns_02_4kvartal.EntireColumn.Hidden = True

' закроем(свернем) все группы
' Worksheets("Детализация").Outline.ShowLevels 1, 1

' перегруппируем столбцы по вхождению в спецификации, поставщика и т.п.
' Colunms_Group_producer_after.Group

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



