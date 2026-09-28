Option Explicit

' =========================================================================
' Расчет ипотеки с растущими ставками для месяцев 1-12 (субсидированный период)
' Логика:
'  1) Ищем остаток долга после 12-го платежа (Bal12) методом деления отрезка
'     пополам, так, чтобы прямой прогон месяцев 1..12 (с найденными по каждому
'     месяцу ставками) давал именно этот остаток (условие самосогласованности).
'  2) Для каждого кандидата Bal12: считаем аннуитет 13-го месяца по известной
'     ставке B4 на срок B5 от Bal12 (обычный PMT). Далее разворачиваем цепочку:
'     annuity(13) = AVERAGE(annuity(1..12)) * (1+B7)
'     annuity(k)  = AVERAGE(annuity(1..k-1)) * (1+B7),  k = 2..12
'     Так как соотношения линейны по annuity(1), можно выразить все 12
'     целевых аннуитетов как annuity(1)*coef(k), где coef(k) считается из
'     базовой (нормированной) рекурсии, а затем масштабируется под нужное
'     среднее annuity(1..12).
'  3) Для каждого месяца m=1..12 ищем ставку r_m (в годовых, дробное число),
'     такую что PMT(r_m/12, B3-(m-1), остаток_до_месяца_m) = целевой annuity(m).
'     Метод - деление отрезка пополам по ставке в диапазоне [1E-9; 10].
'  4) Проценты = остаток_до_месяца * r_m/12; Основной долг = annuity - проценты;
'     новый остаток = старый остаток - основной долг.
'  5) Месяцы 13..B3: ставка = B4, аннуитет постоянный = annuity(13),
'     обычная амортизация. Месяцы после B3 (до конца таблицы) - платежи 0,
'     остаток 0 (кредит выплачен).
' =========================================================================

Function PmtAnnual(rateAnnual As Double, nper As Double, pv As Double) As Double
    Dim r As Double
    r = rateAnnual / 12
    If Abs(r) < 0.0000000001 Then
        PmtAnnual = pv / nper
    Else
        PmtAnnual = pv * r / (1 - (1 + r) ^ (-nper))
    End If
End Function

' Ищет годовую ставку r в [loR;hiR], при которой PmtAnnual(r,nper,pv)=targetPmt.
' Возвращает True, если решение найдено (функция монотонно растет по r,
' поэтому бисекция применима), False - если целевой платеж вне достижимого
' диапазона (targetPmt меньше платежа при r->0 или больше платежа при r=hiR).
Function SolveRateBisect(targetPmt As Double, nper As Double, pv As Double, _
                          ByRef rOut As Double) As Boolean
    Dim loR As Double, hiR As Double
    Dim fLo As Double, fHi As Double, fMid As Double, midR As Double
    Dim i As Long
    loR = 0.000000001
    hiR = 10
    fLo = PmtAnnual(loR, nper, pv) - targetPmt
    fHi = PmtAnnual(hiR, nper, pv) - targetPmt
    If fLo * fHi > 0 Then
        SolveRateBisect = False
        Exit Function
    End If
    For i = 1 To 200
        midR = (loR + hiR) / 2
        fMid = PmtAnnual(midR, nper, pv) - targetPmt
        If fMid = 0 Then Exit For
        If fLo * fMid < 0 Then
            hiR = midR
            fHi = fMid
        Else
            loR = midR
            fLo = fMid
        End If
    Next i
    rOut = (loR + hiR) / 2
    SolveRateBisect = True
End Function

' Прогоняет месяцы 1..12 при заданном остатке долга после 12 месяцев (Bal12Guess).
' Возвращает True при успехе (все ставки найдены) и заполняет:
'   rates(1..12)      - годовые ставки по месяцам
'   annuities(1..12)  - аннуитетные платежи по месяцам
'   principals(1..12) - платежи по основному долгу
'   interests(1..12)  - платежи по процентам
'   annuity13         - аннуитет с 13-го месяца (постоянный далее)
'   balanceAfter12    - фактический остаток после прогона (сравнивается с Bal12Guess)
Function SimulateForward(Bal12Guess As Double, B2 As Double, B3 As Double, _
                          B4 As Double, B5 As Double, B7 As Double, _
                          ByRef rates() As Double, ByRef annuities() As Double, _
                          ByRef principals() As Double, ByRef interests() As Double, _
                          ByRef annuity13 As Double, ByRef balanceAfter12 As Double) As Boolean

    Dim coef(1 To 12) As Double
    Dim k As Long, j As Long
    Dim avgPrev As Double, sumC As Double, avg12C As Double
    Dim avg12Target As Double, a1 As Double
    Dim balance As Double, targetPmt As Double, nper As Double
    Dim rOut As Double, interest As Double, principal As Double
    Dim ok As Boolean

    annuity13 = PmtAnnual(B4, B5, Bal12Guess)
    avg12Target = annuity13 / (1 + B7)

    coef(1) = 1
    For k = 2 To 12
        sumC = 0
        For j = 1 To k - 1
            sumC = sumC + coef(j)
        Next j
        avgPrev = sumC / (k - 1)
        coef(k) = avgPrev * (1 + B7)
    Next k

    sumC = 0
    For j = 1 To 12
        sumC = sumC + coef(j)
    Next j
    avg12C = sumC / 12
    a1 = avg12Target / avg12C

    balance = B2
    For k = 1 To 12
        targetPmt = a1 * coef(k)
        nper = B3 - (k - 1)
        ok = SolveRateBisect(targetPmt, nper, balance, rOut)
        If Not ok Then
            SimulateForward = False
            Exit Function
        End If
        rates(k) = rOut
        annuities(k) = targetPmt
        interest = balance * rOut / 12
        principal = targetPmt - interest
        interests(k) = interest
        principals(k) = principal
        balance = balance - principal
    Next k

    balanceAfter12 = balance
    SimulateForward = True
End Function

Sub РасчетИпотеки()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim B2 As Double, B3 As Double, B4 As Double, B5 As Double, B6 As Double, B7 As Double
    B2 = ws.Range("B2").Value      ' Сумма кредита
    B3 = ws.Range("B3").Value      ' Срок кредита, мес
    B4 = ws.Range("B4").Value      ' Рынок, ставка после субсидии
    B6 = ws.Range("B6").Value      ' Субсидия, срок, мес
    B5 = B3 - B6                   ' Рынок, срок после субсидии
    ws.Range("B5").Value = B5
    B7 = ws.Range("B7").Value      ' Допуск по шагу аннуитета

    If B6 <> 12 Then
        MsgBox "Внимание: макрос рассчитан на субсидию длительностью 12 месяцев (B6=12)." & _
               " Текущее значение B6=" & B6 & ". Логика поиска ставок 1-12 месяца требует правки.", vbExclamation
        Exit Sub
    End If

    ' --- Поиск остатка долга после 12-го месяца (Bal12) методом деления отрезка ---
    Dim rates(1 To 12) As Double
    Dim annuities(1 To 12) As Double
    Dim principals(1 To 12) As Double
    Dim interests(1 To 12) As Double
    Dim annuity13 As Double, balanceAfter12 As Double

    Dim loB As Double, hiB As Double, midB As Double
    Dim fLo As Double, fHi As Double, fMid As Double
    Dim ok As Boolean, foundBracket As Boolean
    Dim i As Long, steps As Long

    ' Грубый скан для поиска интервала со сменой знака функции f(Bal12) = balanceAfter12(Bal12) - Bal12
    steps = 400
    foundBracket = False
    Dim xPrev As Double, fPrev As Double, hasPrev As Boolean
    Dim xCur As Double, fCur As Double
    hasPrev = False

    For i = 0 To steps
        xCur = B2 * 0.1 + (B2 * 2.9) * i / steps   ' диапазон от 0.1*B2 до 3*B2
        ok = SimulateForward(xCur, B2, B3, B4, B5, B7, rates, annuities, principals, interests, annuity13, balanceAfter12)
        If ok Then
            fCur = balanceAfter12 - xCur
            If hasPrev Then
                If fPrev * fCur <= 0 Then
                    loB = xPrev
                    hiB = xCur
                    fLo = fPrev
                    fHi = fCur
                    foundBracket = True
                    Exit For
                End If
            End If
            xPrev = xCur
            fPrev = fCur
            hasPrev = True
        Else
            hasPrev = False
        End If
    Next i

    If Not foundBracket Then
        MsgBox "Не удалось найти решение (не обнаружен интервал со сменой знака)." & _
               " Проверьте входные параметры B2:B7.", vbCritical
        Exit Sub
    End If

    ' Уточнение корня Bal12 методом деления отрезка пополам
    Dim bestBal12 As Double
    For i = 1 To 200
        midB = (loB + hiB) / 2
        ok = SimulateForward(midB, B2, B3, B4, B5, B7, rates, annuities, principals, interests, annuity13, balanceAfter12)
        If Not ok Then
            ' если середина не решается - сдвигаем к границе, которая решается
            loB = midB
            GoTo ContinueLoop
        End If
        fMid = balanceAfter12 - midB
        If fMid = 0 Then
            Exit For
        End If
        If fLo * fMid < 0 Then
            hiB = midB
            fHi = fMid
        Else
            loB = midB
            fLo = fMid
        End If
ContinueLoop:
    Next i

    bestBal12 = (loB + hiB) / 2
    ok = SimulateForward(bestBal12, B2, B3, B4, B5, B7, rates, annuities, principals, interests, annuity13, balanceAfter12)
    If Not ok Then
        MsgBox "Ошибка на финальном шаге расчета.", vbCritical
        Exit Sub
    End If

    ' --- Запись результатов в таблицу ---
    Dim firstRow As Long
    firstRow = 22   ' строка для месяца 1

    Dim lastMonthRow As Long
    lastMonthRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row  ' последняя строка с номером месяца

    Dim balance As Double
    balance = B2

    Dim m As Long, r As Long
    For m = 1 To (lastMonthRow - firstRow + 1)
        r = firstRow + m - 1

        If m <= 12 Then
            ws.Cells(r, "A").Value = rates(m)                 ' ставка на месяц (только A22:A33)
            ws.Cells(r, "C").Value = principals(m)             ' основной долг
            ws.Cells(r, "D").Value = interests(m)               ' проценты
            ws.Cells(r, "E").Value = annuities(m)               ' аннуитет
            balance = balance - principals(m)
            ws.Cells(r, "F").Value = -balance                  ' остаток со знаком минус

        ElseIf m <= B3 Then
            ' A34:A(21+B3) не трогаем - там уже стоит формула =$B$4 (по условию)
            Dim interestM As Double, principalM As Double
            interestM = balance * B4 / 12
            principalM = annuity13 - interestM
            If m = B3 Then
                ' последний месяц - гасим остаток полностью (страхуемся от накопленной погрешности округления)
                principalM = balance
            End If
            ws.Cells(r, "C").Value = principalM
            ws.Cells(r, "D").Value = interestM
            ws.Cells(r, "E").Value = annuity13
            balance = balance - principalM
            ws.Cells(r, "F").Value = -balance

        Else
            ' после окончания срока кредита - платежей нет
            ws.Cells(r, "C").Value = 0
            ws.Cells(r, "D").Value = 0
            ws.Cells(r, "E").Value = 0
            ws.Cells(r, "F").Value = 0
        End If
    Next m

    ' Числовые форматы (проценты для ставок, денежный для сумм)
    ws.Range(ws.Cells(firstRow, "A"), ws.Cells(firstRow + 11, "A")).NumberFormat = "0.0000%"
    ws.Range(ws.Cells(firstRow, "C"), ws.Cells(lastMonthRow, "F")).NumberFormat = "#,##0.00"

    MsgBox "Расчет завершен." & vbCrLf & _
           "Остаток долга после 12 мес.: " & Format(bestBal12, "#,##0.00") & vbCrLf & _
           "Аннуитет с 13-го месяца: " & Format(annuity13, "#,##0.00") & vbCrLf & _
           "Итоговый остаток после " & B3 & " мес.: " & Format(balance, "#,##0.00"), vbInformation

End Sub
