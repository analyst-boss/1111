Option Explicit

' =========================================================================
' Raschet ipoteki s rastuschimi stavkami dlya mesyacev 1-12 (subsidirovanny period)
' Logika:
'  1) Ischem ostatok dolga posle 12-go platezha (Bal12) metodom deleniya otrezka
'     popolam, tak, chtoby pryamoy progon mesyacev 1..12 (s naydennymi po kazhdomu
'     mesyacu stavkami) daval imenno etot ostatok (uslovie samosoglasovannosti).
'  2) Dlya kazhdogo kandidata Bal12: schitaem annuitet 13-go mesyaca po izvestnoy
'     stavke B4 na srok B5 ot Bal12 (obychny PMT). Dalee razvorachivaem cepochku:
'     annuity(13) = AVERAGE(annuity(1..12)) * (1+B7)
'     annuity(k)  = AVERAGE(annuity(1..k-1)) * (1+B7),  k = 2..12
'     Tak kak sootnosheniya lineyny po annuity(1), mozhno vyrazit vse 12
'     celevyh annuitetov kak annuity(1)*coef(k), gde coef(k) schitaetsya iz
'     bazovoy (normirovannoy) rekursii, a potom masshtabiruetsya pod nuzhnoe
'     srednee annuity(1..12).
'  3) Dlya kazhdogo mesyaca m=1..12 ischem stavku r_m (v godovyh, drobnoe chislo),
'     takuyu chto PMT(r_m/12, B3-(m-1), ostatok_do_mesyaca_m) = celevoy annuity(m).
'     Metod - delenie otrezka popolam po stavke v diapazone [1E-9; 10].
'  4) Procenty = ostatok_do_mesyaca * r_m/12; Osnovnoy dolg = annuity - procenty;
'     novy ostatok = stary ostatok - osnovnoy dolg.
'  5) Mesyacy 13..B3: stavka = B4, annuitet postoyanny = annuity(13),
'     obychnaya amortizaciya. Mesyacy posle B3 (do konca tablicy) - platezhi 0,
'     ostatok 0 (kredit vyplachen).
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

' Ischet godovuyu stavku r v [loR;hiR], pri kotoroy PmtAnnual(r,nper,pv)=targetPmt.
' Vozvraschaet True, esli reshenie naydeno (funkciya monotonno rastet po r,
' poetomu bisekciya primenima), False - esli celevoy platezh vne dostizhimogo
' diapazona (targetPmt menshe platezha pri r->0 ili bolshe platezha pri r=hiR).
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

' Progonyaet mesyacy 1..12 pri zadannom ostatke dolga posle 12 mesyacev (Bal12Guess).
' Vozvraschaet True pri uspehe (vse stavki naydeny) i zapolnyaet:
'   rates(1..12)      - godovye stavki po mesyacam
'   annuities(1..12)  - annuitetnye platezhi po mesyacam
'   principals(1..12) - platezhi po osnovnomu dolgu
'   interests(1..12)  - platezhi po procentam
'   annuity13         - annuitet s 13-go mesyaca (postoyanny dalee)
'   balanceAfter12    - fakticheskiy ostatok posle progona (sravnivaetsya s Bal12Guess)
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

Sub RaschetIpoteki()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim B2 As Double, B3 As Double, B4 As Double, B5 As Double, B6 As Double, B7 As Double
    B2 = ws.Range("B2").Value      ' Summa kredita
    B3 = ws.Range("B3").Value      ' Srok kredita, mes
    B4 = ws.Range("B4").Value      ' Rynok, stavka posle subsidii
    B6 = ws.Range("B6").Value      ' Subsidiya, srok, mes
    B5 = B3 - B6                   ' Rynok, srok posle subsidii
    ws.Range("B5").Value = B5
    B7 = ws.Range("B7").Value      ' Dopusk po shagu annuiteta

    If B6 <> 12 Then
        MsgBox "Vnimanie: makros rasschitan na subsidiyu dlitelnostyu 12 mesyacev (B6=12)." & _
               " Tekuschee znachenie B6=" & B6 & ". Logika poiska stavok 1-12 mesyaca trebuet pravki.", vbExclamation
        Exit Sub
    End If

    ' --- Poisk ostatka dolga posle 12-go mesyaca (Bal12) metodom deleniya otrezka ---
    Dim rates(1 To 12) As Double
    Dim annuities(1 To 12) As Double
    Dim principals(1 To 12) As Double
    Dim interests(1 To 12) As Double
    Dim annuity13 As Double, balanceAfter12 As Double

    Dim loB As Double, hiB As Double, midB As Double
    Dim fLo As Double, fHi As Double, fMid As Double
    Dim ok As Boolean, foundBracket As Boolean
    Dim i As Long, steps As Long

    ' Gruby skan dlya poiska intervala so smenoy znaka funkcii f(Bal12) = balanceAfter12(Bal12) - Bal12
    steps = 400
    foundBracket = False
    Dim xPrev As Double, fPrev As Double, hasPrev As Boolean
    Dim xCur As Double, fCur As Double
    hasPrev = False

    For i = 0 To steps
        xCur = B2 * 0.1 + (B2 * 2.9) * i / steps   ' diapazon ot 0.1*B2 do 3*B2
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
        MsgBox "Ne udalos nayti reshenie (ne obnaruzhen interval so smenoy znaka)." & _
               " Proverte vhodnye parametry B2:B7.", vbCritical
        Exit Sub
    End If

    ' Utochnenie kornya Bal12 metodom deleniya otrezka popolam
    Dim bestBal12 As Double
    For i = 1 To 200
        midB = (loB + hiB) / 2
        ok = SimulateForward(midB, B2, B3, B4, B5, B7, rates, annuities, principals, interests, annuity13, balanceAfter12)
        If Not ok Then
            ' esli seredina ne reshaetsya - sdvigaem k granice, kotoraya reshaetsya
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
        MsgBox "Oshibka na finalnom shage rascheta.", vbCritical
        Exit Sub
    End If

    ' --- Zapis rezultatov v tablicu ---
    Dim firstRow As Long
    firstRow = 22   ' stroka dlya mesyaca 1

    Dim lastMonthRow As Long
    lastMonthRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row  ' poslednyaya stroka s nomerom mesyaca

    Dim balance As Double
    balance = B2

    Dim m As Long, r As Long
    For m = 1 To (lastMonthRow - firstRow + 1)
        r = firstRow + m - 1

        If m <= 12 Then
            ws.Cells(r, "A").Value = rates(m)                 ' stavka na mesyac (tolko A22:A33)
            ws.Cells(r, "C").Value = principals(m)             ' osnovnoy dolg
            ws.Cells(r, "D").Value = interests(m)               ' procenty
            ws.Cells(r, "E").Value = annuities(m)               ' annuitet
            balance = balance - principals(m)
            ws.Cells(r, "F").Value = -balance                  ' ostatok so znakom minus

        ElseIf m <= B3 Then
            ' A34:A(21+B3) ne trogaem - tam uzhe stoit formula =$B$4 (po usloviyu)
            Dim interestM As Double, principalM As Double
            interestM = balance * B4 / 12
            principalM = annuity13 - interestM
            If m = B3 Then
                ' posledniy mesyac - gasim ostatok polnostyu (strahuemsya ot nakoplennoy pogreshnosti okrugleniya)
                principalM = balance
            End If
            ws.Cells(r, "C").Value = principalM
            ws.Cells(r, "D").Value = interestM
            ws.Cells(r, "E").Value = annuity13
            balance = balance - principalM
            ws.Cells(r, "F").Value = -balance

        Else
            ' posle okonchaniya sroka kredita - platezhey net
            ws.Cells(r, "C").Value = 0
            ws.Cells(r, "D").Value = 0
            ws.Cells(r, "E").Value = 0
            ws.Cells(r, "F").Value = 0
        End If
    Next m

    ' Chislovye formaty (procenty dlya stavok, denezhny dlya summ)
    ws.Range(ws.Cells(firstRow, "A"), ws.Cells(firstRow + 11, "A")).NumberFormat = "0.0000%"
    ws.Range(ws.Cells(firstRow, "C"), ws.Cells(lastMonthRow, "F")).NumberFormat = "#,##0.00"

    MsgBox "Raschet zavershen." & vbCrLf & _
           "Ostatok dolga posle 12 mes.: " & Format(bestBal12, "#,##0.00") & vbCrLf & _
           "Annuitet s 13-go mesyaca: " & Format(annuity13, "#,##0.00") & vbCrLf & _
           "Itogovy ostatok posle " & B3 & " mes.: " & Format(balance, "#,##0.00"), vbInformation

End Sub
