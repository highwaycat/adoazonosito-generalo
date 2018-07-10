Attribute VB_Name = "AdoazonositoModul"
Public Function ADOAZONOSITO(birthDate As Date, counter As String)
    If Len(counter) > 3 Then
        ADOAZONOSITO = "HIBA! TÚL MAGAS COUNTER! (MAX 999)"
    ElseIf Len(counter) < 1 Then
        ADOAZONOSITO = "HIBA! TÚL ALACSONY COUNTER! (MIN 000)"
    End If
    
    Dim result As String
    
    Dim maganyszemelyJel As String
    maganszemelyJel = "8"
    
    Dim daysSinceReferenceDate As Long
    daysSinceReferenceDate = DateDiff("d", "01/01/1867", birthDate)
    
    result = maganszemelyJel + CStr(daysSinceReferenceDate) + CStr(counter)
    
    Dim checkSum As Integer
    checkSum = CALCULATE_CHECKSUM(result)
    
    If checkSum = 10 Then
        ADOAZONOSITO = "HIBA! NEM ADHATÓ KI EZEKKEL AZ ADATOKKAL ADÓAZONOSÍTÓ JEL!"
    Else
        result = result + CStr(checkSum)
        ADOAZONOSITO = result
    End If
End Function


Private Function CALCULATE_CHECKSUM(text As String)
    Dim i As Integer
    Dim sum As Integer
    For i = 1 To Len(text)
        sum = sum + ((Asc(Mid(text, i, 1)) - Asc("0")) * i)
    Next
    CALCULATE_CHECKSUM = sum Mod 11
End Function

Public Function ADOAZONOSITO_WITH_RANDOM_COUNTER(birthDate As Date)
    Dim randomCounter As Integer
    randomCounter = Int(900 * Rnd + 100)
    ADOAZONOSITO_WITH_RANDOM_COUNTER = ADOAZONOSITO(birthDate, CStr(randomCounter))
End Function
