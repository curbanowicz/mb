Option Explicit


'=========================================================
' FUNKCJA: WYCIĄGANIE NUMERU WNIOSKU Z TEMATU
'=========================================================
Function WyciagnijNumerWniosku(ByVal Temat As String) As String

    Dim Regex As Object
    Dim Matches As Object
    
    Set Regex = CreateObject("VBScript.RegExp")
    
    With Regex
        .Global = False
        .IgnoreCase = True
        
        '3-6 liter + 8-10 cyfr
        .Pattern = "(^|[^A-Za-z])([A-Za-z]{3,6}[0-9]{8,10})(?![0-9])"
    End With
    
    Set Matches = Regex.Execute(Temat)
    
    If Matches.Count > 0 Then
        WyciagnijNumerWniosku = Matches(0).SubMatches(1)
    Else
        WyciagnijNumerWniosku = ""
    End If

End Function


'=========================================================
' UZUPEŁNIANIE NUMERÓW WNIOSKÓW W ISTNIEJĄCYCH DANYCH
'=========================================================
Sub UzupelnijNumeryWnioskow()

    Dim Ws As Worksheet
    Dim OstatniWiersz As Long
    Dim i As Long
    
    Dim Temat As String
    Dim Numer As String
    
    Dim LicznikUzupelnionych As Long
    Dim LicznikBrakNumeru As Long
    
    Dim StaryScreenUpdating As Boolean
    Dim StaryCalculation As XlCalculation
    Dim StaryEnableEvents As Boolean
    
    On Error GoTo Blad
    
    Set Ws = ThisWorkbook.Worksheets("Outlook")
    
    StaryScreenUpdating = Application.ScreenUpdating
    StaryCalculation = Application.Calculation
    StaryEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Ostatni wiersz według kolumny D - Temat
    OstatniWiersz = Ws.Cells(Ws.Rows.Count, "D").End(xlUp).Row
    
    If OstatniWiersz < 2 Then
        MsgBox "Brak danych do uzupełnienia.", vbInformation
        GoTo Koniec
    End If
    
    For i = 2 To OstatniWiersz
        
        'Uzupełniamy tylko puste numery
        If Trim(CStr(Ws.Cells(i, "A").Value)) = "" Then
            
            Temat = CStr(Ws.Cells(i, "D").Value)
            
            If Len(Temat) > 0 Then
                
                Numer = WyciagnijNumerWniosku(Temat)
                
                If Numer <> "" Then
                    Ws.Cells(i, "A").Value = Numer
                    LicznikUzupelnionych = LicznikUzupelnionych + 1
                Else
                    LicznikBrakNumeru = LicznikBrakNumeru + 1
                End If
                
            End If
            
        End If
        
    Next i
    
Koniec:

    Application.ScreenUpdating = StaryScreenUpdating
    Application.Calculation = StaryCalculation
    Application.EnableEvents = StaryEnableEvents
    
    MsgBox "Uzupełnianie zakończone." & vbCrLf & vbCrLf & _
           "Uzupełniono numerów: " & LicznikUzupelnionych & vbCrLf & _
           "Nie znaleziono numeru: " & LicznikBrakNumeru, _
           vbInformation

    Exit Sub

Blad:

    Application.ScreenUpdating = StaryScreenUpdating
    Application.Calculation = StaryCalculation
    Application.EnableEvents = StaryEnableEvents
    
    MsgBox "Wystąpił błąd podczas uzupełniania:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbCritical

End Sub


'=========================================================
' GŁÓWNE MAKRO - POBIERANIE MAILI Z OUTLOOKA
'=========================================================
Sub PobierzMaile()

    Dim OutlookApp As Object
    Dim OutlookNS As Object
    Dim Folder As Object
    Dim Items As Object
    Dim Mail As Object
    
    Dim Ws As Worksheet
    Dim WsProblemy As Worksheet
    
    Dim Dict As Object
    
    Dim OstatniWiersz As Long
    Dim OstatniWierszProblemow As Long
    Dim NowyWiersz As Long
    
    Dim EntryID As String
    Dim Temat As String
    Dim Nadawca As String
    Dim NumerWniosku As String
    Dim DataMaila As Date
    
    Dim DataOd As Date
    
    Dim LicznikNowych As Long
    Dim LicznikPominietych As Long
    Dim LicznikProblemow As Long
    
    Dim CzasStart As Double
    
    Dim StaryScreenUpdating As Boolean
    Dim StaryCalculation As XlCalculation
    Dim StaryEnableEvents As Boolean
    
    On Error GoTo BladGlowny
    
    CzasStart = Timer
    
    '-----------------------------------------------------
    ' ARKUSZE
    '-----------------------------------------------------
    
    Set Ws = ThisWorkbook.Worksheets("Outlook")
    Set WsProblemy = ThisWorkbook.Worksheets("Problemy")
    
    '-----------------------------------------------------
    ' USTAWIENIA EXCELA
    '-----------------------------------------------------
    
    StaryScreenUpdating = Application.ScreenUpdating
    StaryCalculation = Application.Calculation
    StaryEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '-----------------------------------------------------
    ' DATA GRANICZNA
    '-----------------------------------------------------
    
    DataOd = DateSerial(2026, 6, 1)
    
    '-----------------------------------------------------
    ' SŁOWNIK ISTNIEJĄCYCH ENTRYID
    '-----------------------------------------------------
    
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbBinaryCompare
    
    OstatniWiersz = Ws.Cells(Ws.Rows.Count, "E").End(xlUp).Row
    
    If OstatniWiersz >= 2 Then
        
        Dim i As Long
        
        For i = 2 To OstatniWiersz
            
            EntryID = Trim(CStr(Ws.Cells(i, "E").Value))
            
            If EntryID <> "" Then
                If Not Dict.Exists(EntryID) Then
                    Dict.Add EntryID, True
                End If
            End If
            
        Next i
        
    End If
    
    '-----------------------------------------------------
    ' POŁĄCZENIE Z OUTLOOKIEM
    '-----------------------------------------------------
    
    On Error Resume Next
    
    Set OutlookApp = GetObject(, "Outlook.Application")
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    On Error GoTo BladGlowny
    
    If OutlookApp Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Nie udało się uruchomić Outlooka."
    End If
    
    Set OutlookNS = OutlookApp.GetNamespace("MAPI")
    
    '-----------------------------------------------------
    ' FOLDER OUTLOOK
    '-----------------------------------------------------
    
    Set Folder = OutlookNS.GetDefaultFolder(6).Folders("PEP Onboarding All")
    
    Set Items = Folder.Items
    
    'Najnowsze maile jako pierwsze
    Items.Sort "[ReceivedTime]", True
    
    '-----------------------------------------------------
    ' POBIERANIE MAILI
    '-----------------------------------------------------
    
    For Each Mail In Items
        
        'Sprawdzamy tylko wiadomości e-mail
        If Mail.Class = 43 Then
            
            On Error GoTo ProblemZMailem
            
            '---------------------------------------------
            ' DATA MAILA
            '---------------------------------------------
            
            DataMaila = Mail.ReceivedTime
            
            'Jeżeli jesteśmy już przed datą graniczną,
            'możemy zakończyć cały proces.
            If DataMaila < DataOd Then
                GoTo KoniecPetli
            End If
            
            '---------------------------------------------
            ' ENTRY ID
            '---------------------------------------------
            
            EntryID = Mail.EntryID
            
            '---------------------------------------------
            ' SPRAWDZENIE DUPLIKATU
            '---------------------------------------------
            
            If Dict.Exists(EntryID) Then
                
                LicznikPominietych = LicznikPominietych + 1
                
                GoTo NastepnyMail
                
            End If
            
            '---------------------------------------------
            ' ODCZYT PODSTAWOWYCH DANYCH
            '---------------------------------------------
            
            Nadawca = Mail.SenderName
            Temat = Mail.Subject
            
            '---------------------------------------------
            ' WYCIĄGNIĘCIE NUMERU WNIOSKU
            '---------------------------------------------
            
            NumerWniosku = WyciagnijNumerWniosku(Temat)
            
            '---------------------------------------------
            ' NOWY WIERSZ
            '---------------------------------------------
            
            NowyWiersz = Ws.Cells(Ws.Rows.Count, "E").End(xlUp).Row + 1
            
            Ws.Cells(NowyWiersz, "A").Value = NumerWniosku
            Ws.Cells(NowyWiersz, "B").Value = DataMaila
            Ws.Cells(NowyWiersz, "C").Value = Nadawca
            Ws.Cells(NowyWiersz, "D").Value = Temat
            Ws.Cells(NowyWiersz, "E").Value = EntryID
            
            'Dodajemy do słownika, żeby podczas tego samego
            'uruchomienia nie pobrać go drugi raz.
            Dict.Add EntryID, True
            
            LicznikNowych = LicznikNowych + 1
            
        End If
        
NastepnyMail:
        
        'Wracamy do normalnego działania obsługi błędów
        On Error GoTo BladGlowny
        
    Next Mail
    
KoniecPetli:

    '-----------------------------------------------------
    ' PRZYWRÓCENIE USTAWIEŃ
    '-----------------------------------------------------
    
    Application.ScreenUpdating = StaryScreenUpdating
    Application.Calculation = StaryCalculation
    Application.EnableEvents = StaryEnableEvents
    
    MsgBox "Pobieranie zakończone." & vbCrLf & vbCrLf & _
           "Nowe maile: " & LicznikNowych & vbCrLf & _
           "Pominięte (już istnieją): " & LicznikPominietych & vbCrLf & _
           "Problematyczne: " & LicznikProblemow & vbCrLf & vbCrLf & _
           "Czas: " & Format((Timer - CzasStart) / 86400, "hh:mm:ss"), _
           vbInformation
    
    Exit Sub


'=========================================================
' OBSŁUGA PROBLEMATYCZNEGO MAILA
'=========================================================

ProblemZMailem:

    On Error Resume Next
    
    'Nowy wiersz w arkuszu Problemy
    OstatniWierszProblemow = WsProblemy.Cells(WsProblemy.Rows.Count, "A").End(xlUp).Row + 1
    
    'Data
    WsProblemy.Cells(OstatniWierszProblemow, "A").Value = Mail.ReceivedTime
    
    'Nadawca
    WsProblemy.Cells(OstatniWierszProblemow, "B").Value = Mail.SenderName
    
    'Temat
    WsProblemy.Cells(OstatniWierszProblemow, "C").Value = Mail.Subject
    
    'BŁĄD
    WsProblemy.Cells(OstatniWierszProblemow, "D").Value = _
        Err.Number & " - " & Err.Description
    
    'EntryID
    WsProblemy.Cells(OstatniWierszProblemow, "E").Value = Mail.EntryID
    
    LicznikProblemow = LicznikProblemow + 1
    
    Err.Clear
    
    On Error GoTo BladGlowny
    
    GoTo NastepnyMail


'=========================================================
' BŁĄD GŁÓWNY MAKRA
'=========================================================

BladGlowny:

    Application.ScreenUpdating = StaryScreenUpdating
    Application.Calculation = StaryCalculation
    Application.EnableEvents = StaryEnableEvents
    
    MsgBox "Makro zostało przerwane." & vbCrLf & vbCrLf & _
           "Błąd: " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical

End Sub
