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
    
    Dim i As Long
    Dim LiczbaElementow As Long
    
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
    
    '=====================================================
    ' ARKUSZE
    '=====================================================
    
    Set Ws = ThisWorkbook.Worksheets("Outlook")
    Set WsProblemy = ThisWorkbook.Worksheets("Problemy")
    
    '=====================================================
    ' USTAWIENIA EXCELA
    '=====================================================
    
    StaryScreenUpdating = Application.ScreenUpdating
    StaryCalculation = Application.Calculation
    StaryEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '=====================================================
    ' DATA GRANICZNA
    '=====================================================
    
    DataOd = DateSerial(2026, 6, 1)
    
    '=====================================================
    ' SŁOWNIK ISTNIEJĄCYCH ENTRYID
    '=====================================================
    
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbBinaryCompare
    
    OstatniWiersz = Ws.Cells(Ws.Rows.Count, "E").End(xlUp).Row
    
    If OstatniWiersz >= 2 Then
        
        For i = 2 To OstatniWiersz
            
            EntryID = Trim(CStr(Ws.Cells(i, "E").Value))
            
            If EntryID <> "" Then
                
                If Not Dict.Exists(EntryID) Then
                    Dict.Add EntryID, True
                End If
                
            End If
            
        Next i
        
    End If
    
    '=====================================================
    ' POŁĄCZENIE Z OUTLOOKIEM
    '=====================================================
    
    On Error Resume Next
    
    Set OutlookApp = GetObject(, "Outlook.Application")
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    On Error GoTo BladGlowny
    
    If OutlookApp Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
                  "Nie udało się uruchomić Outlooka."
    End If
    
    Set OutlookNS = OutlookApp.GetNamespace("MAPI")
    
    '=====================================================
    ' FOLDER
    '=====================================================
    
    Set Folder = OutlookNS.GetDefaultFolder(6).Folders("PEP Onboarding All")
    
    Set Items = Folder.Items
    
    'Najnowsze maile jako pierwsze
    Items.Sort "[ReceivedTime]", True
    
    LiczbaElementow = Items.Count
    
    '=====================================================
    ' PRZEJŚCIE PO MAILACH
    '=====================================================
    
    For i = 1 To LiczbaElementow
        
        '-------------------------------------------------
        ' BARDZO WAŻNE:
        ' Obsługę błędu włączamy PRZED odczytem maila.
        '-------------------------------------------------
        
        On Error GoTo ProblemZMailem
        
        Set Mail = Items.Item(i)
        
        '-------------------------------------------------
        ' Sprawdzamy czy to wiadomość e-mail
        '-------------------------------------------------
        
        If Mail.Class <> 43 Then
            GoTo NastepnyMail
        End If
        
        '-------------------------------------------------
        ' DATA
        '-------------------------------------------------
        
        DataMaila = Mail.ReceivedTime
        
        'Jeżeli jesteśmy już przed 01.06.2026,
        'kończymy dalsze przeglądanie.
        If DataMaila < DataOd Then
            GoTo KoniecPetli
        End If
        
        '-------------------------------------------------
        ' ENTRY ID
        '-------------------------------------------------
        
        EntryID = Mail.EntryID
        
        '-------------------------------------------------
        ' SPRAWDZENIE DUPLIKATU
        '-------------------------------------------------
        
        If Dict.Exists(EntryID) Then
            
            LicznikPominietych = LicznikPominietych + 1
            
            GoTo NastepnyMail
            
        End If
        
        '-------------------------------------------------
        ' DANE MAILA
        '-------------------------------------------------
        
        Nadawca = Mail.SenderName
        Temat = Mail.Subject
        
        '-------------------------------------------------
        ' NUMER WNIOSKU
        '-------------------------------------------------
        
        NumerWniosku = WyciagnijNumerWniosku(Temat)
        
        '-------------------------------------------------
        ' NOWY WIERSZ W EXCELU
        '-------------------------------------------------
        
        NowyWiersz = Ws.Cells(Ws.Rows.Count, "E").End(xlUp).Row + 1
        
        Ws.Cells(NowyWiersz, "A").Value = NumerWniosku
        Ws.Cells(NowyWiersz, "B").Value = DataMaila
        Ws.Cells(NowyWiersz, "C").Value = Nadawca
        Ws.Cells(NowyWiersz, "D").Value = Temat
        Ws.Cells(NowyWiersz, "E").Value = EntryID
        
        'Dodajemy EntryID do słownika
        Dict.Add EntryID, True
        
        LicznikNowych = LicznikNowych + 1
        
NastepnyMail:

        'Reset obsługi błędu przed kolejnym elementem
        On Error GoTo BladGlowny
        
    Next i


KoniecPetli:

    '=====================================================
    ' PRZYWRÓCENIE USTAWIEŃ
    '=====================================================
    
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
' PROBLEM Z KONKRETNYM MAILEM
'=========================================================

ProblemZMailem:

    Dim NumerBledu As Long
    Dim OpisBledu As String
    
    'Zapamiętujemy błąd zanim go wyczyścimy
    NumerBledu = Err.Number
    OpisBledu = Err.Description
    
    '-----------------------------------------------------
    ' Tutaj NIE zakładamy, że właściwości maila
    ' na pewno dadzą się odczytać.
    ' Każdy zapis próbujemy osobno.
    '-----------------------------------------------------
    
    On Error Resume Next
    
    OstatniWierszProblemow = _
        WsProblemy.Cells(WsProblemy.Rows.Count, "A").End(xlUp).Row + 1
    
    'Data
    Err.Clear
    WsProblemy.Cells(OstatniWierszProblemow, "A").Value = Mail.ReceivedTime
    
    'Nadawca
    Err.Clear
    WsProblemy.Cells(OstatniWierszProblemow, "B").Value = Mail.SenderName
    
    'Temat
    Err.Clear
    WsProblemy.Cells(OstatniWierszProblemow, "C").Value = Mail.Subject
    
    'Opis błędu
    WsProblemy.Cells(OstatniWierszProblemow, "D").Value = _
        NumerBledu & " - " & OpisBledu
    
    'EntryID
    Err.Clear
    WsProblemy.Cells(OstatniWierszProblemow, "E").Value = Mail.EntryID
    
    LicznikProblemow = LicznikProblemow + 1
    
    Err.Clear
    
    'Wracamy do normalnej obsługi
    On Error GoTo BladGlowny
    
    GoTo NastepnyMail


'=========================================================
' GŁÓWNY BŁĄD MAKRA
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
