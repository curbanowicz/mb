Sub PobierzJIRA()

    Dim OutlookApp As Object
    Dim OutlookNS As Object
    Dim JIRAFolder As Object
    Dim Items As Object
    Dim Mail As Object
    
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Licznik As Long
    
    Dim OstatniaData As Date
    Dim PierwszeUruchomienie As Boolean
    
    Set ws = ThisWorkbook.Worksheets("Test_JIRA")
    
    '----------------------------------------
    ' POŁĄCZENIE Z OUTLOOKIEM
    '----------------------------------------
    
    On Error Resume Next
    
    Set OutlookApp = GetObject(, "Outlook.Application")
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    On Error GoTo 0
    
    Set OutlookNS = OutlookApp.GetNamespace("MAPI")
    
    '----------------------------------------
    ' FOLDER JIRA
    '----------------------------------------
    
    Set JIRAFolder = OutlookNS.GetDefaultFolder(6).Folders("JIRA")
    
    If JIRAFolder Is Nothing Then
        MsgBox "Nie znaleziono folderu JIRA.", vbExclamation
        Exit Sub
    End If
    
    Set Items = JIRAFolder.Items
    
    '----------------------------------------
    ' SORTOWANIE OD NAJNOWSZYCH
    '----------------------------------------
    
    Items.Sort "[ReceivedTime]", True
    
    '----------------------------------------
    ' SPRAWDZENIE, CZY TO PIERWSZE URUCHOMIENIE
    '----------------------------------------
    
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If LastRow < 2 Or Not IsDate(ws.Cells(2, 1).Value) Then
        
        PierwszeUruchomienie = True
        
    Else
        
        PierwszeUruchomienie = False
        
        'Najstarsza data już pobrana
        OstatniaData = ws.Cells(LastRow, 1).Value
        
    End If
    
    '----------------------------------------
    ' PIERWSZY WOLNY WIERSZ
    '----------------------------------------
    
    LastRow = LastRow + 1
    
    Licznik = 0
    
    '----------------------------------------
    ' POBIERANIE MAILI
    '----------------------------------------
    
    For Each Mail In Items
        
        If Mail.Class = 43 Then
            
            'Jeżeli to nie pierwsze uruchomienie
            'i dotarliśmy do starych wiadomości
            If Not PierwszeUruchomienie Then
                
                If Mail.ReceivedTime < OstatniaData Then
                    Exit For
                End If
                
            End If
            
            '--------------------------------
            ' ZAPIS DO EXCELA
            '--------------------------------
            
            ws.Cells(LastRow, 1).Value = Mail.ReceivedTime
            ws.Cells(LastRow, 2).Value = Mail.SenderName
            ws.Cells(LastRow, 3).Value = Mail.Subject
            ws.Cells(LastRow, 4).Value = Mail.EntryID
            
            LastRow = LastRow + 1
            Licznik = Licznik + 1
            
        End If
        
    Next Mail
    
    '----------------------------------------
    ' FORMATOWANIE
    '----------------------------------------
    
    ws.Columns("A:D").AutoFit
    
    '----------------------------------------
    ' INFORMACJA
    '----------------------------------------
    
    If PierwszeUruchomienie Then
        
        MsgBox "Pierwsze pobieranie zakończone." & vbCrLf & _
               "Pobrano: " & Licznik & " wiadomości.", _
               vbInformation
        
    Else
        
        MsgBox "Pobrano nowych wiadomości: " & Licznik, _
               vbInformation
        
    End If

End Sub
