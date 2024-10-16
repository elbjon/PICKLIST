'Code is old and parts had to be left out
'Just part 1 -> Preparing the extracted data on PC


Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Global nr, G As Integer





Sub Zwischenspeichern()

    On Error GoTo Fehler
    ThisWorkbook.Save

    Exit Sub
Fehler:
    MsgBox "Automatische Speicherung hat nicht geklappt, wird der USB-Stick erkannt? Ihr könnt einfach weiterarbeiten"
End Sub


Sub Vorbereitung()
'ruft Spaltensortierung auf, ermittelt letzte Zeile, normalisiert Lagerorte, sortiert nach Lagerort, aber umständlich über einen generierten wert, ruft Speichern auf-entweder aud SPeedport oder freie Pfadwahl

Dim wort, zahl, buchstabe As String
Dim l As Integer
Dim j As Integer

i = 1
k = 2
l = 2 'zeile
j = 1 'spalte 


Call Spaltensortieren 'Reihenfolge EAN,Menge,Bezeichnung,VkEinheit,Hersteller,Lagerort,Bestand,tournr, WG, SummeVk"

Columns("A:A").Select
Selection.NumberFormat = "000000000000" ' EAN-Formatierung

letztezeile = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

'Sortieren nach BNN-Hersteller-Kürzel
'BNN-Hersteller-Kürzel ergänzen um Klarnamen
'wenn neues Kürzel, dann Select Case alles durchsuchen, Klarnamen eintragen, lastHrstll erneuern
    'Ausbaustufe: wenn nach select case lstHrstll unverändert, dann schreibe Kürzen und EAN in TXT-Datei oder sonstwohin
'Else Klarnamen aus vorheriger Zelle in neue Schreiben
'NEXT

'Dim lstHrstll As String
'lstHrstll = "XYX"

'Cells.Select
 '  Selection.Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlYes, _
  '      OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        

For G = 0 To letztezeile Step 1

If lstHrstll = Sheets(1).Cells(l + G, j + 4).Value Then
Sheets(1).Cells(l + G, j + 7) = Sheets(1).Cells(l + G - 1, j + 7)
Else



  

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'BNN-Herstellerkürzel werden an dieser Stelle durch Herstellernamen ersetzt
'Liste beim BNN erhaeltlich. Dieser Codeteil kann nicht veröffentlicht werden.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    End Select



   lstHrstll = Sheets(1).Cells(l + G, j + 4)
    
End If
    
    
    
Next G 



' Spaltenbenennungen setzen


Sheets(1).Cells(l - 1, j + 15) = "TourSortierhilfe"
Sheets(1).Cells(l - 1, j + 14) = "VkHof"
Sheets(1).Cells(l - 1, j + 13) = "Warengr."
Sheets(1).Cells(l - 1, j + 12) = "Tour"
Sheets(1).Cells(l - 1, j + 11) = "Intern. Sort."
Sheets(1).Cells(l - 1, j + 10) = "LAOT"
Sheets(1).Cells(l - 1, j + 9) = "Kommentar"
Sheets(1).Cells(l - 1, j + 8) = "gepackt"
Sheets(1).Cells(l - 1, j + 7) = "Hersteller"
Sheets(1).Cells(l - 1, j + 6) = "Bestand"
Sheets(1).Cells(l - 1, j + 5) = "Lagerort"
Sheets(1).Cells(l - 1, j + 4) = "Hersteller"
Sheets(1).Cells(l - 1, j + 3) = "Verp.-Größe"
Sheets(1).Cells(l - 1, j + 2) = "Bezeichnung"
Sheets(1).Cells(l - 1, j + 1) = "Packmenge"
Sheets(1).Cells(l - 1, j) = "EAN"


letztezeile = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row




Do Until l = letztezeile + 1
wort = Sheets(1).Cells(l, j + 5).Value


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Lagerortbezeichnungen normalisieren
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

If wort Like "[A-Za-z]#*" Then 'Lagerort normalisieren
    buchstabe = Left$(wort, 1) 'Buchstabe
    zahl = Mid$(wort, 2) ' Zahl

    
    
    
 If zahl Like "#####*" Then
    '   Sheets(1).Cells(l, j + 10) = buchstabe & zahl  '' Andere Lagerorte, Zahl zu lang
       
 End If
     If zahl Like "####" Then
    'Sheets(1).Cells(l, j + 10) = buchstabe & zahl  'Normalisierter Lagerort

      End If
    
    If zahl Like "###" Then
    zahl = "0" & zahl
    'Sheets(1).Cells(l, j + 10) = buchstabe & zahl  'Normalisierter Lagerort
    End If
    If zahl Like "##" Then
    zahl = "00" & zahl
    'Sheets(1).Cells(l, j + 10) = buchstabe & zahl  'Normalisierter Lagerort
    End If
    If zahl Like "#" Then
    zahl = "000" & zahl
 
    End If
    
     Sheets(1).Cells(l, j + 10) = buchstabe & zahl  'Normalisierter Lagerort
      
    Select Case True
    Case buchstabe Like "[Aa]"
        Sheets(1).Cells(l, j + 11) = "12" & "." & zahl
    Case buchstabe Like "[Bb]"
        Sheets(1).Cells(l, j + 11) = "14" & "." & zahl
    Case buchstabe Like "[Cc]"
        Sheets(1).Cells(l, j + 11) = "16" & "." & zahl
    Case buchstabe Like "[Dd]"
        Sheets(1).Cells(l, j + 11) = "18" & "." & zahl
    Case buchstabe Like "[Ee]"
        Sheets(1).Cells(l, j + 11) = "20" & "." & zahl
    Case buchstabe Like "[Ff]"
        Sheets(1).Cells(l, j + 11) = "22" & "." & zahl
    Case buchstabe Like "[Gg]"
        Sheets(1).Cells(l, j + 11) = "24" & "." & zahl
    Case buchstabe Like "[Hh]"
        Sheets(1).Cells(l, j + 11) = "36" & "." & zahl
    Case buchstabe Like "K"
        Sheets(1).Cells(l, j + 11) = "51" & "." & 10 - zahl
    Case buchstabe Like "[Ll]"
        Sheets(1).Cells(l, j + 11) = "34" & "." & zahl
    End Select
        
        
      
 
   
    Else
    Sheets(1).Cells(l, j + 10).Value = Sheets(1).Cells(l, j + 5).Value ' Andere Lagerorte ohne Zahl
        
        
Select Case Sheets(1).Cells(l, j + 10)
Case "Tiefkühl"
Sheets(1).Cells(l, j + 11) = "1"
Case "TK"
Sheets(1).Cells(l, j + 11) = "3"
Case "TK1"
Sheets(1).Cells(l, j + 11) = "4"
Case "TK2"
Sheets(1).Cells(l, j + 11) = "5"
Case "TK3"
Sheets(1).Cells(l, j + 11) = "6"
Case "TK4"
Sheets(1).Cells(l, j + 11) = "7"

Case "Tür"
Sheets(1).Cells(l, j + 11) = "11"


Case "K"
Sheets(1).Cells(l, j + 11) = "31"

Case "Mühle"
Sheets(1).Cells(l, j + 11) = "34"
Case "Haupt"
Sheets(1).Cells(l, j + 11) = "35"



Case "Tresen"
Sheets(1).Cells(l, j + 11) = "39"
Case "Echt Bio"
Sheets(1).Cells(l, j + 11) = "33"
Case "ANG"
Sheets(1).Cells(l, j + 11) = "33"

Case "Brot"
Sheets(1).Cells(l, j + 11) = "41"
Case "BR"
Sheets(1).Cells(l, j + 11) = "42"
Case "aS"
Sheets(1).Cells(l, j + 11) = "44"
Case "VB"
Sheets(1).Cells(l, j + 11) = "46"




Case ""
Sheets(1).Cells(l, j + 11) = "47.5"
Case "?"
Sheets(1).Cells(l, j + 11) = "47"
Case " "
Sheets(1).Cells(l, j + 11) = "48"
Case "L4/06"
Sheets(1).Cells(l, j + 11) = "49"
Case "ELB"
Sheets(1).Cells(l, j + 11) = "50"
Case "OG"
Sheets(1).Cells(l, j + 11) = "51"

Case "Käse"
Sheets(1).Cells(l, j + 11) = "53"
Case "KT"
Sheets(1).Cells(l, j + 11) = "55"
End Select
    

   If Left$(Sheets(1).Cells(l, j + 10), 2) = "AA" Then
        Sheets(1).Cells(l, j + 11) = "10"
    
    End If
    If Left$(Sheets(1).Cells(l, j + 10), 5) = "Kasse" Then
        Sheets(1).Cells(l, j + 11) = "32"
      End If
      
      

If Left$(Sheets(1).Cells(l, j + 10), 8) = "Käsethek" Then
        Sheets(1).Cells(l, j + 11) = "52"
      End If

   buchstabe = Z 'hier stand buchtabe
   zahl = 0

       End If
       




Select Case Sheets(1).Cells(l, j + 12).Value
    Case 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

    Sheets(1).Cells(l, j + 15) = "1"
    
Case 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

    Sheets(1).Cells(l, j + 15) = "2"
 Case'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

   Sheets(1).Cells(l, j + 15) = "3"
   
   Case 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   Sheets(1).Cells(l, j + 15) = "4"

    
    Case Else
    'If Sheets(1).Cells(l, j + 10).Value Like "3#" Then
    'Sheets(1).Cells(l, j + 14) = "4"
    'Else
    'End If
    Cells(l, j + 15) = "5"
    
End Select


l = l + 1

Loop ' Ende Lagerortnorm.

Cells.Select
   Selection.Sort Key1:=Range("P1"), Order1:=xlAscending, Key2:=Range("L1"), Order2:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        

        
        
        
        
        Call Speichern
       ' Application.Dialogs(xlDialogSaveAs).Show (sDatei)
End Sub

Sub SpeicherVersuch()
'
' Makro6 Makro
Dim sDatei, Test, sPfad, sZielDatei As String

  
  
  

  'Pfad = "\\speedport.ip\ALL\LexarEcho\Lexar\Naturel"
  'Pfad = "\\Arbeitsplatz"
  sPfad = "G:\"
  nr = nr + 1
  sDatei = Format(Date, "yyyy_MM_dd_") & "Nr_" & nr
  edg = ".xls"
  sZielDatei = sPfad & sDatei & edg
Application.GetSaveAsFilename sZielDatei
nr = nr + 1


' sPfad &
  
 'If Dir("Arbeitsplatz") = "" Then
    
  
  'MsgBox ("gibt den Pfad nicht!")
  'Application.Dialogs(xlDialogSaveAs).Show "G:\"
 'Application.Dialogs(xlDialogSaveAs).Show sDatei & edg
  
  
  '& sZielDatei
  'Else
  'MsgBox ("gibt den Pfad")
  
   
 ' End If



    '
    

    

    
End Sub

Sub Speichern()
    Dim sOrdner As String
    Dim sblattname As String
    Dim sFilename As String
   
    sOrdner = "G:\"
    sblattname = Format(Date, "yyyy_MM_dd_") & "Nr_" & nr
   
    sFilename = Application.GetSaveAsFilename _
    (sOrdner & sblattname, "Microsoft Excel-Dateien (*.xlsx),*.xlsx")
   
   ' If sFilename <> False Then
        ActiveWorkbook.SaveAs sFilename _
        , FileFormat:=51, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
     '   nr = nr + 1
    'End If

End Sub


Sub Spaltensortieren()

    
    
       Cells.Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "A1:Z1"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "EAN,Menge,Bezeichnung,VkEinheit,Hersteller,Lagerort,Bestand,tournr, WG, SummeVk", DataOption:= _
        xlSortNormal
        

    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:Z" & Cells(Rows.Count, 1).End(xlUp).Row) 'ist sicherer, da leere Spalten bis "Z" berücksichtigt werden
        '("A1:Z" & Cells(Rows.Count, 1).End(xlUp).Row)
        '("A1").CurrentRegion
        .Header = xlTrue 'xlGuess
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
    
      Columns("H:J").Select
    Selection.Cut
    Columns("M:O").Select
    ActiveSheet.Paste
    
    Columns("P:Z").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub

Sub Test()
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub


Sub SortienrungBEGINN()
'
' SortienrungBEGINN Makro
'

'
    Cells.Select
    Range("M54").Activate
    
    Application.AddCustomList ListArray:=Array('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
    ActiveWorkbook.Worksheets("$I66156512498855").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("$I66156512498855").Sort.SortFields.Add Key:=Range( _
        "M2:M124"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("$I66156512498855").Sort
        .SetRange Range("A1:O124")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    

    Rows("2:7").Select
    ActiveWorkbook.Worksheets("$I66156512498855").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("$I66156512498855").Sort.SortFields.Add Key:=Range( _
        "L2:L7"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("$I66156512498855").Sort
        .SetRange Range("A2:O7")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    
    
    
    
    
    Rows("2:7").Select
    Range("A7").Activate
    ActiveWorkbook.Worksheets("$I66156512498855").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("$I66156512498855").Sort.SortFields.Add Key:=Range( _
        "L2:L7"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("$I66156512498855").Sort
        .SetRange Range("A2:O7")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub








