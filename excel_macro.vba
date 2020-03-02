Sub XML_Export()
Dim strDateiname As String
Dim strDateinameZusatz As String
Dim strMappenpfad As String
Dim intCutExt

'Datename ohne Ext. (nach Punkt suchen):
intCutExt = Len(ActiveWorkbook.Name) - InStrRev(ActiveWorkbook.Name, ".") + 1
strMappenpfad = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - intCutExt)

'strDateinameZusatz = "-" & Year(ActiveSheet.Cells(3, 1).Value) & "-" & Month(ActiveSheet.Cells(3, 1).Value) & ".xml"
strDateinameZusatz = "-" & Format(Now, "YYYY-MM-DD-HH-MM-SS") & ".xml"

strDateiname = InputBox("Bitte den Namen der XML-Datei angeben.", "XML-Export", strMappenpfad & strDateinameZusatz)
If strDateiname = "" Then Exit Sub

Range("A2").Select

'Erstellt die Telefonbuchdatei (hier: xxx.xml)
'Dateiname kann frei gewählt werden
'Der entsprechende Ordner MUSS vorhanden sein, da sonst ein Fehler auftritt
    Set fs = CreateObject("scripting.filesystemobject")
    
    Set a = fs.createtextfile(strDateiname, True)

'Schreibt den allgemeinen Teil der Telefonbuchdatei
    a.writeline ("<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>")
'    a.writeline ("<?xml version=" & """1.0""" & " encoding=" & """UTF-16""" & "?>")
'    a.writeline ("<?xml version=" & """1.0""" & " encoding=" & """ISO-8859-1""" & "?>")
    a.writeline ("<phonebooks>")
    a.writeline ("<phonebook>")
    'a.writeline ("<phonebook name=" & """Telefonbuch 1""" & " owner=" & """1""" & ">")

'Schleife zur Ermittlung aller Einträge
'Benutzt alle Datensätze, die einen Namen enthalten
    i = 0
    While ActiveCell.Offset(i, 0) <> ""
    
    Dim realName As String
    realName = Umlaut(ActiveCell.Offset(i, 0))
    Dim home As String
    home = ActiveCell.Offset(i, 1)
    Dim work As String
    work = ActiveCell.Offset(i, 2)
    Dim mobile As String
    mobile = ActiveCell.Offset(i, 3)
    Dim fax_work As String
    fax_work = ActiveCell.Offset(i, 4)

'Schreibt den Telefonbucheintrag
    a.writeline ("<contact><category>0</category>")
    a.writeline ("<person><realName>" + realName + "</realName></person><telephony>")
    If home <> "" Then
        a.writeline ("<number type=" & """home""" & " prio=" & """1""" & " id=" & """0""" & ">" + home + "</number>")
    End If
    If work <> "" Then
        a.writeline ("<number type=" & """work""" & " prio=" & """1""" & " id=" & """1""" & ">" + work + "</number>")
    End If
    If mobile <> "" Then
        a.writeline ("<number type=" & """mobile""" & " prio=" & """1""" & " id=" & """2""" & ">" + mobile + "</number>")
    End If
    If fax_work <> "" Then
        a.writeline ("<number type=" & """fax_work""" & " prio=" & """1""" & " id=" & """3""" & ">" + fax_work + "</number>")
    End If
    a.writeline ("</telephony></contact>")
    
    i = i + 1
    Wend
'Ende der Schleife
    
'Ende der Telefonbuchdatei
    a.writeline ("</phonebook>")
    a.writeline ("</phonebooks>")
    
MsgBox "Export erfolgreich. Datei wurde exportiert nach" & vbCrLf & strDateiname
End Sub

Public Function Umlaut(Anything As Variant) As Variant
' https://dbwiki.net/wiki/VBA_Tipp:_Umlaute_ersetzen
   Dim i        As Long
   Dim Ch       As String * 1
   Dim Ch1      As String * 1
   Dim Res      As String
   Dim IsUpCase As Boolean
 
   If IsNull(Anything) Then Umlaut = Null: Exit Function
 
   For i = 1 To Len(Anything)
      Ch = Mid$(Anything, i, 1)
      Ch1 = IIf(i < Len(Anything), Mid$(Anything, i + 1, 1), " ")
      ' Nächstes Zeichen ist kein Kleinbuchstabe:
      IsUpCase = CBool((Asc(Ch1) = Asc(UCase(Ch1))))
      Select Case Asc(Ch)
         Case Asc("Ä"): Res = Res & "Ã„"
         Case Asc("Ü"): Res = Res & "Ãœ"
         Case Asc("Ö"): Res = Res & "Ã–"
         Case Asc("ü"): Res = Res & "Ã¼"
         Case Asc("ö"): Res = Res & "Ã¶"
         Case Asc("ä"): Res = Res & "Ã¤"
         Case Asc("ß"): Res = Res & "ÃŸ"
         Case Else:     Res = Res & Ch
      End Select
   Next
   Umlaut = Res
End Function
