Attribute VB_Name = "Module1"
Sub nettoyerCSV_crawled()
' nettoyerCSV_crawled Macro
'
' Touche de raccourci du clavier: Ctrl+Shift+N
        Workbooks.Add

        Dim requete As String
        Dim domaine As String
        
        requete = "sante_mtl"
        path = "C:\Users\p1115145\Documents\text-mining-project\03-corpus\1-crawler\"
        Source = path & requete & ".csv"
        
    ActiveWorkbook.Queries.Add Name:=requete, Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""" & Source & """),[Delimiter=""#(tab)"", Columns=15, Encoding=65001, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""En-t�tes promus"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Type modifié"" = Table.TransformColumnTypes(#""En-têtes promus"",{{""Address"", type text}, {""Status-Co" & _
        "de"", Int64.Type}, {""Status-Text"", type text}, {""Type"", type text}, {""Size"", Int64.Type}, {""Title"", type text}, {""Date"", type text}, {""Level"", Int64.Type}, {""Links Out"", Int64.Type}, {""Links In"", Int64.Type}, {""Server"", type text}, {""Error"", type text}, {""Duration"", type text}, {""Charset"", type text}, {""Description"", type text}})," & Chr(13) & "" & Chr(10) & "    #""L" & _
        "ignes filtrées"" = Table.SelectRows(#""Type modifié"", each ([#""Status-Text""] <> ""skip external""))," & Chr(13) & "" & Chr(10) & "    #""Autres colonnes supprim�es"" = Table.SelectColumns(#""Lignes filtr�es"",{""Address"", ""Type"", ""Title"", ""Charset"", ""Description""})," & Chr(13) & "" & Chr(10) & "    #""Lignes filtrées1"" = Table.SelectRows(#""Autres colonnes supprimées"", each ([Type] = """" or [Type] = ""appl" & _
        "ication/pdf"" or [Type] = ""text/html""))" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Lignes filtr�es1"""
    
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & requete & """;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & requete & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = requete
        .Refresh BackgroundQuery:=False
    End With


    Application.CommandBars("Queries and Connections").Visible = False
    Application.DisplayAlerts = False
    ActiveSheet.ListObjects(requete).TableStyle = ""
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Columns("C:C").ColumnWidth = 54
    

    'Supprimer la feuille qui ne nous servira plus
    Sheets("Feuil1").Select
    ActiveWindow.SelectedSheets.Delete
    
    Sheets(1).Name = requete
    
    
    'Enregistrer le classeur
    Range("A1").Select
    Dim save As String
    save = requete & ".xlsx"
    ActiveWorkbook.SaveAs fileName:= _
        path & save, FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
    
End Sub




