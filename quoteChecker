Option Explicit

'global variables to be used in popup and sub routines
Global MarshallProfiling As Boolean
Global RealSteel As Boolean
'sub routine to clear data from the quote comp sheet
Sub ClearTables()

Dim CompBook As Workbook: Set CompBook = ThisWorkbook
       
    'Declear worksheets
    Dim QuoteCheck As Worksheet
    Set QuoteCheck = CompBook.Worksheets("Quote Check")
    
    'Declear listobjects
    Dim BOMCopy As ListObject
    Dim Quote As ListObject
    Set BOMCopy = QuoteCheck.ListObjects("BOMCopy")
    Set Quote = QuoteCheck.ListObjects("Quote")
    
    'clear the data range of both tables and reset the colour
    If Not BOMCopy.DataBodyRange Is Nothing Then
        BOMCopy.DataBodyRange.Interior.Color = RGB(38, 38, 38)
        BOMCopy.DataBodyRange.Delete
    End If
    If Not Quote.DataBodyRange Is Nothing Then
        Quote.DataBodyRange.Interior.Color = RGB(38, 38, 38)
        Quote.DataBodyRange.Delete
    End If
    'clear text from the output
    Dim output As Shape
    Set output = QuoteCheck.Shapes("output")
    output.TextFrame.Characters.Delete

End Sub
'sub that imports the profiles from the bom worksheet
Sub ImportData()
    
    Dim CompBook As Workbook: Set CompBook = ThisWorkbook
    'Declear worksheets
    Dim QuoteCheck As Worksheet
    Dim MBOM As Worksheet
    Set QuoteCheck = CompBook.Worksheets("Quote Check")
    Set MBOM = CompBook.Worksheets("Machine BOM")
    
    'Declear listobjects
    Dim BOMCopy As ListObject
    Dim Quote As ListObject
    Set BOMCopy = QuoteCheck.ListObjects("BOMCopy")
    Set Quote = QuoteCheck.ListObjects("Quote")
    
    'Copy required columns from orginal bom sheet
    Dim i As Integer
    i = 1
    Dim cols As Integer
    cols = MBOM.ListObjects("MBOM").ListColumns.count
    
    'get the range required for the new table
    Dim colr As Integer
    colr = QuoteCheck.ListObjects("BOMCopy").ListColumns.count
    Dim rows As Integer
    rows = MBOM.ListObjects("MBOM").ListRows.count
        
    'import bom to an array
    Dim MBOMArray() As Variant
    MBOMArray = MBOM.ListObjects("MBOM").DataBodyRange.Value
    
    'declear bom to be pasted into quote comp sheet
    Dim BOMFinalArray() As Variant
    Dim c As Integer
    c = 0
    
    'get the number of rows to dimension the final array that will be pasted into the comp book
    For i = 1 To UBound(MBOMArray)
        If Not IsEmpty(MBOMArray(i, 11)) = True Then
            c = c + 1
        End If
    Next i
    
    If c = 0 Then
        MsgBox "Profile column needs to be filled"
        Exit Sub
    End If
    
    'size the array
    ReDim Preserve BOMFinalArray(1 To c, 1 To colr)
    
    'put data into array if the part is profiled
    c = 1
    For i = 1 To UBound(MBOMArray)
        If Not IsEmpty(MBOMArray(i, 11)) = True Then
                BOMFinalArray(c, 1) = MBOMArray(i, 2)
                BOMFinalArray(c, 2) = MBOMArray(i, 4)
                BOMFinalArray(c, 3) = MBOMArray(i, 5)
                BOMFinalArray(c, 4) = MBOMArray(i, 7)
                BOMFinalArray(c, 5) = MBOMArray(i, 9)
                BOMFinalArray(c, 6) = MBOMArray(i, 12)
                BOMFinalArray(c, 7) = MBOMArray(i, 13)
                BOMFinalArray(c, 8) = MBOMArray(i, 14)
                BOMFinalArray(c, 9) = MBOMArray(i, 17)
                c = c + 1
        End If
    Next i

    'resize the BOMCopy table so that the required columns can be copied to it
    BOMCopy.Resize BOMCopy.range.Resize(UBound(BOMFinalArray) + 1, colr)
    
    'paste array into table
    QuoteCheck.ListObjects("BOMCopy").DataBodyRange.Value = BOMFinalArray()
                     
    'quote type selection popup
    QuoteSelect.Show
    
    'call sub depending on user selection
    If MarshallProfiling = True Then
        Call MarshallsDataImport
    End If
    If RealSteel = True Then
        Call RealSteelImport
    End If
    
    'call the sub routine to tidy the section and material columns
    Call SectionClean


End Sub
    
Sub MarshallsDataImport()

    Dim CompBook As Workbook: Set CompBook = ThisWorkbook
       
    'Declear worksheets
    Dim QuoteCheck As Worksheet
    Set QuoteCheck = CompBook.Worksheets("Quote Check")
    
    'Declear listobjects
    Dim BOMCopy As ListObject
    Dim Quote As ListObject
    Set BOMCopy = QuoteCheck.ListObjects("BOMCopy")
    Set Quote = QuoteCheck.ListObjects("Quote")

    Dim QuoteFilePath As String
    Dim QuoteWorkbook As Workbook
    Dim QuoteSheet As Worksheet
    Dim PNrow As Integer
    Dim PNrowString As String
    
 
        'Get the quote file name via pop up window
        QuoteFilePath = Application.GetOpenFilename
    
        'Make the CSV a workbook that can be referenced and define what sheet is to be used
        Set QuoteWorkbook = Workbooks.Open(QuoteFilePath)
        QuoteWorkbook.Activate
        Set QuoteSheet = QuoteWorkbook.Sheets(1)
        
        'find the description col so that part num can be cleaned
        Dim csvrawArray() As Variant
        csvrawArray() = QuoteSheet.range("A1").CurrentRegion
        Dim h As Integer
        Dim descriptioncol As Integer
        
        For h = 1 To UBound(csvrawArray, 2)
                If InStr(1, csvrawArray(1, h), "Description", 1) > 0 Then
                descriptioncol = h
                End If
        Next h
        
        
        'Clean up part number in CSV
        'create part num array
        Dim PartNumArray() As Variant
        
        'create range to get dynamic size
        Dim PartNumRng As range
        Set PartNumRng = Cells(1, descriptioncol).CurrentRegion
        
        'get number of rows
        PNrow = PartNumRng.CurrentRegion.rows.count
        
        'declear partnum array
        PartNumArray() = QuoteSheet.range(Cells(2, descriptioncol), Cells(PNrow, descriptioncol))
        
        Dim PartNumArrayClean() As Variant
        
        'Creating regular expression object
        Dim regEx As New RegExp
        regEx.Global = True
           
        'regex expression for identifing pfs part number
        regEx.Pattern = "(.*((-)(\d{4})(-)(\d{3,4})(-\d{3})?))"
        Dim item As Variant
        Dim text As Variant
        Dim x As Integer
        x = 1
        'clean the part numbers
        For Each text In PartNumArray()
            Dim mc As MatchCollection
            Set mc = regEx.Execute(text)
                For Each item In mc
                    ReDim Preserve PartNumArrayClean(1 To x)
                    PartNumArrayClean(x) = item
                    x = x + 1
                Next item
        Next text
        
        'Put clean part numbers back into CSV
        QuoteSheet.Cells(1, h).Value = "Part Number"
        QuoteSheet.Cells(2, h).Resize(PNrow - 1).Value = Application.Transpose(PartNumArrayClean)
    
        'Import the quote to an array
        Dim QuoteImportArray() As Variant
        QuoteImportArray() = QuoteSheet.range("A1").CurrentRegion
                
        'gets last row and column number
        Dim QuoteNumRows As Integer
        Dim QuoteNumCols As Integer
        QuoteNumRows = UBound(QuoteImportArray, 1)
        QuoteNumCols = UBound(QuoteImportArray, 2)
        
        Dim partnumcol As Integer
        Dim sectioncol As Integer
        Dim materialcol As Integer
        Dim quantitycol As Integer
        Dim processcol As Integer
        Dim s As Integer
        
        For s = 1 To QuoteNumCols
            
            If InStr(1, QuoteImportArray(1, s), "Part Number", 1) > 0 Then
                partnumcol = s
            End If
            If InStr(1, QuoteImportArray(1, s), "Thickness", 1) > 0 Then
                sectioncol = s
            End If
            If InStr(1, QuoteImportArray(1, s), "Grade", 1) > 0 Then
                materialcol = s
            End If
            If InStr(1, QuoteImportArray(1, s), "Quantity", 1) > 0 Then
                quantitycol = s
            End If
            If InStr(1, QuoteImportArray(1, s), "Process", 1) > 0 Then
                processcol = s
            End If
        Next s
        
        'generates checking quote by only getting required columns from import array
        Dim j As Integer
        Dim c As Integer
        Dim i As Integer
        Dim QuoteCheckArray() As Variant
        ReDim QuoteCheckArray(1 To QuoteNumRows - 1, 1 To 5)
        
            For i = 1 To QuoteNumRows - 1
                For j = 1 To QuoteNumCols
                    c = 1
                    If j = processcol Then
                        QuoteCheckArray(i, c) = QuoteImportArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = sectioncol Then
                        QuoteCheckArray(i, c) = QuoteImportArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = materialcol Then
                        QuoteCheckArray(i, c) = QuoteImportArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = quantitycol Then
                        QuoteCheckArray(i, c) = QuoteImportArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = partnumcol Then
                            QuoteCheckArray(i, c) = QuoteImportArray(i + 1, j)
                    End If
                Next j
            Next i
        
        'close the csv and active the quote checker
        CompBook.Activate
        QuoteWorkbook.Close savechanges:=True
                  
        Dim colr As Integer
        colr = QuoteCheck.ListObjects("BOMCopy").ListColumns.count
        'resize the quote table so that the data can be pasted in the table
        Quote.Resize Quote.range.Resize(QuoteNumRows, colr)
        
        'paste the quote array in desired order in the quote check sheet
        With QuoteCheck
            'part number
            .ListObjects("Quote").ListColumns(1).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 5)
            'section
            .ListObjects("Quote").ListColumns(2).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 2)
            'grade
            .ListObjects("Quote").ListColumns(3).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 3)
            'process
            .ListObjects("Quote").ListColumns(4).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 1)
            'quantity
            .ListObjects("Quote").ListColumns(5).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 4)
        End With

         
        For i = 1 To [Quote].ListRows.count
            For j = 6 To [Quote].ListColumns.count
                If [Quote].DataBodyRange.Cells(i, j) = "" Then
                    [Quote].DataBodyRange.Cells(i, j) = "N"
                End If
            Next j
        Next i
                
        

End Sub
    
Sub RealSteelImport()
        Dim CompBook As Workbook: Set CompBook = ThisWorkbook
       
        'Declear worksheets
        Dim QuoteCheck As Worksheet
        Set QuoteCheck = CompBook.Worksheets("Quote Check")
        Dim QuoteFilePath As String
        Dim QuoteWorkbook As Workbook
        Dim QuoteSheet As Worksheet
        Dim PNrow As Integer
        Dim PNrowString As String
        'Declear listobjects
        Dim BOMCopy As ListObject
        Dim Quote As ListObject
        Set BOMCopy = QuoteCheck.ListObjects("BOMCopy")
        Set Quote = QuoteCheck.ListObjects("Quote")
        QuoteFilePath = Application.GetOpenFilename
    
        'Make the CSV a workbook that can be referenced and define what sheet is to be used
        Set QuoteWorkbook = Workbooks.Open(QuoteFilePath)
        QuoteWorkbook.Activate
        Set QuoteSheet = QuoteWorkbook.Sheets(1)
       
        'create part num array
        Dim ProcessingArray() As Variant
        
        'create range to get dynamic size
        Dim ProcessingRng As range
        Set ProcessingRng = range("K2").CurrentRegion
        
        'get number of rows
        PNrow = ProcessingRng.CurrentRegion.rows.count
        
        'create range out of string so that it can be called
        PNrowString = "K2:K" & PNrow
        
        'declear partnum array
        ProcessingArray() = QuoteSheet.range(PNrowString)
        
        Dim SplitArray() As Variant
        ReDim SplitArray(1 To UBound(ProcessingArray), 1 To 4)
        
        
        'for loop searching for additional processing on part split in to separate columns
        Dim x As Integer
        For x = 1 To UBound(ProcessingArray)
            If Not IsEmpty(ProcessingArray(x, 1)) Then
                'check if it is formed
                If InStr(1, CStr(ProcessingArray(x, 1)), "Formed", 1) > 0 Then
                    SplitArray(x, 1) = "Y"
                End If
                'check if it tapped
                If InStr(1, CStr(ProcessingArray(x, 1)), "Tap", 1) > 0 Then
                    SplitArray(x, 2) = "Y"
                End If
                'check if it is machined
                If InStr(1, CStr(ProcessingArray(x, 1)), "Machined", 1) > 0 Then
                    SplitArray(x, 3) = "Y"
                End If
                'check if it is beveled
                If InStr(1, CStr(ProcessingArray(x, 1)), "Bevel", 1) > 0 Then
                    SplitArray(x, 4) = "Y"
                End If
            End If
        Next x
        
        'add headings to the section
        QuoteSheet.range("L1").Value = "Pressed"
        QuoteSheet.range("M1").Value = "Drill/Tap"
        QuoteSheet.range("N1").Value = "Machined"
        QuoteSheet.range("O1").Value = "Bevel"
        'add array with attional processing to csv
        QuoteSheet.range("L2").Resize(PNrow - 1, 4).Value = SplitArray
        
        'import all data into array
        Dim RealSteelQuoteArray() As Variant
        'import all data
        RealSteelQuoteArray() = QuoteSheet.range("A1").CurrentRegion
        
        'gets last row and column number
        Dim QuoteNumRows As Integer
        Dim QuoteNumCols As Integer
        QuoteNumRows = UBound(RealSteelQuoteArray, 1)
        QuoteNumCols = UBound(RealSteelQuoteArray, 2)
        
        Dim partnumcol As Integer
        Dim materialcol As Integer
        Dim sectioncol As Integer
        Dim quantitycol As Integer
        Dim processcol As Integer
        Dim pressedcol As Integer
        Dim drilltapcol As Integer
        Dim machinedcol As Integer
        Dim bevelcol As Integer
        Dim s As Integer
        
        For s = 1 To QuoteNumCols
            
            If InStr(1, RealSteelQuoteArray(1, s), "Short Description", 1) > 0 Then
                partnumcol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Material", 1) > 0 Then
                materialcol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Thickness", 1) > 0 Then
                sectioncol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "QTY", 1) > 0 Then
                quantitycol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Cut on", 1) > 0 Then
                processcol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Pressed", 1) > 0 Then
                pressedcol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Drill/Tap", 1) > 0 Then
                drilltapcol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Machined", 1) > 0 Then
                machinedcol = s
            End If
            If InStr(1, RealSteelQuoteArray(1, s), "Bevel", 1) > 0 Then
                bevelcol = s
            End If
            
        Next s
                
        'create final array to be pasted into quote comp table, j variable indicates the column in the csv that is required
        Dim j As Integer
        Dim c As Integer
        Dim i As Integer
        Dim QuoteCheckArray() As Variant
        ReDim QuoteCheckArray(1 To QuoteNumRows - 1, 1 To 12)
        
            For i = 1 To QuoteNumRows - 1
                For j = 1 To QuoteNumCols
                    c = 1
                    If j = partnumcol Then
                        QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = materialcol Then
                        QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = sectioncol Then
                        QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = quantitycol Then
                        QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = processcol Then
                            QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = pressedcol Then
                            QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = drilltapcol Then
                            QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = machinedcol Then
                            QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                    c = c + 1
                    If j = bevelcol Then
                            QuoteCheckArray(i, c) = RealSteelQuoteArray(i + 1, j)
                    End If
                Next j
            Next i
            
            'put N in empty cells
            For i = 1 To UBound(QuoteCheckArray, 1)
                For j = 1 To UBound(QuoteCheckArray, 2)
                    If QuoteCheckArray(i, j) = "" Then
                        QuoteCheckArray(i, j) = "N"
                    End If
                Next j
            Next i
        
            
        'close the csv and active the quote checker
        CompBook.Activate
        QuoteWorkbook.Close savechanges:=True
                  
        Dim colr As Integer
        colr = QuoteCheck.ListObjects("BOMCopy").ListColumns.count
        'resize the quote table so that the data can be pasted in the table
        Quote.Resize Quote.range.Resize(QuoteNumRows, colr)
        
        'paste the quote array in desired order in the quote check sheet
        With QuoteCheck
            'part number
            .ListObjects("Quote").ListColumns(1).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 1)
            'section
            .ListObjects("Quote").ListColumns(2).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 3)
            'grade
            .ListObjects("Quote").ListColumns(3).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 2)
            'process
            .ListObjects("Quote").ListColumns(4).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 5)
            'quantity
            .ListObjects("Quote").ListColumns(5).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 4)
            'Press
            .ListObjects("Quote").ListColumns(6).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 6)
            'Drill/Tap
            .ListObjects("Quote").ListColumns(7).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 7)
            'Machined
            .ListObjects("Quote").ListColumns(8).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 8)
            'Bevel
            .ListObjects("Quote").ListColumns(9).DataBodyRange.Resize = Application.Index(QuoteCheckArray, , 9)
        End With
End Sub

Sub CheckQuote()
       
     Dim CompBook As Workbook: Set CompBook = ThisWorkbook
        
    'Declear worksheets
    Dim QuoteCheck As Worksheet
    Dim MBOM As Worksheet
    Set QuoteCheck = CompBook.Worksheets("Quote Check")
    Set MBOM = CompBook.Worksheets("Machine BOM")
       
    Dim BOMCopy As ListObject
    Dim Quote As ListObject
    Set BOMCopy = QuoteCheck.ListObjects("BOMCopy")
    Set Quote = QuoteCheck.ListObjects("Quote")
    
    'if there is no data to be check create popup msg and exit sub
    If BOMCopy.DataBodyRange Is Nothing Or Quote.DataBodyRange Is Nothing Then
        MsgBox "Please import data to be checked."
        Exit Sub
    Else
    
    'declear array to store errors found in quote
    Dim NotesArray() As Variant
    Dim note As Integer
    note = 1
    ReDim Preserve NotesArray(1 To note)
    
    'first row of array that can be copied into email
    NotesArray(note) = "Thank you for the quote. Can you please make the following corrections:"
    note = note + 1
    
    'declear loop variables
    Dim partnum As String
    Dim comprow As Integer
    Dim q As Integer
    Dim BOMrow As Integer
    Dim c As Integer
    Dim i As Integer
    c = 1
    
    'number of rows to be checked
    BOMrow = QuoteCheck.ListObjects("BOMCopy").DataBodyRange.rows.count
        
    'declear variable for finding matches
    Dim matchfound As Boolean
    matchfound = False
                  
    'loop through the BOM table to get the part number to be checked
    For q = 1 To BOMrow
        
        matchfound = False
        'part number to be checked
        partnum = [BOMCopy].DataBodyRange.Cells(q, 1)
        
        'Loop through quote table to find part number to be checked
        For i = 1 To Quote.ListRows.count
                If partnum = CStr([Quote].DataBodyRange.Cells(i, 1)) Then
                    comprow = c
                    matchfound = True
                    Exit For
                End If
                c = c + 1
        Next i
        
        'If part isnt in the quote store it in an array and highlight the row red in BOM
        If matchfound = False Then
                ReDim Preserve NotesArray(1 To note)
                NotesArray(note) = "    -  " & partnum & " is not in quote"
                note = note + 1
                QuoteCheck.ListObjects("BOMCopy").ListRows(q).range.Font.Color = RGB(255, 74, 74)
        Else
        
            'Check if the quote matches the BOM
            For i = 1 To 9
                If InStr(1, CStr([BOMCopy].DataBodyRange.Cells(q, i)), CStr([Quote].DataBodyRange.Cells(comprow, i)), 1) > 0 Then
                    'Change cell colour to green if they match
                    [BOMCopy].DataBodyRange.Cells(q, i).Font.ColorIndex = 4
                    [Quote].DataBodyRange.Cells(comprow, i).Font.ColorIndex = 4
                    
                    'exception to account for PFS vs marshall eg sheet naming
                ElseIf CStr([BOMCopy].DataBodyRange.Cells(q, i)) = "EG SHEET" And CStr([Quote].DataBodyRange.Cells(comprow, i)) = "ElectroGalv" Then
                    [BOMCopy].DataBodyRange.Cells(q, i).Font.ColorIndex = 4
                    [Quote].DataBodyRange.Cells(comprow, i).Font.ColorIndex = 4
                    
                    'exception for 350 and corten (when 350 required but under min thickness)
                ElseIf CStr([BOMCopy].DataBodyRange.Cells(q, i)) = "350" And CStr([Quote].DataBodyRange.Cells(comprow, i)) = "Corten" Then
                    [BOMCopy].DataBodyRange.Cells(q, i).Font.ColorIndex = 4
                    [Quote].DataBodyRange.Cells(comprow, i).Font.ColorIndex = 4
                    
                    'exception for stainless steel naming
               ElseIf CStr([BOMCopy].DataBodyRange.Cells(q, i)) = "304" And CStr([Quote].DataBodyRange.Cells(comprow, i)) = "304 N1" Then
                    [BOMCopy].DataBodyRange.Cells(q, i).Font.ColorIndex = 4
                    [Quote].DataBodyRange.Cells(comprow, i).Font.ColorIndex = 4
                    
                Else
                    'change cell colour to red if they do not match
                    [BOMCopy].DataBodyRange.Cells(q, i).Font.Color = RGB(255, 74, 74)
                    [Quote].DataBodyRange.Cells(comprow, i).Font.Color = RGB(255, 74, 74)
                    
                    'record the difference in a notes array
                    'section mistake catch
                    If i = 2 Then
                        ReDim Preserve NotesArray(1 To note)
                        NotesArray(note) = "    -  " & partnum & " section should be " & [BOMCopy].DataBodyRange.Cells(q, i).Value & "mm"
                        note = note + 1
                    End If
                    'material mistake catch
                    If i = 3 Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " material should be " & [BOMCopy].DataBodyRange.Cells(q, i).Value
                            note = note + 1
                    End If
                    
                    'process mistake catch, ignore if no process given in PFS BOM
                    If i = 4 Then
                        If [BOMCopy].DataBodyRange.Cells(q, i).Value = "" Then
                            Else
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " process should be " & [BOMCopy].DataBodyRange.Cells(q, i).Value
                            note = note + 1
                        End If
                    End If
                    
                    'quantity mistake catch
                    If i = 5 Then
                        ReDim Preserve NotesArray(1 To note)
                        NotesArray(note) = "    -  " & partnum & " quantity should be " & [BOMCopy].DataBodyRange.Cells(q, i).Value
                        note = note + 1
                    End If
                                        
                    'press mistake catch
                    If i = 6 Then
                        If [BOMCopy].DataBodyRange.Cells(q, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " requires pressing"
                            note = note + 1
                        End If
                        If [Quote].DataBodyRange.Cells(comprow, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " does not require pressing"
                            note = note + 1
                        End If
                    End If
                                        
                    'Drill/tap mistake catch
                    If i = 7 Then
                        If [BOMCopy].DataBodyRange.Cells(q, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " requires drilling/tapping"
                            note = note + 1
                        End If
                        If [Quote].DataBodyRange.Cells(comprow, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " does not require drilling/tapping"
                            note = note + 1
                        End If
                    End If
                                      
                    'Machine mistake catch
                    If i = 8 Then
                        If [BOMCopy].DataBodyRange.Cells(q, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " requires machining"
                            note = note + 1
                        End If
                        If [Quote].DataBodyRange.Cells(comprow, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " does not require machining"
                            note = note + 1
                        End If
                    End If
                    
                    'Bevel mistake catch
                    If i = 9 Then
                        If [BOMCopy].DataBodyRange.Cells(q, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " requires beveling"
                            note = note + 1
                        End If
                        If [Quote].DataBodyRange.Cells(comprow, i) = "Y" Then
                            ReDim Preserve NotesArray(1 To note)
                            NotesArray(note) = "    -  " & partnum & " does not require beveling"
                            note = note + 1
                        End If
                    End If
                End If
            Next i
        End If
            c = 1
    Next q
    
    'Check if there is a part in the quote that is not in the bom
    'number of rows to be checked
    Dim Quoterow As Integer
    Quoterow = QuoteCheck.ListObjects("Quote").DataBodyRange.rows.count
    matchfound = False
                  
    'loop through the BOM table to get the part number to be checked
    For q = 1 To Quoterow
        
        matchfound = False
        'part number to be checked
        partnum = [Quote].DataBodyRange.Cells(q, 1)
        
        'Loop through quote table to find part number to be checked
        For i = 1 To Quote.ListRows.count
                If partnum = CStr([BOMCopy].DataBodyRange.Cells(i, 1)) Then
                    matchfound = True
                    Exit For
                End If
         Next i
        
        'If part isnt in the quote store it in an array and highlight the row red in BOM
        If matchfound = False Then
                ReDim Preserve NotesArray(1 To note)
                NotesArray(note) = "    -  Remove " & partnum & " from quote"
                note = note + 1
                QuoteCheck.ListObjects("Quote").ListRows(q).range.Font.Color = RGB(255, 74, 74)
        End If
    
    Next q
    

    'declear shape to put output notes in for copying
    Dim output As Shape
    Set output = QuoteCheck.Shapes("output")
    With output.TextFrame.Characters
        .Font.Name = "Source Code Pro"
        .Font.Size = "8"
        .text = Join(NotesArray, vbCr)
    End With
End If
    
End Sub

'cleans section and material for comparison with quote
Sub SectionClean()
       
        Dim BOMImportArray() As Variant
        Dim BOM As ListObject
        Set BOM = QuoteComp.ListObjects("BOMCopy")
                
        'Import bom to an array
        BOMImportArray() = BOM.range
        
        'get number of rows
        Dim rows As Integer
        rows = UBound(BOMImportArray()) - 1
        
        'declear column variables
        Dim section As Variant
        Dim material As Variant
        Dim j As Integer
        Set section = BOM.ListColumns(2).DataBodyRange
        Set material = BOM.ListColumns(3).DataBodyRange
        Dim i As Integer
        Dim item As Variant
        
        'define section regex
        Dim regexSection As New RegExp
        regexSection.Global = True
        regexSection.Pattern = "((\d{1,2}\.?\d{1})|\d{1})"
        Dim text As Variant
        i = 1
        
        'loop through and clean section column
        For Each text In section
            Dim mc As MatchCollection
            Set mc = regexSection.Execute(text)
            For Each item In mc
                section(i) = item
                i = i + 1
           Next item
        Next text
                
        'define material regex
        Dim regexMaterial As New RegExp
        regexMaterial.Global = True
        regexMaterial.Pattern = "(\d{3,4}|EG SHEET)"
        i = 1
        
        'loop through and clean material column
        For Each text In material
            Dim mc2 As MatchCollection
            Set mc2 = regexMaterial.Execute(text)
            For Each item In mc2
                material(i) = item
                i = i + 1
            Next item
        Next text
        
        'clean real steel material column so that it can be compared to pfs BOM
        If RealSteel = True Then
            Dim Quote As ListObject
            Set Quote = QuoteComp.ListObjects("Quote")
            Dim material2 As Variant
            Set material2 = Quote.ListColumns(3).DataBodyRange
            i = 1
            'loop through and clean material column
            For Each text In material2
                Dim mc3 As MatchCollection
                Set mc3 = regexMaterial.Execute(text)
                For Each item In mc3
                    material2(i) = item
                    i = i + 1
                Next item
            Next text
        End If

End Sub
