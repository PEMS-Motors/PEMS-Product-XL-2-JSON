Attribute VB_Name = "ProductsXL2JSON"
Option Explicit

Sub ConvertActiveSheetToJSON()
    
    DeleteUnusedRange
    
    'for exporting the file
    Dim filepath As String: filepath = ActiveWorkbook.Path
    Dim filename As String: filename = "Excel2JSON.json"
    
    'for iterating through sheet
    Dim tr As Range: Set tr = ActiveSheet.UsedRange
    Dim r As Range 'row
    Dim c As Range 'cell
    
    'for assembling the string for export
    Dim product As String
    Dim products As String: products = ""
    
    'iterate through sheet
    For Each r In tr.Rows
        If r.Row Mod 500 = 0 Then Debug.Print r.Row & "/" & tr.Rows.count
        If r.Row Mod 341 = 0 Then
            Application.StatusBar = "Processing row " & r.Row & " of " & tr.Rows.count
            DoEvents
        End If
        If r.Row > 1 Then
            product = Row2Product(r)
            If Len(Trim(product)) > 0 Then products = products & vbNewLine & vbTab & product
        End If
    Next r
    products = Left(products, Len(products) - 1) 'removes extra ","
    products = "[" & products & vbNewLine & vbTab & "]"
    
    'export the file
    ExportFile filepath, filename, products
End Sub

Function Row2Product(r As Range) As String: Row2Product = ""
    
    Dim PEMS_Number As String
    Dim Identification As Dictionary
    Dim Specifications As Dictionary
    Dim Dimensions As Dictionary
    Dim Brands As ArrayList
    Dim References As ArrayList
    Dim Replaces As ArrayList
    Dim Fits As ArrayList
    Dim Applications As ArrayList
    Dim Features As Dictionary
    Dim Details As ArrayList
    Dim Certifications As ArrayList
    Dim Warnings As ArrayList
    
    Dim sIdentification As String
    Dim sSpecifications As String
    Dim sDimensions As String
    Dim sBrands As String
    Dim sReferences As String
    Dim sReplaces As String
    Dim sFits As String
    Dim sApplications As String
    Dim sFeatures As String
    Dim sDetails As String
    Dim sCertifications As String
    Dim sWarnings As String
    
    
    Dim c As Range 'cell
    
    Dim hdr As String
    Dim section As String
    Dim attr As String
    
    Dim val As String
    
    Dim list() As String
    
    Dim i As Long
    
    Dim s As String: s = ""
    
    PEMS_Number = ""
    Set Identification = New Dictionary
    Set Specifications = New Dictionary
    Set Dimensions = New Dictionary
    Set Brands = New ArrayList
    Set References = New ArrayList
    Set Replaces = New ArrayList
    Set Fits = New ArrayList
    Set Applications = New ArrayList
    Set Features = New Dictionary
    Set Details = New ArrayList
    Set Certifications = New ArrayList
    Set Warnings = New ArrayList

    
    For Each c In r.Cells
        hdr = c.Parent.Rows(1).Cells(c.Column).Value2 'header
        section = ""
        attr = ""
        If IsError(c.Value) Then c.Value = ""
        val = StandardizeForHTML(c.Value)
        
        If InStr(hdr, ".") <> 0 Then
            section = Left(hdr, InStr(hdr, ".") - 1)
            attr = Right(hdr, Len(hdr) - InStr(hdr, "."))
        Else
            section = hdr
        End If
        
        'Debug.Print "hdr: " & hdr
        'Debug.Print "sect: " & section
        'Debug.Print "attr: " & attr
        'Debug.Print "val: " & val
        
        
        
        Select Case section
            Case "PEMS Number"
                PEMS_Number = val
            Case "Identification"
                If Len(Trim(val)) > 0 And val <> "0" Then Identification.Add attr, val
            Case "Specifications"
                If Len(Trim(val)) > 0 And val <> "0" Then Specifications.Add attr, val
            Case "Dimensions"
                If Len(Trim(val)) > 0 And val <> "0" Then Dimensions.Add attr, val
            Case "Cross Reference"
                Select Case attr
                    Case "Brands"
                        If Len(Trim(val)) > 0 And val <> "0" Then
                            list = Split(val, "|")
                            For i = 0 To UBound(list)
                                Brands.Add list(i)
                            Next i
                        End If
                    Case "References"
                        If Len(Trim(val)) > 0 And val <> "0" Then
                            list = Split(val, "|")
                            For i = 0 To UBound(list)
                                References.Add list(i)
                            Next i
                        End If
                    Case "Replaces"
                        If Len(Trim(val)) > 0 And val <> "0" Then
                            list = Split(val, "|")
                            For i = 0 To UBound(list)
                                Replaces.Add list(i)
                            Next i
                        End If
                    Case "Fits"
                        If Len(Trim(val)) > 0 And val <> "0" Then
                            list = Split(val, "|")
                            For i = 0 To UBound(list)
                                Fits.Add list(i)
                            Next i
                        End If
                    Case Else
                        'invalid entry, do nothing
                End Select
            Case "Applications"
                If Len(Trim(val)) > 0 And val <> "0" Then
                    list = Split(val, "|")
                    For i = 0 To UBound(list)
                        Applications.Add list(i)
                    Next i
                End If
            Case "Features"
                If Len(Trim(val)) > 0 And val <> "0" Then Features.Add attr, val
            Case "Details"
                If Len(Trim(val)) > 0 And val <> "0" Then
                    list = Split(val, "|")
                    For i = 0 To UBound(list)
                        Details.Add list(i)
                    Next i
                End If
            Case "Certifications"
                If Len(Trim(val)) > 0 And val <> "0" Then
                    list = Split(val, "|")
                    For i = 0 To UBound(list)
                        Certifications.Add list(i)
                    Next i
                End If
            Case "Warnings"
                If Len(Trim(val)) > 0 And val <> "0" Then
                    list = Split(val, "|")
                    For i = 0 To UBound(list)
                        Warnings.Add list(i)
                    Next i
                End If
            Case Else
                'invalid entry, do nothing
        End Select
    Next c
        
    sIdentification = Dictionary2JSON(Identification)
    sSpecifications = Dictionary2JSON(Specifications)
    sDimensions = Dictionary2JSON(Dimensions)
    sBrands = ArrayList2JSON(Brands)
    sReferences = ArrayList2JSON(References)
    sReplaces = ArrayList2JSON(Replaces)
    sFits = ArrayList2JSON(Fits)
    sApplications = ArrayList2JSON(Applications)
    sFeatures = Dictionary2JSON(Features)
    sDetails = ArrayList2JSON(Details)
    sCertifications = ArrayList2JSON(Certifications)
    sWarnings = ArrayList2JSON(Warnings)
    
    s = vbTab & vbTab & """" & "PEMS Number" & """" & ":" & """" & PEMS_Number & """" & ","
    s = s & vbNewLine & vbTab & vbTab & """" & "Identification" & """" & ":" & sIdentification
    s = s & vbNewLine & vbTab & vbTab & """" & "Specifications" & """" & ":" & sSpecifications
    s = s & vbNewLine & vbTab & vbTab & """" & "Dimensions" & """" & ":" & sDimensions
    s = s & vbNewLine & vbTab & vbTab & """" & "Cross Reference" & """" & ":{"
    s = s & vbNewLine & vbTab & vbTab & vbTab & """" & "Brands" & """" & ":" & sBrands
    s = s & vbNewLine & vbTab & vbTab & vbTab & """" & "References" & """" & ":" & sReferences
    s = s & vbNewLine & vbTab & vbTab & vbTab & """" & "Replaces" & """" & ":" & sReplaces
    s = s & vbNewLine & vbTab & vbTab & vbTab & """" & "Fits" & """" & ":" & sFits
    If Len(s) > 0 Then s = Left(s, Len(s) - 1) 'Remove extra ","
    s = s & vbNewLine & vbTab & vbTab & "},"
    s = s & vbNewLine & vbTab & vbTab & """" & "Applications" & """" & ":" & sApplications
    s = s & vbNewLine & vbTab & vbTab & """" & "Features" & """" & ":" & sFeatures
    s = s & vbNewLine & vbTab & vbTab & """" & "Details" & """" & ":" & sDetails
    s = s & vbNewLine & vbTab & vbTab & """" & "Certifications" & """" & ":" & sCertifications
    s = s & vbNewLine & vbTab & vbTab & """" & "Warnings" & """" & ":" & sWarnings
    
    If Len(s) > 0 Then s = Left(s, Len(s) - 1) 'Remove extra ","
    s = "{" & vbNewLine & s & vbNewLine & vbTab & "},"
    Row2Product = s
End Function
Function StandardizeForHTML(text As String) As String: StandardizeForHTML = ""
    Dim val As String: val = text
    
    Do While InStr(val, vbNewLine) <> 0
        val = WorksheetFunction.Substitute(val, vbNewLine, "|")
    Loop
    
    Do While InStr(val, Chr(10)) <> 0
        val = WorksheetFunction.Substitute(val, Chr(10), "|")
    Loop
    
    Do While InStr(val, vbTab) <> 0
        val = WorksheetFunction.Substitute(val, vbTab, " ")
    Loop
    
    Do While InStr(val, "  ") <> 0
        val = WorksheetFunction.Substitute(val, "  ", " ")
    Loop
    
    val = WorksheetFunction.Substitute(val, "†", "")
    val = WorksheetFunction.Substitute(val, "®", "")
    val = WorksheetFunction.Substitute(val, "&", " and ")
    val = WorksheetFunction.Substitute(val, "°", "")
    val = WorksheetFunction.Substitute(val, """", "")
    val = WorksheetFunction.Substitute(val, "'", "")
    val = WorksheetFunction.Substitute(val, "`", "")
    val = WorksheetFunction.Substitute(val, "Ø", " phase")
    val = WorksheetFunction.Substitute(val, "&", " and ")
    val = WorksheetFunction.Substitute(val, "*", "")
    
    StandardizeForHTML = val
    
End Function
Function Dictionary2JSON(d As Dictionary) As String: Dictionary2JSON = ""
    Dim key As Variant
    Dim val As String
    For Each key In d.Keys
        val = Trim(d(key))
        If Len(Trim(val)) > 0 And Trim(val) <> "0" Then
            Dictionary2JSON = Dictionary2JSON & """" & key & """" & ":" & """" & val & """" & ","
        End If
    Next key
    If Len(Dictionary2JSON) > 0 Then Dictionary2JSON = Left(Dictionary2JSON, Len(Dictionary2JSON) - 1) 'remove extra ","
    Dictionary2JSON = "{" & Dictionary2JSON & "},"
End Function
Function ArrayList2JSON(l As ArrayList) As String: ArrayList2JSON = ""
    Dim key As Variant
    Dim val As String
    Dim i As Long
    If l.count > 0 Then
        For i = 0 To l.count - 1
            val = Trim(l(i))
            If Len(Trim(val)) > 0 And Trim(val) <> "0" Then
                ArrayList2JSON = ArrayList2JSON & """" & val & """" & ","
            End If
        Next i
    End If
    If Len(ArrayList2JSON) > 0 Then ArrayList2JSON = Left(ArrayList2JSON, Len(ArrayList2JSON) - 1) 'remove extra ","
    ArrayList2JSON = "[" & ArrayList2JSON & "],"
End Function
Sub ExportFile(filepath As String, filename As String, filecontent As String)
    Dim output_file As String: output_file = filepath & "\" & filename
    Open output_file For Output As #1
    Print #1, filecontent
    Close #1
    Application.StatusBar = "Processing Complete; File saved as: " & output_file
    Debug.Print "Processing Complete"
    Debug.Print "Saved As:"
    Debug.Print output_file
End Sub
Sub DeleteUnusedRange()
    Dim i As Long
    Dim r As Range
    Dim startcell As Range
    Dim endcell As Range
    Dim deleteme As Range
    
    'Debug.Print "UsedRange Before: " & ActiveSheet.UsedRange.Address
    
    For i = ActiveSheet.UsedRange.Rows.count To 1 Step -1
        Set r = ActiveSheet.UsedRange.Rows(i)
        If Len(Trim(WorksheetFunction.Concat(r))) > 0 Then
            Set startcell = ActiveSheet.Cells(r.Row, 1).Offset(1, 0)
            Set endcell = ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Cells.count).Offset(1, 0)
            Set deleteme = Range(startcell, endcell)
            'Debug.Print "Deleting: " & deleteme.Address
            deleteme.EntireRow.Delete
            Exit For
        End If
    Next i
    
    Debug.Print "Processing Complete"
    Debug.Print "UsedRange After: " & ActiveSheet.UsedRange.Address
End Sub
