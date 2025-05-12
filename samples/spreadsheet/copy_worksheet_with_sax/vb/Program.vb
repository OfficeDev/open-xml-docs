Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Diagnostics

Module Program
    Sub Main(args As String())
        CopySheetDOM(args(0))
        CopySheetSAX(args(1))
    End Sub

    ' <Snippet0>
    Sub CopySheetDOM(path As String)
        Console.WriteLine("Starting DOM method")

        Dim sw As Stopwatch = New Stopwatch()
        sw.Start()
        ' <Snippet1>
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(path, True)
            ' Get the first sheet
            Dim worksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault()

            If worksheetPart IsNot Nothing Then
                ' </Snippet1>
                ' <Snippet2>
                ' Add a new WorksheetPart
                Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

                ' Make a copy of the original worksheet
                Dim newWorksheet As Worksheet = CType(worksheetPart.Worksheet.Clone(), Worksheet)

                ' Add the new worksheet to the new worksheet part
                newWorksheetPart.Worksheet = newWorksheet
                ' </Snippet2>

                Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()

                If sheets Is Nothing Then
                    spreadsheetDocument.WorkbookPart.Workbook.AddChild(New Sheets())
                End If

                ' <Snippet3>
                ' Find the new WorksheetPart's Id and create a new sheet id
                Dim id As String = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)
                Dim newSheetId As UInteger = CUInt(sheets.ChildElements.Count + 1)

                ' Append a new Sheet with the WorksheetPart's Id and sheet id to the Sheets element
                sheets.AppendChild(New Sheet() With {
                    .Name = "My New Sheet",
                    .SheetId = newSheetId,
                    .Id = id
                })
                ' </Snippet3>
            End If
        End Using

        sw.Stop()

        Console.WriteLine($"DOM method took {sw.Elapsed.TotalSeconds} seconds")
    End Sub
    ' </Snippet0>

    ' <Snippet99>
    Sub CopySheetSAX(path As String)
        Console.WriteLine("Starting SAX method")

        Dim sw As Stopwatch = New Stopwatch()
        sw.Start()
        ' <Snippet4>
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(path, True)
            ' Get the first sheet
            Dim worksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart?.WorksheetParts?.FirstOrDefault()

            If worksheetPart IsNot Nothing Then
                ' </Snippet4>
                ' <Snippet5>
                Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()
                ' </Snippet5>

                ' <Snippet6>
                Using reader As OpenXmlReader = OpenXmlPartReader.Create(worksheetPart)
                    Using writer As OpenXmlWriter = OpenXmlPartWriter.Create(newWorksheetPart)
                        ' </Snippet6>
                        ' <Snippet7>
                        ' Write the XML declaration with the version "1.0".
                        writer.WriteStartDocument()

                        ' Read the elements from the original worksheet part
                        While reader.Read()
                            ' If the ElementType is CellValue it's necessary to explicitly add the inner text of the element
                            ' or the CellValue element will be empty
                            If reader.ElementType Is GetType(CellValue) Then
                                If reader.IsStartElement Then
                                    writer.WriteStartElement(reader)
                                    writer.WriteString(reader.GetText())
                                ElseIf reader.IsEndElement Then
                                    writer.WriteEndElement()
                                End If
                                ' For other elements write the start and end elements
                            Else
                                If reader.IsStartElement Then
                                    writer.WriteStartElement(reader)
                                ElseIf reader.IsEndElement Then
                                    writer.WriteEndElement()
                                End If
                            End If
                        End While
                        ' </Snippet7>
                    End Using
                End Using

                ' <Snippet8>
                Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()

                If sheets Is Nothing Then
                    spreadsheetDocument.WorkbookPart.Workbook.AddChild(New Sheets())
                End If

                Dim id As String = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)
                Dim newSheetId As UInteger = CUInt(sheets.ChildElements.Count + 1)

                sheets.AppendChild(New Sheet() With {
                    .Name = "My New Sheet",
                    .SheetId = newSheetId,
                    .Id = id
                })
                ' </Snippet8>

                sw.Stop()

                Console.WriteLine($"SAX method took {sw.Elapsed.TotalSeconds} seconds")
            End If
        End Using
    End Sub
    ' </Snippet99>
End Module
