Imports DocumentFormat.OpenXml.CustomProperties
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.VariantTypes
Imports System
Imports System.IO
Imports System.Linq

Module MyModule

    ' <Snippet0>
    ' <Snippet2>
    Function SetCustomProperty(fileName As String, propertyName As String, propertyValue As Object, propertyType As PropertyTypes) As String
        ' </Snippet2>
        ' Given a document name, a property name/value, and the property type, 
        ' add a custom property to a document. The method returns the original
        ' value, if it existed.

        ' <Snippet4>
        Dim returnValue As String = String.Empty

        Dim newProp As New CustomDocumentProperty()
        Dim propSet As Boolean = False

        Dim propertyValueString As String = propertyValue.ToString()
        If propertyValueString Is Nothing Then
            Throw New ArgumentNullException("propertyValue can't be converted to a string.")
        End If

        ' Calculate the correct type.
        Select Case propertyType
            Case PropertyTypes.DateTime
                ' Be sure you were passed a real date, 
                ' and if so, format in the correct way. 
                ' The date/time value passed in should 
                ' represent a UTC date/time.
                If TypeOf propertyValue Is DateTime Then
                    newProp.VTFileTime = New VTFileTime(String.Format("{0:s}Z", Convert.ToDateTime(propertyValue)))
                    propSet = True
                End If

            Case PropertyTypes.NumberInteger
                If TypeOf propertyValue Is Integer Then
                    newProp.VTInt32 = New VTInt32(propertyValueString)
                    propSet = True
                End If

            Case PropertyTypes.NumberDouble
                If TypeOf propertyValue Is Double Then
                    newProp.VTFloat = New VTFloat(propertyValueString)
                    propSet = True
                End If

            Case PropertyTypes.Text
                newProp.VTLPWSTR = New VTLPWSTR(propertyValueString)
                propSet = True

            Case PropertyTypes.YesNo
                If TypeOf propertyValue Is Boolean Then
                    ' Must be lowercase.
                    newProp.VTBool = New VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower())
                    propSet = True
                End If
        End Select

        If Not propSet Then
            ' If the code was not able to convert the 
            ' property to a valid value, throw an exception.
            Throw New InvalidDataException("propertyValue")
        End If
        ' </Snippet4>

        ' <Snippet5>
        ' Now that you have handled the parameters, start
        ' working on the document.
        newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        newProp.Name = propertyName
        ' </Snippet5>

        ' <Snippet6>
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
            ' </Snippet6>

            ' <Snippet7>
            If customProps Is Nothing Then
                ' No custom properties? Add the part, and the
                ' collection of properties now.
                customProps = document.AddCustomFilePropertiesPart()
                customProps.Properties = New Properties()
            End If
            ' </Snippet7>

            ' <Snippet8>
            Dim props = customProps.Properties

            If props IsNot Nothing Then
                ' </Snippet8>

                ' This will trigger an exception if the property's Name 
                ' property is null, but if that happens, the property is damaged, 
                ' and probably should raise an exception.

                ' <Snippet9>
                Dim prop = props.FirstOrDefault(Function(p) CType(p, CustomDocumentProperty).Name.Value = propertyName)

                ' Does the property exist? If so, get the return value, 
                ' and then delete the property.
                If prop IsNot Nothing Then
                    returnValue = prop.InnerText
                    prop.Remove()
                End If
                ' </Snippet9>

                ' <Snippet10>
                ' Append the new property, and 
                ' fix up all the property ID values. 
                ' The PropertyId value must start at 2.
                props.AppendChild(newProp)
                Dim pid As Integer = 2
                For Each item As CustomDocumentProperty In props
                    item.PropertyId = pid
                    pid += 1
                Next
                ' </Snippet10>
            End If
        End Using

        ' <Snippet11>
        Return returnValue
        ' </Snippet11>
    End Function
    ' </Snippet0>

    ' <Snippet3>
    Sub Main(args As String())
        Dim fileName As String = args(0)

        Console.WriteLine(String.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Pedro", PropertyTypes.Text)))

        Console.WriteLine(String.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Shweta", PropertyTypes.Text)))

        Console.WriteLine(String.Join("ReviewDate = ", SetCustomProperty(fileName, "ReviewDate", DateTime.Parse("01/26/2024"), PropertyTypes.DateTime)))
    End Sub
    ' </Snippet3>

    ' <Snippet1>
    Enum PropertyTypes As Integer
        YesNo
        Text
        [DateTime]
        NumberInteger
        NumberDouble
    End Enum
    ' </Snippet1>

End Module
