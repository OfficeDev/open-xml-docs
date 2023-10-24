Imports System.IO
Imports DocumentFormat.OpenXml.CustomProperties
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.VariantTypes

Module Program
    Sub Main(args As String())
    End Sub



    Public Enum PropertyTypes
        YesNo
        Text
        DateTime
        NumberInteger
        NumberDouble
    End Enum

    Public Function SetCustomProperty( _
        ByVal fileName As String,
        ByVal propertyName As String, _
        ByVal propertyValue As Object,
        ByVal propertyType As PropertyTypes) As String

        ' Given a document name, a property name/value, and the property type, 
        ' add a custom property to a document. The method returns the original 
        ' value, if it existed.

        Dim returnValue As String = Nothing

        Dim newProp As New CustomDocumentProperty
        Dim propSet As Boolean = False

        ' Calculate the correct type:
        Select Case propertyType

            Case PropertyTypes.DateTime
                ' Make sure you were passed a real date, 
                ' and if so, format in the correct way. 
                ' The date/time value passed in should 
                ' represent a UTC date/time.
                If TypeOf (propertyValue) Is DateTime Then
                    newProp.VTFileTime = _
                        New VTFileTime(String.Format("{0:s}Z",
                            Convert.ToDateTime(propertyValue)))
                    propSet = True
                End If

            Case PropertyTypes.NumberInteger
                If TypeOf (propertyValue) Is Integer Then
                    newProp.VTInt32 = New VTInt32(propertyValue.ToString())
                    propSet = True
                End If

            Case PropertyTypes.NumberDouble
                If TypeOf propertyValue Is Double Then
                    newProp.VTFloat = New VTFloat(propertyValue.ToString())
                    propSet = True
                End If

            Case PropertyTypes.Text
                newProp.VTLPWSTR = New VTLPWSTR(propertyValue.ToString())
                propSet = True

            Case PropertyTypes.YesNo
                If TypeOf propertyValue Is Boolean Then
                    ' Must be lowercase.
                    newProp.VTBool = _
                      New VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower())
                    propSet = True
                End If
        End Select

        If Not propSet Then
            ' If the code was not able to convert the 
            ' property to a valid value, throw an exception.
            Throw New InvalidDataException("propertyValue")
        End If

        ' Now that you have handled the parameters, start
        ' working on the document.
        newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
        newProp.Name = propertyName

        Using document = WordprocessingDocument.Open(fileName, True)
            Dim customProps = document.CustomFilePropertiesPart
            If customProps Is Nothing Then
                ' No custom properties? Add the part, and the
                ' collection of properties now.
                customProps = document.AddCustomFilePropertiesPart
                customProps.Properties = New Properties
            End If

            Dim props = customProps.Properties
            If props IsNot Nothing Then
                ' This will trigger an exception is the property's Name property 
                ' is null, but if that happens, the property is damaged, and 
                ' probably should raise an exception.
                Dim prop = props.
                  Where(Function(p) CType(p, CustomDocumentProperty).
                          Name.Value = propertyName).FirstOrDefault()
                ' Does the property exist? If so, get the return value, 
                ' and then delete the property.
                If prop IsNot Nothing Then
                    returnValue = prop.InnerText
                    prop.Remove()
                End If

                ' Append the new property, and 
                ' fix up all the property ID values. 
                ' The PropertyId value must start at 2.
                props.AppendChild(newProp)
                Dim pid As Integer = 2
                For Each item As CustomDocumentProperty In props
                    item.PropertyId = pid
                    pid += 1
                Next
                props.Save()
            End If
        End Using

        Return returnValue

    End Function
End Module