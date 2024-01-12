' <Snippet0>
Imports DocumentFormat.OpenXml.Office.CustomUI
Imports DocumentFormat.OpenXml.Packaging

Module MyModule

    Sub Main(args As String())
        ' <Snippet4>
        Dim xml As String =
"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
	<ribbon>
		<tabs>
			<tab idMso=""TabAddIns"">
				<group id=""Group1"" label=""Group1"">
					<button id=""Button1"" label=""Click Me!"" showImage=""false"" onAction=""SampleMacro""/>
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>"
        ' </Snippet4>
        ' args(0) contains the path to the AddCustomUI.xlsm file earlier in the tutorial.
        XLAddCustomUI(args(0), xml)
    End Sub


    Public Sub XLAddCustomUI(ByVal fileName As String, ByVal customUIContent As String)

        ' <Snippet1>
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, True)
            ' </Snippet1>

            ' <Snippet2>
            ' You can have only a single ribbon extensibility part.
            ' If the part doesn't exist, add it.
            Dim part = document.RibbonExtensibilityPart
            If part Is Nothing Then
                part = document.AddRibbonExtensibilityPart
            End If
            ' </Snippet2>

            ' <Snippet3>
            part.CustomUI = New CustomUI(customUIContent)
            part.CustomUI.Save()
            ' </Snippet3>

        End Using
    End Sub
End Module
' </Snippet0>