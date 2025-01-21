using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;

// <Snippet4>
string xml =
@"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
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
// </Snippet4>
;

// args[0] should be the absolute path to the AddCustomUI.xlsm created earlier in the tutorial.
AddCustomUI(args[0], xml);

// <Snippet0>
static void AddCustomUI(string fileName, string customUIContent)
{
    // <Snippet1>
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
    // </Snippet1>
    {
        // <Snippet2>
        // You can have only a single ribbon extensibility part.
        // If the part doesn't exist, create it.
        RibbonExtensibilityPart part = document.RibbonExtensibilityPart ?? document.AddRibbonExtensibilityPart();
        // </Snippet2>

        // <Snippet3>
        part.CustomUI = new CustomUI(customUIContent);
        // </Snippet3>
    }
}
// </Snippet0>

