#nullable disable

using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;

static void AddCustomUI(string fileName, string customUIContent)
{
    // Add a custom UI part to the document.
    // Use this sample XML to test:
    //<customUI xmlns="https://schemas.microsoft.com/office/2006/01/customui">
    //    <ribbon>
    //        <tabs>
    //            <tab idMso="TabAddIns">
    //                <group id="Group1" label="Group1">
    //                    <button id="Button1" label="Button1" 
    //                    showImage="false" onAction="SampleMacro"/>
    //                </group>
    //            </tab>
    //        </tabs>
    //    </ribbon>
    //</customUI>

    // In the sample XLSM file, create a module and create a procedure
    // named SampleMacro, using this 
    // signature: Public Sub SampleMacro(control As IRibbonControl)
    // Add some code, and then save and close the XLSM file. Run this
    // example to add a button to the Add-Ins tab that calls the macro,
    // given the XML content above in the AddCustomUI.xml file.

    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
    {
        // You can have only a single ribbon extensibility part.
        // If the part doesn't exist, create it.
        var part = document.RibbonExtensibilityPart;
        if (part == null)
        {
            part = document.AddRibbonExtensibilityPart();
        }
        part.CustomUI = new CustomUI(customUIContent);
        part.CustomUI.Save();
    }
}