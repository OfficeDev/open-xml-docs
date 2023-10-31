---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fdb29547-c295-4e7d-9fc5-d86d8d8c2967
title: 'How to: Add custom UI to a spreadsheet document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---

# Add custom UI to a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for Office to programmatically add custom UI, modifying the ribbon, to an Microsoft Excel 2010 or Microsoft Excel 2013 worksheet. It contains an example **AddCustomUI** method to illustrate
this task.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). Explicitly reference the following assemblies in your project:

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using** directives or **Imports** statements to compile the code in this topic.

```csharp
    using DocumentFormat.OpenXml.Office.CustomUI;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports DocumentFormat.OpenXml.Office.CustomUI
    Imports DocumentFormat.OpenXml.Packaging
```

## Creating Custom UI

Before using the Open XML SDK to create a ribbon customization in an Excel workbook, you must first create the customization content. Describing the XML required to create a ribbon customization is beyond the scope of this topic. In addition, you will find it far easier to use the Ribbon Designer in Visual Studio 2010 to create the customization for you. For more information about customizing the ribbon by using the Visual Studio Ribbon Designer, see [Ribbon Designer](https://msdn.microsoft.com/library/26617206-f4da-416f-a18a-d817b2d4872d(Office.15).aspx) and [Walkthrough: Creating a Custom Tab by Using the Ribbon Designer](https://msdn.microsoft.com/library/312865e6-950f-46ab-88de-fe7eb8036bfe(Office.15).aspx).
For the purposes of this demonstration, you will need an XML file that contains a customization, and the following code provides a simple customization (or you can create your own by using the Visual Studio Ribbon Designer, and then right-click to export the customization to an XML file). Copy the following content into a text file that is named AddCustomUI.xml for use as part of this example. This XML content describes a ribbon customization that includes a button labeled "Click Me!" in a group named Group1 on the **Add-Ins** tab in Excel. When you click the button, it attempts to run a macro named **SampleMacro** in the host workbook.

```xml
    <customUI xmlns="https://schemas.microsoft.com/office/2006/01/customui">
        <ribbon>
            <tabs>
                <tab idMso="TabAddIns">
                    <group id="Group1" label="Group1">
                        <button id="Button1" label="Click Me!" showImage="false" onAction="SampleMacro"/>
                    </group>
                </tab>
            </tabs>
        </ribbon>
    </customUI>
```

## Create the Macro

For this demonstration, the ribbon customization includes a button that attempts to run a macro in the host workbook. To complete the demonstration, you must create a macro in a sample workbook for the button's Click action to call.

1. Create a new workbook.

2. Press Alt+F11 to open the Visual Basic Editor.

3. On the **Insert** tab, click **Module** to create a new module.

4. Add code such as the following to the new module.

    ```vb
        Sub SampleMacro(button As IRibbonControl)
            MsgBox "You Clicked?"
        End Sub
    ```

5. Save the workbook as an Excel Macro-Enabled Workbook named AddCustomUI.xlsm.

## AddCustomUI Method

The **AddCustomUI** method accepts two parameters:

- *filename* — A string that contains a file name that specifies the workbook to modify.

- customUIContent*—A string that contains the custom content (that is, the XML markup that describes the customization).

The following code shows the two parameters.

```csharp
    static public void AddCustomUI(string fileName, string customUIContent)
```

```vb
    Public Sub XLAddCustomUI(ByVal fileName As String,
                                 ByVal customUIContent As String)
```

## Call the AddCustomUI Method

The method modifies the ribbon in an Excel workbook. To call the method, pass the file name of the workbook to modify, and a string that contains the customization XML, as shown in the following example code.

```csharp
    const string SAMPLEXML = "AddCustomUI.xml";
    const string DEMOFILE = "AddCustomUI.xlsm";

    string content = System.IO.File.OpenText(SAMPLEXML).ReadToEnd();
    AddCustomUI(DEMOFILE, content);
```

```vb
    Const SAMPLEXML As String = "AddCustomUI.xml"
    Const DEMOFILE As String = "AddCustomUI.xlsm"

    Dim content As String = System.IO.File.OpenText(SAMPLEXML).ReadToEnd()
    AddCustomUI(DEMOFILE, content)
```

## Interact with the Workbook

The sample method, **AddCustomUI**, starts by opening the requested workbook in read/write mode, as shown in the following code.

```csharp
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, true))
```

```vb
    Using document As SpreadsheetDocument =
        SpreadsheetDocument.Open(fileName, True)
```

## Work with the Ribbon Extensibility Part

Next, as shown in the following code, the sample method attempts to retrieve a reference to the single ribbon extensibility part. If the part does not yet exist, the code creates it and stores a reference to the new part.

```csharp
    // You can have only a single ribbon extensibility part.
    // If the part doesn't exist, create it.
    var part = document.RibbonExtensibilityPart;
    if (part == null)
    {
        part = document.AddRibbonExtensibilityPart();
    }
```

```vb
    ' You can have only a single ribbon extensibility part.
    ' If the part doesn't exist, add it.
    Dim part = document.RibbonExtensibilityPart
    If part Is Nothing Then
        part = document.AddRibbonExtensibilityPart
    End If
```

## Add the Customization

Given a reference to the ribbon extensibility part, the following code finishes by setting the part's **CustomUI** property to a new [CustomUI](https://msdn.microsoft.com/library/office/documentformat.openxml.office.customui.customui.aspx) object that contains the supplied customization. Once the customization is in place, the code saves the custom UI.

```csharp
    part.CustomUI = new CustomUI(customUIContent);
    part.CustomUI.Save();
```

```vb
    part.CustomUI = New CustomUI(customUIContent)
    part.CustomUI.Save()
```

## Sample Code

The following is the complete **AddCustomUI** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/spreadsheet/add_custom_ui/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/add_custom_ui/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
- [Ribbon Designer](https://msdn.microsoft.com/library/26617206-f4da-416f-a18a-d817b2d4872d(Office.15).aspx)
- [Walkthrough: Creating a Custom Tab by Using the Ribbon Designer](https://msdn.microsoft.com/library/312865e6-950f-46ab-88de-fe7eb8036bfe(Office.15).aspx)
