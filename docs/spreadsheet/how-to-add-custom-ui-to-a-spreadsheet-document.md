---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fdb29547-c295-4e7d-9fc5-d86d8d8c2967
title: 'How to: Add custom UI to a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/10/2024
ms.localizationpriority: high
---

# Add custom UI to a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically add custom UI, modifying the ribbon, to a Microsoft Excel worksheet. It contains an example **AddCustomUI** method to illustrate
this task.



## Creating Custom UI

Before using the Open XML SDK to create a ribbon customization in an Excel workbook, you must first create the customization content. Describing the XML required to create a ribbon customization is beyond the scope of this topic. In addition, you will find it far easier to use the Ribbon Designer in Visual Studio to create the customization for you. For more information about customizing the ribbon by using the Visual Studio Ribbon Designer, see [Ribbon Designer](https://msdn.microsoft.com/library/26617206-f4da-416f-a18a-d817b2d4872d(Office.15).aspx) and [Walkthrough: Creating a Custom Tab by Using the Ribbon Designer](https://msdn.microsoft.com/library/312865e6-950f-46ab-88de-fe7eb8036bfe(Office.15).aspx).
For the purposes of this demonstration, you will need an XML file that contains a customization, and the following code provides a simple customization (or you can create your own by using the Visual Studio Ribbon Designer, and then right-click to export the customization to an XML file). The samples below are the xml strings used in this example. This XML content describes a ribbon customization that includes a button labeled "Click Me!" in a group named Group1 on the **Add-Ins** tab in Excel. When you click the button, it attempts to run a macro named **SampleMacro** in the host workbook.

### [C#](#tab/cs-xml)
[!code-csharp[](../../samples/spreadsheet/add_custom_ui/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-xml)
[!code-vb[](../../samples/spreadsheet/add_custom_ui/vb/Program.vb#snippet4)]
***

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

## Interact with the Workbook

The sample method, **AddCustomUI**, starts by opening the requested workbook in read/write mode, as shown in the following code.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/add_custom_ui/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/add_custom_ui/vb/Program.vb#snippet1)]
***


## Work with the Ribbon Extensibility Part

Next, as shown in the following code, the sample method attempts to retrieve a reference to the single ribbon extensibility part. If the part does not yet exist, the code creates it and stores a reference to the new part.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/add_custom_ui/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/add_custom_ui/vb/Program.vb#snippet2)]
***


## Add the Customization

Given a reference to the ribbon extensibility part, the following code finishes by setting the part's **CustomUI** property to a new [CustomUI](https://msdn.microsoft.com/library/office/documentformat.openxml.office.customui.customui.aspx) object that contains the supplied customization. Once the customization is in place, the code saves the custom UI.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/add_custom_ui/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/add_custom_ui/vb/Program.vb#snippet3)]
***


## Sample Code

The following is the complete **AddCustomUI** code sample in C\# and Visual Basic. The first argument passed to the **AddCustomUI** should be the absolute
path to the AddCustomUI.xlsm file created from the instructions above.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/add_custom_ui/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/add_custom_ui/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

