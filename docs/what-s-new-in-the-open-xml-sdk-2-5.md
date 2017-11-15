---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 4fbda0e3-5676-4a8f-ba62-3fba59fa418b
title: What's new in the Open XML SDK 2.5 for Office
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# What's new in the Open XML SDK 2.5 for Office

This topic describes the new and improved features included in the Open
XML SDK 2.5 for Office in addition to known issues and limitations.


---------------------------------------------------------------------------------

-   [Introduction](#BKMK_Introduction)

-   [System Requirements](#BKMK_Requirements)

-   [New Software Requirements](#BKMK_OpenXMLSDK2.0ImprovedArchitecture)

-   [Support of Office 2013 Preview File
    Format](#BKMK_OpenXMLProductivityToolforMicrosoftOffice)

-   [Reads ISO Strict Document Files](#BKMK_MarkupCompatibility)

-   [Deprecated API Information](#BKMK_DocumentValidation)

-   [Updated API information](#BKMK_StreamReadingandWriting)


--------------------------------------------------------------------------------

The Open XML SDK 2.5 is a collection of classes that let you create and
manipulate Open XML documents - documents that adhere to the [Office
Open XML File Formats
Standard](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463).
Because the SDK provides an application program interface that lets you
manipulate Open XML documents directly, you can do so without the need
for the Office client products themselves in both client and server
operating environments. The SDK is designed to let you build high
performance client-side or server-side solutions that perform complex
operations using only a small amount of program code.

This release of the SDK greatly extends support for the file formats
while adding new features.


--------------------------------------------------------------------------------

The Open XML SDK 2.5 has the following system requirements:

**Supported operating systems:** Windows 8 Preview, Windows 7, Windows
Server 2003 Service Pack 2, Windows Server 2008 R2, Windows Server 2008
Service Pack 2, Windows Vista Service Pack 2, Windows XP Service Pack 3

**System Prerequisites:** .NET Framework version 4.0, Up to 300 MB of
available disk space


--------------------------------------------------------------------------------

Open XML SDK 2.5 requires .NET Framework 4.0 or the greater version.
Accordingly, the supported operating systems are updated to be the same
as the requirements of the .NET Framework 4.0.


--------------------------------------------------------------------------------

In addition to offering compatibility with the Open XML SDK 1.0 classes
and the Open XML SDK 2.0 for Microsoft Office classes, Open XML SDK 2.5
provides new classes that enable you to write and build applications to
manipulate Open XML file extensions of the new Microsoft Office 2013
features (see Table 1). By using the Open XML SDK 2.5 Productivity Tool
for Office, those new extensions can be browsed inside the Open XML SDK
documentation in the left pane.

**Table 1. DocumentFormat.OpenXml.Office15 classes**

<table>
<colgroup>
<col width="50%" />
<col width="50%" />
</colgroup>
<thead>
<tr class="header">
<th align="left"><p>Class</p></th>
<th align="left"><p>Description</p></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td align="left"><p>**DocumentFormat.OpenXml.Office15.Excel**</p></td>
<td align="left"><p>Supports new PivotTable features, timeline, and the other new features of Excel</p></td>
</tr>
<tr class="even">
<td align="left"><p>**DocumentFormat.OpenXml.Office15.Word**</p></td>
<td align="left"><p>Supports new Comment features (e.g. Comments pane) and other new features of Word. For example, the **CommentEx</span> class reads the comments author; The <span class="keyword">WebVideoProperty** property is used to insert a video in a Word document</p></td>
</tr>
<tr class="odd">
<td align="left"><p>**DocumentFormat.OpenXml.Office15.PowerPoint, Theme**</p></td>
<td align="left"><p>Supports comment hint, theme family, and the other new features of PowerPoint</p></td>
</tr>
<tr class="even">
<td align="left"><p>**DocumentFormat.OpenXml.Office15.Drawing**</p></td>
<td align="left"><p>Supports new Charts, PivotCharts, and other new Drawing and Chart features</p></td>
</tr>
<tr class="odd">
<td align="left"><p>**DocumentFormat.OpenXml.Office15.WebExtension, WebExtentionPane**</p></td>
<td align="left"><p>Supports app for Office and Task Pane app for Office features. The classes are viable for inserting or modifying app for Office into Word and Excel document files</p></td>
</tr>
</tbody>
</table>

For code samples demonstrating how to use these new classes, please
refer to new articles posted to [Open XML Format SDK
Forum](http://social.msdn.microsoft.com/Forums/en-US/oxmlsdk/threads) in
the Microsoft Developer Network.


---------------------------------------------------------------------------------

The Open XML SDK 2.5 can read ISO/IEC 29500 Strict Format files. Its
document contents are recognized as an Open XML Transitional Format file
when the document is opened. When the file is saved, the document is
saved as an Open XML Transitional Format file.

The Open XML SDK 2.5 converts ISO Strict files to Transitional Formatted
files when any changes are made to the document or when the document is
saved. Unless the document is saved or modified, the document is left as
an ISO Strict Format file.


--------------------------------------------------------------------------------

Because the file format extension of Office 2013 extends members of the
**\<extLst\>** element which did not have any
member elements, Open XML SDK 2.0 classes associated with the empty
**\<extLst\>** (e.g. <span
class="keyword">DocumentFormat.OpenXml.Spreadsheet.PivotFilter:ExtensionList</span>)
are updated to the new variants of <span
class="keyword">ExtensionList</span> classes of Open XML SDK 2.5 (e.g.
<span
class="keyword">DocumentFormat.OpenXml.Spreadsheet.PivotFilter:PivotFilterExtensionList</span>).
The following empty **ExtensionList** in each
class are replaced with a new **ExtensionList**
class including new **Off Open XML** child
element members.

**ExtensionList:**

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.ConnectionShapeLocks</span>

-   **DocumentFormat.OpenXml.Drawing.Theme**

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.ChartDrawing.NonVisualGroupShapeDrawingProperties</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Charts.MultiLevelStringReference</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Charts.NumberReference</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Charts.StringReference</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Charts.SurfaceChartSeries</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Diagrams.DataModelRoot</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGroupShapeDrawingProperties</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Office.Drawing.NonVisualGroupDrawingShapeProperties</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Office2010.Excel.SlicerCacheDefinition</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Office2010.Word.DrawingGroup.NonVisualGroupDrawingShapeProperties</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.CommentAuthor</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Spreadsheet.PivotFilter</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Spreadsheet.QueryTable</span>

**ExtensionListWithModification:**

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.Comment</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.HandoutMaster</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.NotesMaster</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.SlideLayout</span>

-   <span
    class="keyword">DocumentFormat.OpenXml.Presentation.SlideMaster</span>


---------------------------------------------------------------------------------

The following section discusses deprecated API members:

**Smart Tags**

Because *smart tags* were deprecated in Office 2010, the Open XML SDK
2.5 doesn't support smart tag related Open XML elements. The Open XML
SDK 2.5 still can process smart tag elements as *unknown* elements,
however the Open XML SDK 2.5 Productivity Tool for Office validates
those elements (see the following list) in Office document files as
*invalid tags*.

DocumentFormat.OpenXml.Spreadsheet:

-   **SmartTagDisplayValues**

-   **SmartTagProperties**

-   **SmartTags**

-   **SmartTagType**

-   **SmartTagTypes**

DocumentFormat.OpenXml.Wordprocessing:

-   **SaveSmartTagAsXml**

-   **SmartTagAttribut**

-   **SmartTagProperties**

-   **SmartTagRun**

-   **SmartTagType**

**Office 2010 Beta only tags**

The Open XML SDK 2.0 classes for Office 2010 *beta only* Open XML tags
are deprecated. For example, the beta only non-visual properties of
<span
class="keyword">DocumentFormat.OpenXml.Office2010.Drawing.ChartDrawing</span>,
**DocumentFormat.OpenXml.Office2010.Word**, and
**DocumentFormat.OpenXml.Office2010.Drawing**
have been removed from the Open XML SDK 2.5.
