---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: dd42a9a3-5c16-4cab-ad6d-506cf822ec7a
title: Introduction to markup compatibility (Open XML SDK)
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---

# Introduction to markup compatibility (Open XML SDK)

This topic introduces the markup compatibility features included in the Open XML SDK 2.5 for Office.

## Introduction

Suppose you have a Microsoft Word 2013 document that employs a feature introduced in Microsoft Office 2013. When you open that document in Microsoft Word 2010, an earlier version, what should happen? Ideally, you want the document to remain interoperable with Word 2010, even though Word 2010 will not understand the new feature.

Consider also what should happen if you open that document in a hypothetical later version of Office. Here too, you want the document to work as expected. That is, you want the later version of Office to understand and support a feature employed in a document produced by Word 2013.

Open XML anticipates these scenarios. The Office Open XML File Formats specification describes facilities for achieving the above desired outcomes in [ECMA-376, Second Edition, Part 3 - Markup Compatibility and Extensibility](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/).

The Open XML SDK 2.5 supports markup compatibility in a way that makes it easy for you to achieve the above desired outcomes for and Office 2013 without having to necessarily become an expert in the specification details.

## What is markup compatibility?

Open XML defines formats for word-processing, spreadsheet and presentation documents in the form of specific markup languages, namely WordprocessingML, SpreadsheetML, and PresentationML. With respect to the Open XML file formats, markup compatibility is the ability for a document expressed in one of the above markup languages to facilitate interoperability between applications, or versions of an application, with different feature sets. This is supported through the use of a defined set of XML elements and attributes in the Markup Compatibility namespace of the Open XML specification. Notice that while the markup is supported in the document format, markup producers and consumers, such as Microsoft Word, must support it as well. In other words, interoperability is a function of support both in the file format and by applications.

## Markup compatibility in the Open XML file formats specification

Markup compatibility is discussed in [ECMA-376, Second Edition, Part 3 - Markup Compatibility and Extensibility](https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Second%20Edition, 20Part%203%20-%20Markup%20Compatibility%20and%20Extensibility.zip), which is recommended reading to understand markup compatibility. The specification defines XML attributes to express compatibility rules, and XML elements to specify alternate content. For example, the **Ignorable** attribute specifies namespaces that can be ignored when they are not understood by the consuming application. Alternate-Content elements specify markup alternatives that can be chosen by an application at run time. For example, Word 2013 can choose only the markup alternative that it recognizes. The complete list of compatibility-rule attributes and alternate-content elements and their details can be found in the specification.

## Open XML SDK 2.5 support for markup compatibility

The work that the Open XML SDK 2.5 does for markup compatibility is detailed and subtle. However, the goal can be summarized as: using settings that you assign when you open a document, preprocess the document to:

1. Filter or remove any elements from namespaces that will not be understood (for example, Office 2013 document opened in Office 2010 context)
2. Process any markup compatibility elements and attributes as specified in the Open XML specification.

The preprocessing performed is in accordance with ECMA-376, Second Edition: Part 3.13.

The Open XML SDK 2.5 support for markup compatibility comes primarily in the form of two classes and in the manner in which content is preprocessed in accordance with ECMA-376, Second Edition. The two classes are **OpenSettings** and **MarkupCompatibilityProcessSettings**. Use the former to provide settings that apply to SDK behavior overall. Use the latter to supply one part of those settings, specifically those that apply to markup compatibility.

## Set the stage when you open

When you open a document using the Open XML SDK 2.5, you have the option of using an overload with a signature that accepts an instance of the **[OpenSettings](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.opensettings.aspx)** class as a parameter. You use the open settings class to provide certain important settings that govern the behavior of the SDK. One set of settings in particular, stored in the **[MarkupCompatibilityProcessSettings](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.opensettings.markupcompatibilityprocesssettings.aspx)** property, determines how markup compatibility elements and attributes are processed. You set the property to an instance of the **[MarkupCompatibilityProcessSettings](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.markupcompatibilityprocesssettings.aspx)** class prior to opening a document.

The class has the following properties:

- **[ProcessMode](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.markupcompatibilityprocesssettings.processmode.aspx)** - Determines the parts that are preprocessed.

- **[TargetFileFormatVersions](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.markupcompatibilityprocesssettings.targetfileformatversions.aspx)** - Specifies the context that applies to preprocessing.

By default, documents are not preprocessed. If however you do specify open settings and provide markup compatibility process settings, then the document is preprocessed in accordance with those settings.

The following code example demonstrates how to call the Open method with an instance of the open settings class as a parameter. Notice that the **ProcessMode** and **TargetFileFormatVersions** properties are initialized as part of the **MarkupCompatiblityProcessSettings** constructor.

```csharp
    // Create instance of OpenSettings
    OpenSettings openSettings = new OpenSettings();

    // Add the MarkupCompatibilityProcessSettings
    openSettings.MarkupCompatibilityProcessSettings =
        new MarkupCompatibilityProcessSettings(
            MarkupCompatibilityProcessMode.ProcessAllParts, 
            FileFormatVersions.Office2007);

    // Open the document with OpenSettings
    using (WordprocessingDocument wordDocument = 
        WordprocessingDocument.Open(filename, 
            true,
            openSettings))
    {
        // ... more code here
    }
```

## What happens during preprocessing

During preprocessing, the Open XML SDK 2.5 removes elements and attributes in the markup compatibility namespace, removing the contents of unselected alternate-content elements, and interpreting compatibility-rule attributes as appropriate. This work is guided by the process mode and target file format versions properties.

The **ProcessMode** property determines the parts to be preprocessed. The content in *those* parts is filtered to contain only elements that are understood by the application version indicated in the **TargetFileFormatVersions** property.

> [!WARNING]
> Preprocessing affects what gets saved. When you save a file, the only markup that is saved is that which remains after preprocessing.

## Understand process mode

The process mode specifies which document parts should be preprocessed. You set this property to a member of the **[MarkupCompatibilityProcessMode](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.markupcompatibilityprocessmode.aspx)** enumeration. The default value, **NoProcess**, indicates that no preprocessing is performed. Your application must be able to understand and handle any elements and attributes present in the document markup, including any of the elements and attributes in the Markup Compatibility namespace.

You might want to work on specific document parts while leaving the rest untouched. For example, you might want to ensure minimal modification to the file. In that case, specify **ProcessLoadedPartsOnly** for the process mode. With this setting, preprocessing and the associated filtering is only applied to the loaded document parts, not the entire document.

Finally, there is **ProcessAllParts**, which specifies what the name implies. When you choose this value, the entire document is preprocessed.

## Set the target file format version

The target file format versions property lets you choose to process markup compatibility content in either Office 2010 or Office 2013 context. Set the **TargetFileFormatVersions** property to a member of the **[FileFormatVersions](https://msdn.microsoft.com/library/office/documentformat.openxml.fileformatversions.aspx)** enumeration.

The default value, **Office2010**, means the SDK will assume that namespaces defined in Office 2010 are understood, but not namespaces defined in Office 2013. Thus, during preprocessing, the SDK will ignore the namespaces defined in Office 2013 and choose the Office 2010 compatible alternate-content.

When you set the target file format versions property to **Office2013**, the Open XML SDK 2.5 assumes that all of the namespaces defined in Office 2010 and Office 2013 are understood, does not ignore any content defined under Office 2013, and will choose
the Office 2013 compatible alternate-content.
