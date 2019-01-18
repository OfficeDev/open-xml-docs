---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 620e86b5-49f2-43dc-85d4-9c7456c09552
title: About the Open XML SDK 2.5 for Office
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Priority
---

# About the Open XML SDK 2.5 for Office

Open XML is an open standard for word-processing documents, presentations, and spreadsheets that can be freely implemented by multiple applications on different platforms. Open XML is designed to faithfully represent existing word-processing documents, presentations, and spreadsheets that are encoded in binary formats defined by Microsoft Office applications. The reason for Open XML is simple: billions of documents now exist but, unfortunately, the information in those documents is tightly coupled with the programs that created them. The purpose of the Open XML standard is to de-couple documents created by Microsoft Office applications so that they can be manipulated by other applications independent of proprietary formats and without the loss of data.


## Structure of an Open XML Package

An Open XML file is stored in a ZIP archive for packaging and compression. You can view the structure of any Open XML file using a ZIP viewer. An Open XML document is built of multiple document parts. The relationships between the parts are themselves stored in document parts. The ZIP format supports random access to each part. For example, an application can move a slide from one presentation to another presentation without parsing the slide content. Likewise, an application can strip all of the comments out of a word processing document without parsing any of its contents.

The document parts in an Open XML package are created as XML markup. Because XML is structured plain text, you can view the contents of a document part using text readers or you can parse the contents using processes such as XPath.

Structurally, an Open XML document is an Open Packaging Conventions (OPC) package. As stated previously, a package is composed of a collection of document parts. Each part has a part name that consists of a sequence of segments or a pathname such as "/word/theme/theme1.xml." The package contains a [Content\_Types].xml part that allows you to determine the content type of all document parts in the package. A set of explicit relationships for a source package or part is contained in a relationships part that ends with the .rels extension.

Word processing documents are described by using WordprocessingML markup. For more information, see [Working with WordprocessingML documents (Open XML SDK)](working-with-wordprocessingml-documents.md). A WordprocessingML document is composed of a collection of stories where each story is one of the following:

-   Main document (the only required story)

-   Glossary document

-   Header and footer

-   Comments

-   Text box

-   Footnote and endnote

Presentations are described by using PresentationML markup. For more information, see [Working with PresentationML documents (Open XML SDK)](working-with-presentationml-documents.md). Presentation packages can contain the following document parts:

-   Slide master

-   Notes master

-   Handout master

-   Slide layout

-   Notes

Spreadsheet workbooks are described by using SpreadsheetML markup. For more information, see [Working with SpreadsheetML documents (Open XML SDK)](working-with-spreadsheetml-documents.md). Workbook packages can contain:

-   Workbook part (required part)

-   One or more worksheets

-   Charts

-   Tables

-   Custom XML


## Open XML SDK 1.0

Version 1 of the Open XML SDK simplified the manipulation of Open XML packages. The Open XML SDK Application Programming Interface (API) encapsulates many of the common tasks that you typically perform on Open XML packages, so you can perform complex operations with just a few lines of code. Some common tasks:

-   **Search**—With a few lines of code, you can search a collection of Excel 2007 worksheets for some arbitrary data.

-   **Document assembly**—You can create documents by combining the document parts of existing documents programmatically. For example, you can pull slides from various PowerPoint 2007 presentations to create a single presentation.

-   **Validation**—With a few lines of code, you can validate the document parts in a package or validate an entire package against a schema.

-   **Data update**—With the Open XML object model, you can easily modify the data in multiple packages.

-   **Privacy**—With a few lines of code, you can remove comments and other personal information from a document before it is distributed.


## Open XML SDK 2.0 for Microsoft Office

The Open XML SDK 2.0 for Microsoft Office extended the strongly typed class support from the part classes, which are provided in version 1.0, to the XML content in each part. All functions available in version 1.0 are still supported. With version 2.0, you are able to program against the XML content inside the part. The SDK supports programming in the style of LINQ to XML which makes coding against the XML content much easier than the traditional W3C XML DOM programming model.

The SDK supports the following common tasks/scenarios:

-   **Strongly Typed Classes and Objects**—Instead of relying on generic XML functionality to manipulate XML, which requires that you be aware of element/attribute/value spelling as well as namespaces, you can use the Open XML SDK to accomplish the same solution simply by manipulating objects that represent elements/attributes/values. All schema types are represented as strongly typed Common Language Runtime (CLR) classes and all attribute values as enumerations.

-   **Content Construction, Search, and Manipulation**—The LINQ technology is built directly into the SDK. As a result, you are able to perform functional constructs and lambda expression queries directly on objects representing Open XML elements. In addition, the SDK allows you to easily traverse and manipulate content by providing support for collections of objects, like tables and paragraphs.

-   **Validation**—The Open XML SDK 2.0 for Microsoft Office provides validation functionality, enabling you to validate Open XML documents against different variations of the Open XML Format.


## Open XML SDK 2.5 for Office

The Open XML SDK 2.5 provides the namespaces and members to support the Microsoft Office 2013. The Open XML SDK 2.5 can also read ISO/IEC 29500 Strict Format files. The Strict format is a subset of the Transitional format that does not include legacy features - this makes it theoretically easier for a new implementer to support since it has a smaller technical footprint.

The SDK supports the following common tasks/scenarios:

-   **Support of Office 2013 Preview file format**—In addition to the Open XML SDK 2.0 for Microsoft Office classes, Open XML SDK 2.5 provides new classes that enable you to write and build applications to manipulate Open XML file extensions of the new Office 2013 features.

-   **Reads ISO Strict Document File**—Open XML SDK 2.5 can read ISO/IEC 29500 Strict Format files. When the Open XML SDK 2.5 API opens a Strict Format file, each Open XML part in the file is loaded to an **OpenXmlPart**  class of the Open XML SDK 2.5 by mapping `http://purl.oclc.org/ooxml/` namespaces to the corresponding `http://schemas.openxmlformats.org/` namespaces.

-   **Fixes to the Open XML SDK 2.0 for Microsoft Office**—Open XML SDK 2.5 includes fixes to known issues in the Open XML SDK 2.0 for Microsoft Office. These include lost whitespaces in PowerPoint presentations and an issue with the Custom UI in Word documents where a specified argument was reported as being out of the range of valid values.

You can find more information about these and other new features of the Open XML SDK 2.5 in the [What's new in the Open XML SDK 2.5 for Office](what-s-new-in-the-open-xml-sdk.md) article.

