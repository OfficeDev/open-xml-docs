---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 403abe97-7ab2-40ba-92c0-d6312a6d10c8
title: 'How to: Add a comment to a slide in a presentation (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Add a comment to a slide in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to add a comment to the first slide in a presentation
programmatically.


[!include[Structure](./includes/presentation/structure.md)]

## The Structure of the Comment Element

A comment is a text note attached to a slide, with the primary purpose
of enabling readers of a presentation to provide feedback to the
presentation author. Each comment contains an unformatted text string
and information about its author, and is attached to a particular
location on a slide. Comments can be visible while editing the
presentation, but do not appear when a slide show is given. The
displaying application decides when to display comments and determines
their visual appearance.

The following XML element specifies a single comment attached to a
slide. It contains the text of the comment (**text**), its position on the slide (**pos**), and attributes referring to its author
(**authorId**), date and time (**dt**), and comment index (**idx**).

```xml
    <p:cm authorId="0" dt="2006-08-28T17:26:44.129" idx="1">
        <p:pos x="10" y="10"/>
        <p:text>Add diagram to clarify.</p:text>
    </p:cm>
```

The following table contains the definitions of the members and
attributes of the **cm** (comment) element.

| Member/Attribute | Definition |
|---|---|
| authorId | Refers to the ID of an author in the comment author list for the document. |
| dt | The date and time this comment was last modified. |
| idx | An identifier for this comment that is unique within a list of all comments by this author in this document. An author's first comment in a document has index 1. |
| pos | The positioning information for the placement of a comment on a slide surface. |
| text | Comment text. |
| extLst | Specifies the extension list with modification ability within which all future extensions of element type ext are defined. The extension list along with corresponding future extensions is used to extend the storage capabilities of the PresentationML framework. This allows for various new kinds of data to be stored natively within the framework. |


The following XML schema code example defines the members of the **cm** element in addition to the required and
optional attributes.

```xml
    <complexType name="CT_Comment">
       <sequence>
           <element name="pos" type="a:CT_Point2D" minOccurs="1" maxOccurs="1"/>
           <element name="text" type="xsd:string" minOccurs="1" maxOccurs="1"/>
           <element name="extLst" type="CT_ExtensionListModify" minOccurs="0" maxOccurs="1"/>
       </sequence>
       <attribute name="authorId" type="xsd:unsignedInt" use="required"/>
       <attribute name="dt" type="xsd:dateTime" use="optional"/>
       <attribute name="idx" type="ST_Index" use="required"/>
    </complexType>
```

## Sample Code

The following code example shows how to add comments to a
presentation document. To run the program, you can pass in the arguments:

```dotnetcli
dotnet run -- [filePath] [initials] [name] [test ...]
```

> [!NOTE]
> To get the exact author name and initials, open the presentation file and click the **File** menu item, and then click **Options**. The **PowerPointOptions** window opens and the content of the **General** tab is displayed. The author name and initials must match the **User name** and **Initials** in this tab.

### [CSharp](#tab/cs)
[!code-csharp[](../samples/presentation/add_comment/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/add_comment/vb/Program.vb)]

## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

