---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 43c49a6d-96b5-4e87-a5bf-01629d61aad4
title: Open XML SDK for Office design considerations
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Open XML SDK for Office design considerations

Before using the Open XML SDK for Office, be aware of the following
design considerations.


--------------------------------------------------------------------------------
## Design Considerations
The Open XML SDK 2.5:

-   Does not replace the Microsoft Office Object Model and provides no
    abstraction on top of the file formats. You must still understand
    the structure of the file formats to use the Open XML SDK 2.5.

-   Does not provide functionality to convert Open XML formats to and
    from other formats, such as HTML or XPS.

-   Does not guarantee document validity of Open XML Formats when you
    use the Open XML SDK or if you decide to manipulate the
    underlying XML directly.

-   Does not provide application behavior such as layout functionality
    in Word or recalculation, data refresh, or adjustment
    functionalities in Excel.
