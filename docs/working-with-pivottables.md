---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 96697c37-3fa7-4814-85b6-657439435ce1
title: Working with PivotTables (Open XML SDK)
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---

# Working with PivotTables (Open XML SDK)

This topic discusses the Open XML SDK **[PivotTableDefinition](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.pivottabledefinition.aspx)** class and how it relates to the Open XML File Format SpreadsheetML schema. For more information about the overall structure of the parts and elements that make up a SpreadsheetML document, see [Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md).

## PivotTable in SpreadsheetML

The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification introduces the **PivotTableDefinition** (\<**pivotTableDefinition**\>) element.

PivotTables display aggregated views of data easily and in an
understandable layout. Hundreds or thousands of pieces of underlying
information can be aggregated on row & column axes, revealing the
meanings behind the data. PivotTable reports are used to organize and
summarize your data in different ways. Creating a PivotTable report is
about moving pieces of information around to see how they fit together.
In a few gestures the pivot rows and columns can be moved into different
arrangements and layouts.

A PivotTable object has a row axis area, a column axis area, a values
area, and a report filter area. Additionally, PivotTables have a
corresponding field list pane displaying all the fields of data which
can be placed on one of the PivotTable areas.

The workbook points to (and owns the longevity of) the
pivotCacheDefinition part, which in turn points to and owns the
pivotCacheRecords part. The workbook also points to and owns the sheet
part, which in turn points to and owns a pivotTable part definition,
when a PivotTable is on the sheet (there can be multiple PivotTables on
a sheet). The pivotTable part points to the appropriate
pivotCacheDefinition which it is using. Since multiple PivotTables can
use the same cache, the pivotTable part does not own the longevity of
the pivotCacheDefinition.

The pivotTable part describes the particulars of the layout of the
PivotTable on the sheet. It indicates what fields are on the row axis,
the column axis, report filter, and values areas of the PivotTable. It
also indicates formatting information about the PivotTable. If
conditional formatting has been applied to the PivotTable, that is also
expressed in the pivotTable part.

The pivot cache definition contains the definitions of all fields in the
PivotTable. If you create a PivotTable based on a regular table, each
column in the table becomes a field of the pivot cache definition. The
pivot cache contains the field definitions and information about the
type of content found in that field. It also maintains a reference to
the source data in the cache markup so that the pivot cache can be
refreshed along with the cached data in the pivot cache records part.

The data that is displayed in the PivotTable is stored in two locations.
The pivot cache records part maintains the actual data for the
PivotTable. The table cells in the worksheet also store a cached version
of the data, but that is only for display purposes. The pivot cache
records part stores data in one of two ways. The unique values for the
data area of the PivotTable are cached inline. The repeating items that
you normally find on the row and column are referenced. This shared data
is actually stored in the pivot cache definition. Each record in the
pivot cache record part consists of N values where N is equal to the
number of fields defined in the pivot cache definition.

The final step is to create the PivotTable itself. The PivotTable
definition part contains the information about which field is present in
which place of the PivotTable. You can place a field in four areas: row,
column, data or filter. The fields are chosen from the cached fields in
the pivot cache definition.

To create a PivotTable that is ready to use when the workbook is opened
you also need to create the markup for the table cells. The PivotTable
is displayed in the cells of a worksheet and therefore you need to
construct them as well. You can also have the user update the PivotTable
cells when opening the document.

The following table lists the common Open XML SDK classes used when working with the **PivotTableDefinition** class.

| **SpreadsheetML Element** | **Open XML SDK Class**                                                      |
|---------------------------|-------------------------------------------------------------------------------------------------------------------|
|  pivotField            | [PivotField](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.pivotfield.aspx)           |
|  pivotCacheDefinition  | [PivotCacheDefinition](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.pivotcachedefinition.aspx) |
|  pivotCacheRecords     | [PivotCacheRecords](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.pivotcacherecords.aspx)    |

## Open XML SDK PivotTableDefinition Class

The Open XML SDK**PivotTableDefinition**
class represents the PivotTable definition (\<**pivotTableDefinition**\>) element defined in the Open XML File Format schema for SpreadsheetML documents. Use the **PivotTableDefinition** class to manipulate individual \<**pivotTableDefinition**\> elements in a SpreadsheetML document.

The main function of the PivotTable definition is to store information
about which field is displayed on which axis of the PivotTable and in
what order. There are many other settings that can be added to the
PivotTable definition, but the following explains the basics.

The root element names the PivotTable so that it can be used as a data
source. The root element also references the pivot cache by using the ID
added to the workbook part, and defines the caption label to display
above the data area of the PivotTable. All of these elements are
required.

The three main pieces of the **pivotTableDefinition** are: the location of the
table, the display information for the cached fields, and the
positioning information of the cached fields. For more information about
these and other additional elements that make up the **pivotTableDefinition**, see the ISO/IEC 29500
specification.

### PivotField Class

The **PivotTableDefinition** element contains the **PivotField** (\<**pivotField**\>) elements. The following information from the ISO/IEC 29500 specification introduces the **PivotField** (\<**pivotField**\>) element.

Represents a single field in the PivotTable. This element contains information about the field, including the collection of items in the field.

First, define the collection of fields that appear on the PivotTable
using the **pivotFields** element. Each field
serves as a cache for the data of that field in the data source. You do
not need to define the cache. Instead, you can set the **item** element equal to **default** and have the user update the table when
they open the document. The **showAll**
attribute is used to hide certain elements for that data dimension.
After defining which fields are part of the table, the fields are
applied to the four areas of the PivotTable.

### Pivot Cache Definition Class

The following information from the ISO/IEC 29500 specification introduces the **PivotCacheDefinition** (\<**pivotCacheDefinition**\>) element.

The pivotCacheDefinition part defines each field in the
pivotCacheRecords part, including field name and information about the
data contained in the field. The pivotCacheDefinition part also defines
pivot items that are shared among the pivotTable and pivotRecords parts.

The pivot cache defines the source of the data in the PivotTable, which
allows it to be updated, and it defines the list of fields in that data.
Be aware that the cache defines all the fields available to the
PivotTable, not the ones actually used. The PivotTable definition
defines which of the available fields are used by a particular
PivotTable.

The data source definition references the data that is displayed in the
PivotTable. The PivotTable also maintains the data in the cache-records
part to allow the table to be updatable when the data connection is not
available. You cannot rely on the cells of the PivotTable to store the
data because the data in these cells is transient in nature, it changes
when you pivot the table. There are various types of data sources, for
example: worksheets, database, OLAP cube, and other PivotTables.

The last part of the cache definition defines the fields of the data
source using the **cacheField** element. The
**cacheField** element is used for two
purposes: it defines the data type and formatting of the field, and it
is used as a cache for shared strings. The pivot values are stored in
the pivot cache records part. When a recurring string is used as a
value, the cache record uses a reference into the **cacheField** collection of shared items.

### Pivot Cache Records Class

The following information from the ISO/IEC 29500 specification introduces the **PivotCacheRecords** (\<**pivotCacheRecords**\>) element.

The pivotCacheRecords part contains the underlying data to be aggregated. It is a cache of the source data.

The cache records part can store any number of cached records. Each record has the same number of values defined as there are fields in the cache definition. Each record is defined with the **r** element. This record contains value items as child elements. You can provide certain typed values, such as Numeric, Boolean, or Date-Time, or you can reference into the shared items.
