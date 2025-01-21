## The Structure of the Modern Comment Element

The following XML element specifies a single comment. 
It contains the text of the comment (`t`) and attributes referring to its author
(`authorId`), date time created (`created`), and comment id (`id`).

```xml
<p188:cm id="{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}" authorId="{CD37207E-7903-4ED4-8AE8-017538D2DF7E}" created="2024-12-30T20:26:06.503">
  <p188:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
      <a:r>
          <a:t>Needs more cowbell</a:t>
      </a:r>
      </a:p>
  </p188:txBody>
</p188:cm>
```

The following tables list the definitions of the possible child elements and attributes
of the `cm` (comment) element. For the complete definition see [MS-PPTX 2.16.3.3 CT_Comment](/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54)


| Attribute | Definition |
|---|---|
| id | Specifies the ID of a comment or a comment reply. |
| authorId | Specifies the author ID of a comment or a comment reply. |
| status | Specifies the status of a comment or a comment reply. |
| created | Specifies the date time when the comment or comment reply is created. |
| startDate | Specifies start date of the comment. |
| dueDate | Specifies due date of the comment. |
| assignedTo | Specifies a list of authors to whom the comment is assigned. |
| complete | Specifies the completion percentage of the comment. |
| title | Specifies the title for a comment. |

| Child Element | Definition |
|------------|---------------|
| pc:sldMkLst | Specifies a content moniker that identifies the slide to which the comment is anchored. |
| ac:deMkLst | Specifies a content moniker that identifies the drawing element to which the comment is anchored. |
| ac:txMkLst | Specifies a content moniker that identifies the text character range to which the comment is anchored. |
| unknownAnchor | Specifies an unknown anchor to which the comment is anchored. |
| pos | Specifies the position of the comment, relative to the top-left corner of the first object to which the comment is anchored. |
| replyLst | Specifies the list of replies to the comment. |
| txBody | Specifies the text of a comment or a comment reply. |
| extLst | Specifies a list of extensions for a comment or a comment reply. |



The following XML schema example defines the members of the `cm` element in addition to the required and
optional attributes.

```xml
 <xsd:complexType name="CT_Comment">
   <xsd:sequence>
     <xsd:group ref="EG_CommentAnchor" minOccurs="1" maxOccurs="1"/>
     <xsd:element name="pos" type="a:CT_Point2D" minOccurs="0" maxOccurs="1"/>
     <xsd:element name="replyLst" type="CT_CommentReplyList" minOccurs="0" maxOccurs="1"/>
     <xsd:group ref="EG_CommentProperties" minOccurs="1" maxOccurs="1"/>
   </xsd:sequence>
   <xsd:attributeGroup ref="AG_CommentProperties"/>
   <xsd:attribute name="startDate" type="xsd:dateTime" use="optional"/>
   <xsd:attribute name="dueDate" type="xsd:dateTime" use="optional"/>
   <xsd:attribute name="assignedTo" type="ST_AuthorIdList" use="optional" default=""/>
   <xsd:attribute name="complete" type="s:ST_PositiveFixedPercentage" default="0%" use="optional"/>
   <xsd:attribute name="title" type="xsd:string" use="optional" default=""/>
 </xsd:complexType>
```