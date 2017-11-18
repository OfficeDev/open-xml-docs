---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fe780fcd-ed8f-4ee1-938e-cf3bb358ccae
title: Structure of a PresentationML document (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Structure of a PresentationML document (Open XML SDK)

The document structure of a PresentationML document consists of the
\<presentation\> (Presentation) element that contains \<sldMaster\>
(Slide Master), \<sldLayout\> (Slide Layout), \<sld \> (Slide), and
\<theme\> (Theme) elements that reference the slides in the
presentation. (The Theme element is the root element of the
DrawingMLTheme part.) These elements are the minimum elements required
for a valid presentation document.

In addition, a presentation document might contain \<notes\> (Notes
Slide), \<handoutMaster\> (Handout Master), \<sp\> (Shape), \<pic\>
(Picture), \<tbl\> (Table), and other slide-related elements. (Table
elements are defined in the DrawingML schema.)

Other features that a PresentationML document can contain include the
following: animation, audio, video, and transitions between slides.

A PresentationML document is not stored as one large body in a single
part. Instead, the elements that implement certain groupings of
functionality are stored in separate parts. For example, all comments in
a document are stored in one comment part, while each slide has its own
part. A separate XML file is created for each slide.


--------------------------------------------------------------------------------

Using the Open XML SDK 2.5, you can create document structure and
content that uses strongly-typed classes that correspond to
PresentationML elements. You can find these classes in the <span
sdata="cer" target="N:DocumentFormat.OpenXml.Presentation"><span
class="nolink">DocumentFormat.OpenXml.Presentation</span></span>
namespace. The following table lists the class names of the classes that
correspond to some of the important presentation elements.

**Package Part**|**Top Level PresentationML Element**|**Open XML SDK 2.5 Class**|**Description**
---|---|---|---
Presentation|<presentation>|Presentation|The root element for the Presentation part. This element specifies within it fundamental presentation-wide properties.
Presentation Properties|<presentationPr>|PresentationProperties|The root element for the Presentation Properties part. This element functions as a parent element within which additional presentation-wide document properties are contained.
Slide Master|<sldMaster>|SlideMaster|The root element for the Slide Master part. Within a slide master slide are contained all elements that describe the objects and their corresponding formatting for within a presentation slide. For more information, see Working with slide masters (Open XML SDK).
Slide Layout|<sldLayout>|SlideLayout|The root element for the Slide Layout part. This element specifies the relationship information for each slide layout that is used within the slide master. For more information, see Working with slide layouts (Open XML SDK).
Theme|<officeStyleSheet>|Theme|The root element for the Theme part. This element holds all the different formatting options available to a document through a theme and defines the overall look and feel of the document when themed objects are used within the document.
Slide|<sld>|Slide|The root element for the Slide part. This element specifies a slide within a slide list. For more information, see Working with presentation slides (Open XML SDK).
Notes Master|<notesMaster>|NotesMaster|The root element for the Notes Master part. Within a notes master slide are contained all elements that describe the objects and their corresponding formatting for within a notes slide.
Notes Slide|<notes>|NotesSlide|The root element of the Notes Slide part. This element specifies the existence of a notes slide along with its corresponding data. Contained within a notes slide are all the common slide elements along with addition properties that are specific to the notes element. For more information, see Working with notes slides (Open XML SDK).
Handout Master|<handoutMaster>|HandoutMaster|The root element of the Handout Master part. Within a handout master slide are contained all elements that describe the objects and their corresponding formatting for within a handout slide. For more information, see Working with handout master slides (Open XML SDK).
Comments|<cmLst>|CommentList|The root element of the Comments part. This element specifies a list of comments for a particular slide. For more information, see Working with comments (Open XML SDK).
Comments Author|<cmAuthorLst>|CommentAuthorList|The root element of the Comments Author part. This element specifies a list of authors with comments in the current document. For more information, see Working with comments (Open XML SDK).

 Descriptions adapted from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification, © ISO/IEC29500: 2008.

### Presentation Part

A PresentationML package's main part starts with a \<presentation\> root
element. That element contains a presentation, which, in turn, refers to
a slide list, a slide master list, a notes master list, and a handout
master list. The slide list refers to all of the slides in the
presentation. The slide master list refers to the entire set of slide
masters used in the presentation. The notes master contains information
about the formatting of notes pages. The handout master describes how a
handout looks. (A handout is a printed set of slides that can be handed
out to an audience for future reference.)

### Presentation Properties Part

The root element of the Presentation Properties part is the
\<presentationPr\> element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Presentation
Properties part as follows:

An instance of this part type contains all the presentation's
properties.

A package shall contain exactly one Presentation Properties part, and
that part shall be the target of an implicit relationship from the
Presentation (§13.3.6) part.

Example: The following Presentation part-relationship item contains a
relationship to the Presentation Properties part, which is stored in the
ZIP item presProps.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId6"  
        Type="http://…/presProps" Target="presProps.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be
presentationPr.

Example:

```xml
<p:presentationPr xmlns:p="…" …>  
    <p:clrMru>  
        …  
    </p:clrMru>  
    …  
</p:presentationPr>
```

A Presentation Properties part shall be located within the package
containing the relationships part (expressed syntactically, the
TargetMode attribute of the Relationship element shall be Internal).

A Presentation Properties part shall not have implicit or explicit
relationships to any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Slide Master Part

The root element of the Slide Master part is the \<sldMaster\> element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Slide Master part as
follows:

An instance of this part type contains the master definition of
formatting, text, and objects that appear on each slide in the
presentation that is derived from this slide master.

A package shall contain one or more Slide Master parts, each of which
shall be the target of an explicit relationship from the Presentation
(§13.3.6) part, as well as an implicit relationship from any Slide
Layout (§13.3.9) part where that slide layout is defined based on this
slide master. Each can optionally be the target of a relationship in a
Slide Layout (§13.3.9) part as well.

Example: The following Presentation part-relationship item contains a
relationship to the Slide Master part, which is stored in the ZIP item
slideMasters/slideMaster1.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId1"  
        Type="http://…/slideMaster"
Target="slideMasters/slideMaster1.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be sldMaster.

Example:

```xml
<p:sldMaster xmlns:p="…">  
    <p:cSld name="">  
        …  
    </p:cSld>  
    <p:clrMap … />  
</p:sldMaster>
```

A Slide Master part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).

A Slide Master part is permitted to have implicit relationships to the
following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1)  
• Bibliography (§15.2.3)  
• Custom XML Data Storage (§15.2.4)  
• Theme (§14.2.7)  
• Thumbnail (§15.2.16)

A Slide Master part is permitted to have explicit relationships to the
following parts defined by ISO/IEC 29500:

• Audio (§15.2.2)  
• Chart (§14.2.1)  
• Content Part (§15.2.4)  
• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram
Layout Definition (§14.2.5), and Diagram Styles (§14.2.6)  
• Embedded Control Persistence (§15.2.9)  
• Embedded Object (§15.2.10)  
• Embedded Package (§15.2.11)  
• Hyperlink (§15.3)  
• Image (§15.2.14)  
• Slide Layout (§13.3.9)  
• Video (§15.2.15)

A Slide Master part shall not have implicit or explicit relationships to
any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Slide Layout Part

The root element of the Slide Layout part is the \<sldLayout\> element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Slide Layout part as
follows:

An instance of this part type contains the definition for a slide layout
template for this presentation. This template defines the default
appearance and positioning of drawing objects on this slide type when it
is created.

A package shall contain one or more Slide Layout parts, and each of
those parts shall be the target of an explicit relationship in the Slide
Master (§13.3.10) part, as well as an implicit relationship from each of
the Slide (§13.3.8) parts associated with this slide layout.

Example: The following Slide Master part-relationship item contains
relationships to several Slide Layout parts, which are stored in the ZIP
items ../slideLayouts/slideLayoutN.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId1"  
        Type="http://…/slideLayout"  
        Target="../slideLayouts/slideLayout1.xml"/>  
    <Relationship Id="rId2"  
        Type="http://…/slideLayout"  
        Target="../slideLayouts/slideLayout2.xml"/>  
    <Relationship Id="rId3"  
        Type="http://…/slideLayout"  
        Target="../slideLayouts/slideLayout3.xml"/>  
</Relationships>
```


The root element for a part of this content type shall be sldLayout.

Example:

```xml
<p:sldLayout xmlns:p="…" matchingName="" type="title" preserve="1">  
    <p:cSld name="Title Slide">  
        …  
    </p:cSld>  
    <p:clrMapOvr>  
        <a:masterClrMapping/>  
    </p:clrMapOvr>  
    <p:timing/>  
</p:sldLayout>
```

A Slide Layout part is permitted to have implicit relationships to the
following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1)  
• Bibliography (§15.2.3)  
• Custom XML Data Storage (§15.2.4)  
• Slide Master (§13.3.10)  
• Theme Override (§14.2.8)  
• Thumbnail (§15.2.16)

A Slide Layout part is permitted to have explicit relationships to the
following parts defined by ISO/IEC 29500:

• Audio (§15.2.2)  
• Chart (§14.2.1)  
• Content Part (§15.2.4)  
• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram
Layout Definition (§14.2.5), and Diagram Styles (§14.2.6)  
• Embedded Control Persistence (§15.2.9)  
• Embedded Object (§15.2.10)  
• Embedded Package (§15.2.11)  
• Hyperlink (§15.3)  
• Image (§15.2.14)  
• Video (§15.2.15)

A Slide Layout part shall not have implicit or explicit relationships to
any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Slide Part

The root element of the Slide part is the \<sld\> element.

As well as text and graphics, each slide can contain comments and notes,
can have a layout, and can be part of one or more custom presentations.
A comment is an annotation intended for the person maintaining the
presentation slide deck. A note is a reminder or piece of text intended
for the presenter or the audience.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Slide part as
follows:

A Slide part contains the contents of a single slide.

A package shall contain one Slide part per slide, and each of those
parts shall be the target of an explicit relationship from the
Presentation (§13.3.6) part.

Example: Consider a PresentationML document having two slides. The
corresponding Presentation part relationship item contains two
relationships to Slide parts, which are stored in the ZIP items
slides/slide1.xml and slides/slide2.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId2"  
        Type="http://…/slide" Target="slides/slide1.xml"/>  
    <Relationship Id="rId3"  
        Type="http://…/slide" Target="slides/slide2.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be sld.

Example: slides/slide1.xml contains:

```xml
<p:sld xmlns:p="…">  
    <p:cSld name="">  
        …  
    </p:cSld>  
    <p:clrMapOvr>  
        …  
    </p:clrMapOvr>  
    <p:timing>  
        <p:tnLst>  
            <p:par>  
                <p:cTn id="1" dur="indefinite" restart="never"
nodeType="tmRoot"/>  
            </p:par>  
        </p:tnLst>  
    </p:timing>  
</p:sld>
```

A Slide part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).

A Slide part is permitted to have implicit relationships to the
following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1)  
• Bibliography (§15.2.3)  
• Comments (§13.3.2)  
• Custom XML Data Storage (§15.2.4)  
• Notes Slide (§13.3.5)  
• Theme Override (§14.2.8)  
• Thumbnail (§15.2.16)  
• Slide Layout (§13.3.9)  
• Slide Synchronization Data (§13.3.11)

A Slide part is permitted to have explicit relationships to the
following parts defined by ISO/IEC 29500:

• Audio (§15.2.2)  
• Chart (§14.2.1)  
• Content Part (§15.2.4)  
• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram
Layout Definition (§14.2.5), and Diagram Styles (§14.2.6)  
• Embedded Control Persistence (§15.2.9)  
• Embedded Object (§15.2.10)  
• Embedded Package (§15.2.11)  
• Hyperlink (§15.3)  
• Image (§15.2.14)  
• User Defined Tags (§13.3.12)  
• Video (§15.2.15)

A Slide part shall not have implicit or explicit relationships to any
other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Theme Part

The root element of the Theme part is the \<officeStyleSheet\> element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML DrawingML Theme part as follows:

An instance of this part type contains information about a document's
theme, which is a combination of color scheme, font scheme, and format
scheme (the latter also being referred to as effects). For a
WordprocessingML document, the choice of theme affects the color and
style of headings, among other things. For a SpreadsheetML document, the
choice of theme affects the color and style of cell contents and charts,
among other things. For a PresentationML document, the choice of theme
affects the formatting of slides, handouts, and notes via the associated
master, among other things.

A WordprocessingML or SpreadsheetML package shall contain zero or one
Theme part, which shall be the target of an implicit relationship in a
Main Document (§11.3.10) or Workbook (§12.3.23) part. A PresentationML
package shall contain zero or one Theme part per Handout Master
(§13.3.3), Notes Master (§13.3.4), Slide Master (§13.3.10) or
Presentation (§13.3.6) part via an implicit relationship.

Example: The following WordprocessingML Main Document part-relationship
item contains a relationship to the Theme part, which is stored in the
ZIP item theme/theme1.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId4"  
        Type="http://…/theme" Target="theme/theme1.xml"/>  
    </Relationships>
```

The root element for a part of this content type shall be
officeStyleSheet.

Example: theme1.xml contains the following, where the name attributes
of the clrScheme, fontScheme, and fmtScheme elements correspond to the
document's color scheme, font scheme, and format scheme, respectively:

```xml
<a:officeStyleSheet xmlns:a="…">  
    <a:baseStyles>  
        <a:clrScheme name="…">  
            …  
        </a:clrScheme>  
        <a:fontScheme name="…">  
            …  
        </a:fontScheme>  
        <a:fmtScheme name="…">  
            …  
        </a:fmtScheme>  
    </a:baseStyles>  
    <a:objectDefaults/>  
</a:officeStyleSheet>
```

A Theme part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).  

A Theme part is permitted to contain explicit relationships to the
following parts defined by ISO/IEC 29500:

• Image (§15.2.14)

A Theme part shall not have any implicit or explicit relationships to
other parts defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Notes Master Part

The root element of the Notes Master part is the \<notesMaster\>
element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Notes Master part as
follows:

An instance of this part type contains information about the content and
formatting of all notes pages.

A package shall contain at most one Notes Master part, and that part
shall be the target of an implicit relationship from the Notes Slide
(§13.3.5) part, as well as an explicit relationship from the
Presentation (§13.3.6) part.

Example: The following Presentation part-relationship item contains a
relationship to the Notes Master part, which is stored in the ZIP item
notesMasters/notesMaster1.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId4"  
        Type="http://…/notesMaster"
Target="notesMasters/notesMaster1.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be notesMaster.

Example:

```xml
<p:notesMaster xmlns:p="…">  
    <p:cSld name="">  
        …  
    </p:cSld\>  
    <p:clrMap … />  
</p:notesMaster>
```

A Notes Master part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).

A Notes Master part is permitted to have implicit relationships to the
following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1)  
• Bibliography (§15.2.3)  
• Custom XML Data Storage (§15.2.4)  
• Theme (§14.2.7)  
• Thumbnail (§15.2.16)

A Notes Master part is permitted to have explicit relationships to the
following parts defined by ISO/IEC 29500:

• Audio (§15.2.2)  
• Chart (§14.2.1)  
• Content Part (§15.2.4)  
• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram
Layout Definition (§14.2.5), and Diagram Styles (§14.2.6)  
• Embedded Control Persistence (§15.2.9)  
• Embedded Object (§15.2.10)  
• Embedded Package (§15.2.11)  
• Hyperlink (§15.3)  
• Image (§15.2.14)  
• Video (§15.2.15)

The Notes Master part shall not have implicit or explicit relationships
to any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Notes Slide Part

The root element of the Notes Slide part is the \<notes\> element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Notes Slide part as
follows:

An instance of this part type contains the notes for a single slide.

A package shall contain one Notes Slide part for each slide that
contains notes. If they exist, those parts shall each be the target of
an implicit relationship from the Slide (§13.3.8) part.

Example: The following Slide part-relationship item contains a
relationship to a Notes Slide part, which is stored in the ZIP item
../notesSlides/notesSlide1.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId3"  
        Type="http://…/notesSlide"
Target="../notesSlides/notesSlide1.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be notes.

Example:

```xml
<p:notes xmlns:p="…">  
    <p:cSld name="">  
         …  
    </p:cSld>  
    <p:clrMapOvr>  
        <a:masterClrMapping/>  
    </p:clrMapOvr>  
</p:notes>
```

A Notes Slide part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).

A Notes Slide part is permitted to have implicit relationships to the
following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1)  
• Bibliography (§15.2.3)  
• Custom XML Data Storage (§15.2.4)  
• Notes Master (§13.3.4)  
• Theme Override (§14.2.8)  
• Thumbnail (§15.2.16)

A Notes Slide part is permitted to have explicit relationships to the
following parts defined by ISO/IEC 29500:

• Audio (§15.2.2)  
• Chart (§14.2.1)  
• Content Part (§15.2.4)  
• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram
Layout Definition (§14.2.5), and Diagram Styles (§14.2.6)  
• Embedded Control Persistence (§15.2.9)  
• Embedded Object (§15.2.10)  
• Embedded Package (§15.2.11)  
• Hyperlink (§15.3)  
• Image (§15.2.14)  
• Video (§15.2.15)

The Notes Slide part shall not have implicit or explicit relationships
to any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Handout Master Part

The root element of the Handout Master part is the \<handoutMaster\>
element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Handout Master part
as follows:

An instance of this part type contains the look, position, and size of
the slides, notes, header and footer text, date, or page number on the
presentation's handout.

A package shall contain at most one Handout Master part, and it shall be
the target of an explicit relationship from the Presentation (§13.3.6)
part.

Example: The following Presentation part-relationship item contains a
relationship to the Handout Master part, which is stored in the ZIP item
handoutMasters/handoutMaster1.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId5"  
        Type="http://…/handoutMaster"  
        Target="handoutMasters/handoutMaster1.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be handoutMaster.

Example:

```xml
<p:handoutMaster xmlns:p="…">  
    <p:cSld name="">  
        …  
    </p:cSld\>  
    <p:clrMap … >  
</p:handoutMaster>
```

A Handout Master part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).

A Handout Master part is permitted to have implicit relationships to the
following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1)  
• Bibliography (§15.2.3)  
• Custom XML Data Storage (§15.2.4)  
• Theme (§14.2.7)  
• Thumbnail (§15.2.16)

A Handout Master part is permitted to have explicit relationships to the
following parts defined by ISO/IEC 29500:

• Audio (§15.2.2)  
• Chart (§14.2.1)  
• Content Part (§15.2.4)  
• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram
Layout Definition (§14.2.5), and Diagram Styles (§14.2.6)  
• Embedded Control Persistence (§15.2.9)  
• Embedded Object (§15.2.10)  
• Embedded Package (§15.2.11)  
• Hyperlink (§15.3)  
• Image (§15.2.14)  
• Video (§15.2.15)

A Handout Master part shall not have implicit or explicit relationships
to any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Comments Part

The root element of the Comments part is the \<cmLst\> element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Comments part as
follows:

An instance of this part type contains the comments for a single slide.
Each comment is tied to its author via an author-ID. Each comment's
index number and author-ID combination are unique.

A package shall contain one Comments part for each slide containing one
or more comments, and each of those parts shall be the target of an
implicit relationship from its corresponding Slide (§13.3.8) part.

Example: The following Slide part-relationship item contains a
relationship to a Comments part, which is stored in the ZIP item
../comments/comment2.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId4"  
        Type="http://…/comments"  
        Target="../comments/comment2.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be cmLst.

Example: The Comments part contains three comments, two created by one
author, and one created by another, all at the dates and times shown.
The index numbers are assigned on a per-author basis, starting at 1 for
an author's first comment:

```xml
<p:cmLst xmlns:p="…" …>  
    <p:cm authorId="0" dt="2005-11-13T17:00:22.071" idx="1">  
        <p:pos x="4486" y="1342"/>  
        <p:text>Comment text goes here.</p:text>  
    </p:cm>  
    <p:cm authorId="0" dt="2005-11-13T17:00:34.849" idx="2">  
        <p:pos x="3607" y="1867"/>  
        <p:text>Another comment's text goes here.</p:text>  
    </p:cm>  
    <p:cm authorId="1" dt="2005-11-15T00:06:46.919" idx="1">  
        <p:pos x="1493" y="2927"/>  
        <p:text>comment …</p:text>  
    </p:cm>  
</p:cmLst>
```

A Comments part shall be located within the package containing the
relationships part (expressed syntactically, the TargetMode attribute of
the Relationship element shall be Internal).

A Comments part shall not have implicit or explicit relationships to any
other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.

### Comments Author Part

The root element of the Comments Author part is the \<cmAuthorLst\>
element.

The [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification describes the Open XML PresentationML Comments Author part
as follows:

An instance of this part type contains information about each author who
has added a comment to the document. That information includes the
author's name, initials, a unique author-ID, a last-comment-index-used
count, and a display color. (The color can be used when displaying
comments to distinguish comments from different authors.)

A package shall contain at most one Comment Authors part. If it exists,
that part shall be the target of an implicit relationship from the
Presentation (§13.3.6) part.

Example: The following Presentation part relationship item contains a
relationship to the Comment Authors part, which is stored in the ZIP
item commentAuthors.xml:

```xml
<Relationships xmlns="…">  
    <Relationship Id="rId8"  
        Type="http://…/commentAuthors" Target="commentAuthors.xml"/>  
</Relationships>
```

The root element for a part of this content type shall be cmAuthorLst.

Example: Two people have authored comments in this document: Mary Smith
and Peter Jones. Her initials are "mas", her author-ID is 0, and her
comments' display color index is 0. Since Mary's last-comment-index-used
value is 3, the next comment-index to be used for her is 4. His initials
are "pjj", his author-ID is 1, and his comments' display color index is
1. Since Peter's last-comment-index-used value is 1, the next
comment-index to be used for him is 2:

```xml
<p:cmAuthorLst xmlns:p="…" …>  
    <p:cmAuthor id="0" name="Mary Smith" initials="mas" lastIdx="3"
clrIdx="0"/>  
    <p:cmAuthor id="1" name="Peter Jones" initials="pjj" lastIdx="1"
clrIdx="1"/>  
</p:cmAuthorLst>
```

A Comment Authors part shall be located within the package containing
the relationships part (expressed syntactically, the TargetMode
attribute of the Relationship element shall be Internal).

A Comment Authors part shall not have implicit or explicit relationships
to any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.


---------------------------------------------------------------------------------

Now that you are familiar with the parts of a PresentationML document,
consider how some of these parts are implemented and connected in an
actual presentation file. As shown in the article <span
sdata="link">[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md)</span>,
you can use the Open XML API to build up a minimum presentation file,
part by part.

A minimum presentation file consists of a presentation part, represented
by the file presentation.xml, as well as a presentation properties part
(presProps.xml), a slide master part (slideMaster.xml), a slide layout
part (slideLayout.xml), and a theme part (theme.xml). One or more slide
parts (slide.xml) are optional.

The packaging structure of a presentation document contains several
references between the parts, including some circular references. For
example, slide layouts reference slide masters, and slide masters
reference slide layouts.


---------------------------------------------------------------------------------

After you run the Open XML SDK 2.5 code to generate a presentation, you
can explore the contents of the .zip package to view the PresentationML
XML code. To view the .zip package, rename the extension on the minimum
presentation from **.pptx** to <span
class="keyword">.zip</span>. Inside the .zip package, there are several
parts that make up the minimum presentation.

Figure 1 shows the structure under the **ppt**
folder of the .zip package for a minimum presentation that contains a
single slide.

Figure 1. Minimum presentation folder structure

  
 ![Minimum presentation folder structure](./media/odc_oxml_ppt_documentstructure_fig01.jpg)
The presentation.xml file contains \<sld\> (Slide) elements that
reference the slides in the presentation. Each slide is associated to
the presentation by means of a slide ID and a relationship ID. The <span
class="keyword">slideID</span> is the identifier (ID) used within the
package to identify a slide and must be unique within the presentation.
The **id** attribute is the relationship ID
that identifies the slide part definition associated with a slide. For
more information about the slide part, see <span sdata="link">[Working with presentation slides (Open XML SDK)](working-with-presentation-slides.md)</span>.

The following XML code is the PresentationML that represents the
presentation part of a presentation document that contains a single
slide. This code is generated when you run the Open XML SDK 2.5 code to
create a minimum presentation

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:sldMasterIdLst>
        <p:sldMasterId id="2147483648"
                       r:id="rId1"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
      </p:sldMasterIdLst>
      <p:sldIdLst>
        <p:sldId id="256"
                 r:id="rId2"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
      </p:sldIdLst>
      <p:sldSz cx="9144000"
               cy="6858000"
               type="screen4x3" />
      <p:notesSz cx="6858000"
                 cy="9144000" />
      <p:defaultTextStyle />
    </p:presentation>
```
The following XML code is the PresentationML that represents the
relationship part of the presentation document. This code is generated
when you run the Open XML SDK 2.5 to create a minimum presentation.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
                    Target="/ppt/slides/slide.xml"
                    Id="rId2" />
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
                    Target="/ppt/slideLayouts/slideMasters/slideMaster.xml"
                    Id="rId1" />
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                    Target="/ppt/slideLayouts/slideMasters/theme/theme.xml"
                    Id="rId5" />
    </Relationships>
```
The following XML code is the PresentationML that represents the slide
part of the presentation document. Each slide in a presentation has a
slide part associated with it. This code is generated when you run the
Open XML SDK 2.5 to create a minimum presentation.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1"
                     name="" />
            <p:cNvGrpSpPr />
            <p:nvPr />
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
          </p:grpSpPr>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2"
                       name="Title 1" />
              <p:cNvSpPr>
                <a:spLocks noGrp="1"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph />
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr />
            <p:txBody>
              <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
              <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
              <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:endParaRPr lang="en-US" />
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
      </p:clrMapOvr>
    </p:sld>
```

---------------------------------------------------------------------------------

A typical presentation does not have a minimum configuration. A typical
presentation might contain several slides, each of which references
slide layouts and slide masters, and which might contain comments. In
addition, a presentation might contain handouts and notes slides, each
of which is represented by separate parts. These additional parts are
contained within the .zip package of the presentation document.

Figure 2 shows most of the elements that you would find in a typical
presentation.

Figure 2. Elements of a PresentationML file

  
 ![Elements of a PresentationML file](./media/odc_oxml_ppt_documentstructure_fig02.jpg)


--------------------------------------------------------------------------------

#### Concepts

[How to: Create a presentation document by providing a file name (Open XML SDK)](how-to-create-a-presentation-document-by-providing-a-file-name.md)  

[Working with presentations (Open XML SDK)](working-with-presentations.md)  

[Working with presentation slides (Open XML SDK)](working-with-presentation-slides.md)  

[Working with slide masters (Open XML SDK)](working-with-slide-masters.md)  

[Working with slide layouts (Open XML SDK)](working-with-slide-layouts.md)  
