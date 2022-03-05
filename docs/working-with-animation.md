---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: d4ef73a6-888a-4476-9e21-4df76782127f
title: Working with animation (Open XML SDK)
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Working with animation (Open XML SDK)

This topic discusses the Open XML SDK 2.5 for Office <span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.Animate">**Animate** class and how it relates to the Open XML File Format PresentationML schema. For more information about the overall structure of the parts and elements that make up a PresentationML document, see [Structure of a PresentationML Document](structure-of-a-presentationml-document.md).

## Animation in PresentationML

The [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Animation section of the Open XML PresentationML framework as follows:

The Animation section of the PresentationML framework stores the movement and related information of objects. This schema is loosely based on the syntax and concepts from the Synchronized Multimedia Integration Language (SMIL), a W3C Recommendation for describing multimedia presentations using XML. The schema describes all the animations effects that reside on a slide and also the animation that occurs when going from slide to slide (slide transition). Animations on a slide are inherently time-based and consist of an animation effects on
an object or text. Slide transitions however do not follow this concept and always appear before any animation on a slide. All elements described in this schema are contained within the slide XML file. More specifically they are in the \<transition\> and the \<timing\> element as shown below:

```xml
<p:sld>  
    <p:cSld> … </p:cSld>  
    <p:clrMapOvr> … </p:clrMapOvr>  
    <p:transition> … </p:transition>  
    <p:timing> … </p:timing>  
</p:sld>
```

Animation consists of several behaviors, the most basic of which is the Animate behavior, represented by the \<anim\> element. The [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML \<anim\> element used to represent basic animation behavior in a PresentationML document as follows:

This element is a generic animation element that requires little or no semantic understanding of the attribute being animated. It can animate text within a shape or even the shape itself.[Example: Consider trying to emphasize text within a shape by changing the size of its font by 150%. The \<anim\> element should be used as follows:

```xml
<p:anim to="1.5" calcmode="lin" valueType="num">  
    <p:cBhvr override="childStyle">  
        <p:cTn id="1" dur="2000" fill="hold">  
        <p:tgtEl>  
            <p:spTgt spid="1">  
                <p:txEl>  
                    <p:charRg st="1" end="4">  
                </p:txEl>  
            </p:spTgt>  
        </p:tgtEl>  
        <p:attrNameLst>  
            <p:attrName>style.fontSize</p:attrName>  
        </p:attrNameLst>  
    </p:cBhvr>  
</p:anim>
```

The following table lists the child elements of the \<anim\> element used when working with animation and the Open XML SDK 2.5 classes that correspond to them.

| **PresentationML Element** | **Open XML SDK 2.5 Class** |
|:---------------------------|:----------------------------|
|         \<cBhvr\>          |       CommonBehavior       |
|         \<tavLst\>         |    TimeAnimateValueList    |

The following table from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the attributes of the \<anim\> element.

| **Attributes** | **Description**                                             |
|:---------------|:-----------------------------------------------------------------|
| by        | This attribute specifies a relative offset value for the animation with respect to its position before the start of the animation.The possible values for this attribute are defined by the W3C XML Schema string data type. |
| calcmode  | This attribute specifies the interpolation mode for the animation.The possible values for this attribute are defined by the ST_TLAnimateBehaviorCalcMode simple type (§19.7.20).      |
| from      | This attribute specifies the starting value of the animation.The possible values for this attribute are defined by the W3C XML Schema string data type.             |
| to        | This attribute specifies the ending value for the animation as a percentage.The possible values for this attribute are defined by the W3C XML Schema string data type.       |
| valueType | This attribute specifies the type of property value.The possible values for this attribute are defined by the ST_TLAnimateBehaviorValueType simple type (§19.7.21).           |

## Open XML SDK 2.5 Animate Class

The OXML SDK **Animate** class represents the \<anim\> element defined in the Open XML File Format schema for PresentationML documents. Use the **Animate**
class to manipulate individual \<anim\> elements in a PresentationML document.

Classes that represent child elements of the \<anim\> element and that are therefore commonly associated with the **Animate** class are shown in the following list.

### CommonBehavior Class

The **CommonBehavior** class corresponds to the \<cBhvr\> element. The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification introduces the \<cBhvr\>element:

This element describes the common behaviors of animations.  

Consider trying to emphasize text within a shape by changing the size of its font. The \<anim\> element should be used as follows:

```xml
<p:anim to="1.5" calcmode="lin" valueType="num">  
    <p:cBhvr override="childStyle">  
        <p:cTn id="6" dur="2000" fill="hold">  
        <p:tgtEl>  
            <p:spTgt spid="3">  
                <p:txEl>  
                   <p:charRg st="4294967295" end="4294967295"/>  
                </p:txEl>  
           </p:spTgt>  
        </p:tgtEl>  
        <p:attrNameLst>  
            <p:attrName>style.fontSize</p:attrName>  
        </p:attrNameLst>  
    </p:cBhvr>  
</p:anim>
```

### TimeAnimateValueList Class

The **TimeAnimateValueList** class corresponds to the \<tavLst\> element. The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification introduces the \<tavLst\> element:

This element specifies a list of time animated value elements.

Example: Consider a shape with a "fly-in" animation. The \<tav\> element should be used as follows:

```xml
<p:anim calcmode="lin" valueType="num">  
    <p:cBhvr additive="base"> … </p:cBhvr>  
    <p:tavLst>  
        <p:tav tm="0">  
            <p:val>  
                <p:strVal val="1+#ppt_h/2"/>  
            </p:val>  
        </p:tav>  
        <p:tav tm="100000"\>  
            <p:val>  
                <p:strVal val="#ppt_y">  
            </p:val>  
        </p:tav>  
    </p:tavLst>  
</p:anim>
```

## Working with the Animate Class

The **Animate** class, which represents the \<anim\> element, is therefore also associated with other classes that represent the child elements of the \<anim\> element, including the
**CommonBehavior** class, which describes common animation behaviors, and the **TimeAnimateValueList** class, which specifies a list of time-animated value elements, as shown in the previous XML code. Other classes associated with the **Animate** class are the <span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.Timing">**Timing** class, which specifies timing information for all the animations on the slide, and the <span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.TargetElement">**TargetElement** class, which specifies the target child elements to which the animation effects are applied.

## See also

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk.md)
[How to: Create a Presentation by Providing a File Name](how-to-create-a-presentation-document-by-providing-a-file-name.md)
[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)
[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)
[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  
