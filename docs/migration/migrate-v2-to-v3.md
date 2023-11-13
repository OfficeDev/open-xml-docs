---
title: 'Migrate from v2.x to v3.0'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Migration to v3.0.0

There are a number of breaking changes between v2.20.0 and v3.0.0 that may require source level changes. As a reminder, there are two kinds of breaking changes:

1. **Binary**: When a binary can no longer be used as drop in replacement
2. **Source**: When the source no longer compiles

The changes made in v3.0.0 were either a removal of obsoletions present in the SDK for a while or changes required for architectural reasons (most notably for better AOT support and trimming). Majority of these changes should be on the *binary* breaking change side, while still supporting compilation and expected behavior with your previous code. However, there are a few source breaking chnages to be made aware of.

## Breaking Changes

### .NET Standard 1.3 support has been dropped

No targets still in support require .NET Standard 1.3 and can use .NET Standard 2.0 instead. The project still supports .NET Framework 3.5+ and any [.NET Standard 2.0 supported platform](/dotnet/standard/net-standard?tabs=net-standard-2-0).

**Action needed**: If using .NET Standard 1.3, please upgrade to a supported version of .NET

### Target frameworks have changed

In order to simplify package creation, the TFMs built have been changed for some of the packages. However, there should be no apparent change to users as the overall supported platforms (besides .NET Standard 1.3 stated above) remains the same.

**Action needed**: None

### OpenXmlPart/OpenXmlContainer/OpenXmlPackage no longer have public constructors

These never initialized correct behavior and should never have been exposed.

**Action needed**: Use `.Create(...)` methods rather than constructor.

### Supporting framework for OpenXML types is now in the DocumentFormat.OpenXml.Framework package

Starting with v3.0.0, the supporting framework for the Open XML SDK is now within a standalone package, [DocumentFormat.OpenXml.Framework](https://www.nuget.org/packages/DocumentFormat.OpenXml.Framework).

**Action needed**: If you would like to operate on just `OpenXmlPackage` types, you no longer need to bring in all the static classes and can just reference the framework library.

### System.IO.Packaging is not directly used anymore

There have been issues with getting behavior we need from the System.IO.Packaging namespace. Starting with v3.0, a new set of interfaces in the `DocumentFormat.OpenXml.Packaging` namespace will be used to access package properties.

> NOTE: These types are currently marked as Obsolete, but only in the sense that we reserve the right to change their shape per feedback. Please be careful using these types as they may change in the future. At some point, we will remove the obsoletions and they will be considered stable APIs.

**Action needed**: If using `OpenXmlPackage.Package`, the package returned is no longer of type `System.IO.Packaging.Package`, but of `DocumentFormat.OpenXml.Packaging.IPackage`.

### Methods on parts to add child parts are now extension methods

There was a number of duplicated methods that would add parts in well defined ways. In order to consolidate this, if a part supports `ISupportedRelationship<T>`, extension methods can be written to support specific behavior that part can provide. Existing methods for this should transparently be retargeted to the new extension methods upon compilation.

**Action needed**: None

### OpenXmlAttribute is now a readonly struct

This type used to have mutable getters and setters. As a struct, this was easy to misuse, and should have been made readonly from the start.

**Action needed**: If expecting to mutate an OpenXmlAttribute in place, please create a new one instead.

### EnumValue&lt;TEnum&gt; now contains structs

Starting with v3.0.0, `EnumValue<T>` wraps a custom type that contains the information about the enum value. Previously, these types were stored in enum values in the C# type system, but required reflection to access, causing very large AOT compiled applications.

**Action needed**: Similar API surface is available, however the exposed enum values for this are no longer constants and will not be available in a few scenarios they had been (i.e. attribute values).

### OpenXmlElementList is now a struct

[OpenXmlElementList](/dotnet/api/documentformat.openxml.openxmlelementlist) is now a struct. It still implements `IEnumerable<OpenXmlElement>` in addition to `IReadOnlyList<OpenXmlElement>` where available.

**Action needed**: None

### IdPartPair is now a readonly struct

This type is used to enumerate pairs within a part and caused many unnecessary allocations. This change should be transparent upon recompilation.

**Action needed**: None

### OpenXmlPartReader no longer knows about all parts

In previous versions, [OpenXmlPartReader](/dotnet/api/documentformat.openxml.openxmlpartreader) knew about about all strongly typed part. In order to reduce coupling required for better AOT scenarios, we now have typed readers for known packages: `WordprocessingDocumentPartReader`, `SpreadsheetDocumentPartReader`, and `PresentationDocumentPartReader`.

**Action needed**: Replace usage of `OpenXmlPartReader` with document specific readers if needed. If creating a part reader from a known package, please use the constructors that take an existing `OpenXmlPart` which will then create the expected strongly typed parts.

### Attributes for schema information have been removed

`SchemaAttrAttribute` and `ChildElementInfoAttribute` have been removed from types and the types themselves are no longer present.

**Action needed**: If these types were required, please engage us at https://github.com/dotnet/open-xml-sdk to figure out the best way forward for you.

### OpenXmlPackage.Close has been removed

This did nothing useful besides call `.Dispose()`, but caused confusion about which should be called. This is now removed with the expectation of calling `.Dispose()`, preferably with the [using pattern](/dotnet/api/system.idisposable#using-an-object-that-implements-idisposable).

**Action needed**: Remove call and ensure package is disposed properly

### OpenXmlPackage.CanSave is now an instance property

This property used to be a static property that was dependent on the framework. Now, it may change per-package instance depending on settings and backing store.

**Action needed**: Replace usage of static property with instance.

### OpenXmlPackage.PartExtensionProvider has been changed

This property provided a dictionary that allowed access to change the extensions used. This is now backed by the `IPartExtensionFeature`.

**Action needed**: Replace usage with `OpenXmlPackage.Features.GetRequired<IPartExtensionFeature>()`.

### Packages with MarkupCompatibilityProcessMode.ProcessAllParts now actually process all parts

Previously, there was a heuristic to potentially minimize processing if no parts had been loaded. However, this caused scenarios such as ones where someone manually edited the XML to not actually process upon saving. v3.0.0 fixes this behavior and processes all part if this has been opted in.
