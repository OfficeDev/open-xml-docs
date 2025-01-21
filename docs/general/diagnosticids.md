---
title: Diagnostic IDs
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/08/2025
ms.localizationpriority: high
---

# Diagnostic IDs

Diagnostic IDs are used to identify APIs or patterns that can raise compiler warnings or errors. This can be done via <xref:System.ObsoleteAttribute.DiagnosticId*> or <xref:System.Diagnostics.CodeAnalysis.ExperimentalAttribute>. These can be suppressed at the consumer level for each diagnostic id.

## Experimental APIs

### OOXML0001

**Title**: IPackage related APIs are currently experimental

As of v3.0, a new abstraction layer was added in between `System.IO.Packaging` and `DocumentFormat.OpenXml.Packaging.OpenXmlPackage`. This is currently experimental, but can be used if needed. This will be stabilized in a future release, and may or may not require code changes.

## Suppress warnings

It's recommended that you use an available workaround whenever possible. However, if you cannot change your code, you can suppress warnings through a `#pragma` directive or a `<NoWarn>` project setting. If you must use the obsolete or experimental APIs and the `OOXMLXXXX` diagnostic does not surface as an error, you can suppress the warning in code or in your project file.

To suppress the warnings in code:

```csharp
// Disable the warning.
#pragma warning disable OOXML0001

// Code that uses obsolete or experimental API.
//...

// Re-enable the warning.
#pragma warning restore OOXML0001
```

To suppress the warnings in a project file:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
   <TargetFramework>net6.0</TargetFramework>
   <!-- NoWarn below suppresses SYSLIB0001 project-wide -->
   <NoWarn>$(NoWarn);OOXML0001</NoWarn>
   <!-- To suppress multiple warnings, you can use multiple NoWarn elements -->
   <NoWarn>$(NoWarn);OOXML0001</NoWarn>
   <NoWarn>$(NoWarn);OTHER_WARNING</NoWarn>
   <!-- Alternatively, you can suppress multiple warnings by using a semicolon-delimited list -->
   <NoWarn>$(NoWarn);OOXML0001;OTHER_WARNING</NoWarn>
  </PropertyGroup>
</Project>
```

> [!NOTE]
> Suppressing warnings in this way only disables the obsoletion warnings you specify. It doesn't disable any other warnings, including obsoletion warnings with different diagnostic IDs.
