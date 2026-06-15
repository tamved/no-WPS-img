# Technical Notes: WPS `DISPIMG`

This document describes the spreadsheet format problem that `noWPS` solves.

## The Symptom

In WPS Office, a spreadsheet cell displays a picture.

In Microsoft Excel, the same cell may display one of these formulas:

```text
=@_xlfn.DISPIMG("ID_6103C1C6BD2E4DDC8EA68AC2DD3E75C1";1)
=DISPIMG("ID_6103C1C6BD2E4DDC8EA68AC2DD3E75C1",1)
```

Excel prefixes unknown functions with `_xlfn`, because `DISPIMG` is not a Microsoft Excel worksheet function.

## Where The Images Are

An `.xlsx` file is a ZIP package. WPS stores the image data inside the package, usually under:

```text
xl/media/image1.png
xl/media/image2.jpeg
...
```

WPS also writes a WPS-specific mapping file:

```text
xl/cellimages.xml
xl/_rels/cellimages.xml.rels
```

`cellimages.xml` contains records like:

```xml
<xdr:cNvPr id="108" name="ID_6103C1C6BD2E4DDC8EA68AC2DD3E75C1"/>
<a:blip r:embed="rId1"/>
```

`cellimages.xml.rels` maps `rId1` to a file such as:

```xml
<Relationship Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="media/image1.png"/>
```

The sheet cell contains the `DISPIMG` formula with the same `ID_...`.

## What `noWPS` Changes

`noWPS` creates a copy of the workbook and adds normal Excel drawing parts:

```text
xl/drawings/drawing1.xml
xl/drawings/_rels/drawing1.xml.rels
xl/worksheets/_rels/sheet1.xml.rels
```

It then adds `oneCellAnchor` pictures that reference the existing image files in `xl/media/`.

The unsupported `DISPIMG` formula is removed from the cell, so Excel no longer shows `_xlfn.DISPIMG`.

## Image Size Handling

The tool reads the real pixel dimensions of each PNG/JPEG image and computes a proportional thumbnail inside the original cell. This avoids flattened or stretched pictures in Excel.

## Why Not Patch Excel?

Excel does not have built-in support for the WPS `DISPIMG` function. A patch/add-in would be harder to distribute and may be blocked by company policy. `noWPS` uses the safer approach: convert the workbook to standard OpenXML that Excel already supports.
