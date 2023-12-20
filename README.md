[![npm version](https://badge.fury.io/js/cx-docx.svg)](https://www.npmjs.com/package/cx-docx)

# cx-docx
Docx rendering library

cx-docx is based on [docx-preview](https://github.com/VolodymyrBaydalka/docxjs)

Fixed:
+ anchor location conflicts with hash routing
+ some style issues
+ some numbering issues

Add:
+ comments
+ outline
+ quick jump to caption
+ preview endnote and footnote

Usage
-----
```html
<!--lib uses jszip-->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="cx-docx.min.js"></script>
<script>
    var docData = <document Blob>;

    docx.renderAsync(docData, document.getElementById("container"))
        .then(x => console.log("docx: finished"));
</script>
<body>
    ...
    <div id="container"></div>
    ...
</body>
```
API
---
```ts
renderAsync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        className: string = "docx", //class name/prefix for default and document style classes
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        ignoreWidth: boolean = false, //disables rendering width of page
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreFonts: boolean = false, //disables fonts rendering
        breakPages: boolean = true, //enables page breaking on page breaks
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        experimental: boolean = false, //enables experimental features (tab stops calculation)
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        renderChanges: false, //enables experimental rendering of document changes (inserions/deletions)
        renderHeaders: true, //enables headers rendering
        renderFooters: true, //enables footers rendering
        renderFootnotes: true, //enables footnotes rendering
        renderEndnotes: true, //enables endnotes rendering
        renderComments: false, // enables comments rendering
        renderOutline: false, // enables outline rendering
        renderNumbering: true, //enable numbering
        debug: boolean = false, //enables additional logging
    }): Promise<any>
```

