import * as __WEBPACK_EXTERNAL_MODULE_jszip__ from "jszip";
/******/ var __webpack_modules__ = ({

/***/ "./src/assets/index.ts":
/*!*****************************!*\
  !*** ./src/assets/index.ts ***!
  \*****************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.createIconPrev = exports.createIconCollapse = exports.createIconPinglun = void 0;
const createIconPinglun = function () {
    let svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 1024 1024");
    svg.setAttribute("width", '20px');
    svg.setAttribute("height", '20px');
    let path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M512.13274 64C229.33274 64 0.13274 250.2 0.13274 480c0 95.2 39.8 182.4 105.8 252.6C76.13274 811.4 14.13274 878.2 13.13274 879c-13.2 14-16.8 34.4-9.2 52S28.93274 960 48.13274 960c123 0 220-51.4 278.2-92.6C384.13274 885.6 446.53274 896 512.13274 896c282.8 0 512-186.2 512-416S794.93274 64 512.13274 64z m0 736c-53.4 0-106.2-8.2-156.8-24.2l-45.4-14.4-39 27.6c-28.6 20.2-67.8 42.8-115 58 14.6-24.2 28.8-51.4 39.8-80.4l21.2-56.2-41.2-43.6C139.53274 628.2 96.13274 564.4 96.13274 480c0-176.4 186.6-320 416-320s416 143.6 416 320-186.6 320-416 320z");
    path.setAttribute("fill", "#777");
    svg.appendChild(path);
    return svg;
};
exports.createIconPinglun = createIconPinglun;
const createIconCollapse = function () {
    let svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 1024 1024");
    svg.setAttribute("width", '24px');
    svg.setAttribute("height", '24px');
    let path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M688.553 587.447L383.686 851.786c-20.841 18.07-52.391 15.83-70.468-5.003A49.927 49.927 0 0 1 301 814.063V209.936c0-27.58 22.365-49.937 49.955-49.937a49.965 49.965 0 0 1 32.731 12.213l304.867 264.34c41.683 36.141 46.165 99.219 10.01 140.887a99.891 99.891 0 0 1-10.01 10.007z");
    path.setAttribute("fill", "#fff");
    svg.appendChild(path);
    return svg;
};
exports.createIconCollapse = createIconCollapse;
const createIconPrev = function () {
    let svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 1024 1024");
    svg.setAttribute("width", '30px');
    svg.setAttribute("height", '30px');
    let path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M640 514.24L445.76 320l-48 48L544 514.24l-146.24 146.24 48 48L640 514.24z");
    path.setAttribute("fill", "#999");
    svg.appendChild(path);
    return svg;
};
exports.createIconPrev = createIconPrev;


/***/ }),

/***/ "./src/comments/comments-part.ts":
/*!***************************************!*\
  !*** ./src/comments/comments-part.ts ***!
  \***************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.CommentsPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const comments_1 = __webpack_require__(/*! ./comments */ "./src/comments/comments.ts");
class CommentsPart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    parseXml(root) {
        Object.assign(this, (0, comments_1.parseCommentsPart)(root, this._package.xmlParser));
    }
}
exports.CommentsPart = CommentsPart;


/***/ }),

/***/ "./src/comments/comments.ts":
/*!**********************************!*\
  !*** ./src/comments/comments.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ParseCommentText = exports.parseComment = exports.parseCommentsPart = void 0;
function parseCommentsPart(elem, xml) {
    let result = {
        comments: []
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "comment":
                result.comments.push(parseComment(e, xml));
                break;
        }
    }
    return result;
}
exports.parseCommentsPart = parseCommentsPart;
function parseComment(elem, xml) {
    let result = {
        id: xml.attr(elem, "id"),
        author: xml.attr(elem, "author"),
        date: xml.attr(elem, "date"),
        paraId: xml.elementAttr(elem, "p", "paraId"),
        text: ParseCommentText(xml.element(elem, "p"), xml),
        noRender: false,
        msg: '',
        type: '',
        children: []
    };
    return result;
}
exports.parseComment = parseComment;
function ParseCommentText(elem, xml) {
    let result = '';
    for (let e of xml.elements(elem, "r")) {
        for (let t of xml.elements(e, "t")) {
            result += t.textContent;
        }
    }
    return result;
}
exports.ParseCommentText = ParseCommentText;


/***/ }),

/***/ "./src/comments/commentsExtended-part.ts":
/*!***********************************************!*\
  !*** ./src/comments/commentsExtended-part.ts ***!
  \***********************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.CommentsExtendedPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const commentsExtended_1 = __webpack_require__(/*! ./commentsExtended */ "./src/comments/commentsExtended.ts");
class CommentsExtendedPart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    parseXml(root) {
        Object.assign(this, (0, commentsExtended_1.parseCommentsExtendedPart)(root, this._package.xmlParser));
    }
}
exports.CommentsExtendedPart = CommentsExtendedPart;


/***/ }),

/***/ "./src/comments/commentsExtended.ts":
/*!******************************************!*\
  !*** ./src/comments/commentsExtended.ts ***!
  \******************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCommentEx = exports.parseCommentsExtendedPart = void 0;
function parseCommentsExtendedPart(elem, xml) {
    let result = {
        commentsEx: []
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "commentEx":
                result.commentsEx.push(parseCommentEx(e, xml));
                break;
        }
    }
    return result;
}
exports.parseCommentsExtendedPart = parseCommentsExtendedPart;
function parseCommentEx(elem, xml) {
    let result = {
        paraIdParent: xml.attr(elem, "paraIdParent"),
        paraId: xml.attr(elem, "paraId")
    };
    return result;
}
exports.parseCommentEx = parseCommentEx;


/***/ }),

/***/ "./src/common/open-xml-package.ts":
/*!****************************************!*\
  !*** ./src/common/open-xml-package.ts ***!
  \****************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.OpenXmlPackage = void 0;
const JSZip = __webpack_require__(/*! jszip */ "jszip");
const xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
const utils_1 = __webpack_require__(/*! ../utils */ "./src/utils.ts");
const relationship_1 = __webpack_require__(/*! ./relationship */ "./src/common/relationship.ts");
class OpenXmlPackage {
    constructor(_zip, options) {
        this._zip = _zip;
        this.options = options;
        this.xmlParser = new xml_parser_1.XmlParser();
    }
    get(path) {
        return this._zip.files[normalizePath(path)];
    }
    update(path, content) {
        this._zip.file(path, content);
    }
    static async load(input, options) {
        const zip = await JSZip.loadAsync(input);
        return new OpenXmlPackage(zip, options);
    }
    save(type = "blob") {
        return this._zip.generateAsync({ type });
    }
    load(path, type = "string") {
        return this.get(path)?.async(type) ?? Promise.resolve(null);
    }
    async loadRelationships(path = null) {
        let relsPath = `_rels/.rels`;
        if (path != null) {
            const [f, fn] = (0, utils_1.splitPath)(path);
            relsPath = `${f}_rels/${fn}.rels`;
        }
        const txt = await this.load(relsPath);
        return txt ? (0, relationship_1.parseRelationships)(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
    }
    parseXmlDocument(txt) {
        return (0, xml_parser_1.parseXmlString)(txt, this.options.trimXmlDeclaration);
    }
}
exports.OpenXmlPackage = OpenXmlPackage;
function normalizePath(path) {
    return path.startsWith('/') ? path.substr(1) : path;
}


/***/ }),

/***/ "./src/common/part.ts":
/*!****************************!*\
  !*** ./src/common/part.ts ***!
  \****************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.Part = void 0;
const xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
class Part {
    constructor(_package, path) {
        this._package = _package;
        this.path = path;
    }
    async load() {
        this.rels = await this._package.loadRelationships(this.path);
        const xmlText = await this._package.load(this.path);
        const xmlDoc = this._package.parseXmlDocument(xmlText);
        if (this._package.options.keepOrigin) {
            this._xmlDocument = xmlDoc;
        }
        this.parseXml(xmlDoc.firstElementChild);
    }
    save() {
        this._package.update(this.path, (0, xml_parser_1.serializeXmlString)(this._xmlDocument));
    }
    parseXml(root) {
    }
}
exports.Part = Part;


/***/ }),

/***/ "./src/common/relationship.ts":
/*!************************************!*\
  !*** ./src/common/relationship.ts ***!
  \************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseRelationships = exports.RelationshipTypes = void 0;
var RelationshipTypes;
(function (RelationshipTypes) {
    RelationshipTypes["OfficeDocument"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    RelationshipTypes["FontTable"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    RelationshipTypes["Image"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    RelationshipTypes["Numbering"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    RelationshipTypes["Styles"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    RelationshipTypes["StylesWithEffects"] = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
    RelationshipTypes["Theme"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    RelationshipTypes["Settings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    RelationshipTypes["WebSettings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    RelationshipTypes["Hyperlink"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    RelationshipTypes["Footnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    RelationshipTypes["Endnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    RelationshipTypes["Footer"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    RelationshipTypes["Header"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    RelationshipTypes["ExtendedProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    RelationshipTypes["CoreProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    RelationshipTypes["CustomProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties";
    RelationshipTypes["Comments"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    RelationshipTypes["CommentsExtended"] = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    RelationshipTypes["CommentsIds"] = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
    RelationshipTypes["CommentsExtensible"] = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
})(RelationshipTypes || (exports.RelationshipTypes = RelationshipTypes = {}));
function parseRelationships(root, xml) {
    return xml.elements(root).map(e => ({
        id: xml.attr(e, "Id"),
        type: xml.attr(e, "Type"),
        target: xml.attr(e, "Target"),
        targetMode: xml.attr(e, "TargetMode")
    }));
}
exports.parseRelationships = parseRelationships;


/***/ }),

/***/ "./src/cx-docx.ts":
/*!************************!*\
  !*** ./src/cx-docx.ts ***!
  \************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.renderAsync = exports.praseAsync = exports.defaultOptions = void 0;
const word_document_1 = __webpack_require__(/*! ./word-document */ "./src/word-document.ts");
const document_parser_1 = __webpack_require__(/*! ./document-parser */ "./src/document-parser.ts");
const html_renderer_1 = __webpack_require__(/*! ./html-renderer */ "./src/html-renderer.ts");
exports.defaultOptions = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    useBase64URL: false,
    renderChanges: false,
    renderComments: false,
    renderOutline: false,
    renderNumbering: true,
};
function praseAsync(data, userOptions = null) {
    const ops = { ...exports.defaultOptions, ...userOptions };
    return word_document_1.WordDocument.load(data, new document_parser_1.DocumentParser(ops), ops);
}
exports.praseAsync = praseAsync;
async function renderAsync(data, bodyContainer, styleContainer = null, userOptions = null) {
    const ops = { ...exports.defaultOptions, ...userOptions };
    const renderer = new html_renderer_1.HtmlRenderer(window.document);
    const doc = await word_document_1.WordDocument.load(data, new document_parser_1.DocumentParser(ops), ops);
    renderer.render(doc, bodyContainer, styleContainer, ops);
    return doc;
}
exports.renderAsync = renderAsync;


/***/ }),

/***/ "./src/document-parser.ts":
/*!********************************!*\
  !*** ./src/document-parser.ts ***!
  \********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.DocumentParser = exports.autos = void 0;
const dom_1 = __webpack_require__(/*! ./document/dom */ "./src/document/dom.ts");
const paragraph_1 = __webpack_require__(/*! ./document/paragraph */ "./src/document/paragraph.ts");
const section_1 = __webpack_require__(/*! ./document/section */ "./src/document/section.ts");
const xml_parser_1 = __webpack_require__(/*! ./parser/xml-parser */ "./src/parser/xml-parser.ts");
const run_1 = __webpack_require__(/*! ./document/run */ "./src/document/run.ts");
const bookmarks_1 = __webpack_require__(/*! ./document/bookmarks */ "./src/document/bookmarks.ts");
const common_1 = __webpack_require__(/*! ./document/common */ "./src/document/common.ts");
const vml_1 = __webpack_require__(/*! ./vml/vml */ "./src/vml/vml.ts");
const comments_1 = __webpack_require__(/*! ./document/comments */ "./src/document/comments.ts");
exports.autos = {
    shd: "inherit",
    color: "black",
    borderColor: "black",
    highlight: "transparent"
};
const supportedNamespaceURIs = [];
const mmlTagMap = {
    "oMath": dom_1.DomType.MmlMath,
    "oMathPara": dom_1.DomType.MmlMathParagraph,
    "f": dom_1.DomType.MmlFraction,
    "func": dom_1.DomType.MmlFunction,
    "fName": dom_1.DomType.MmlFunctionName,
    "num": dom_1.DomType.MmlNumerator,
    "den": dom_1.DomType.MmlDenominator,
    "rad": dom_1.DomType.MmlRadical,
    "deg": dom_1.DomType.MmlDegree,
    "e": dom_1.DomType.MmlBase,
    "sSup": dom_1.DomType.MmlSuperscript,
    "sSub": dom_1.DomType.MmlSubscript,
    "sPre": dom_1.DomType.MmlPreSubSuper,
    "sup": dom_1.DomType.MmlSuperArgument,
    "sub": dom_1.DomType.MmlSubArgument,
    "d": dom_1.DomType.MmlDelimiter,
    "nary": dom_1.DomType.MmlNary,
    "eqArr": dom_1.DomType.MmlEquationArray,
    "lim": dom_1.DomType.MmlLimit,
    "limLow": dom_1.DomType.MmlLimitLower,
    "m": dom_1.DomType.MmlMatrix,
    "mr": dom_1.DomType.MmlMatrixRow,
    "box": dom_1.DomType.MmlBox,
    "bar": dom_1.DomType.MmlBar,
    "groupChr": dom_1.DomType.MmlGroupChar
};
class DocumentParser {
    constructor(options) {
        this.options = {
            ignoreWidth: false,
            debug: false,
            ...options
        };
    }
    parseNotes(xmlDoc, elemName, elemClass) {
        var result = [];
        for (let el of xml_parser_1.default.elements(xmlDoc, elemName)) {
            const node = new elemClass();
            node.id = xml_parser_1.default.attr(el, "id");
            node.noteType = xml_parser_1.default.attr(el, "type");
            node.children = this.parseBodyElements(el);
            result.push(node);
        }
        return result;
    }
    parseDocumentFile(xmlDoc) {
        var xbody = xml_parser_1.default.element(xmlDoc, "body");
        var background = xml_parser_1.default.element(xmlDoc, "background");
        var sectPr = xml_parser_1.default.element(xbody, "sectPr");
        return {
            type: dom_1.DomType.Document,
            children: this.parseBodyElements(xbody),
            props: sectPr ? (0, section_1.parseSectionProperties)(sectPr, xml_parser_1.default) : {},
            cssStyle: background ? this.parseBackground(background) : {},
        };
    }
    parseBackground(elem) {
        var result = {};
        var color = xmlUtil.colorAttr(elem, "color");
        if (color) {
            result["background-color"] = color;
        }
        return result;
    }
    parseBodyElements(element) {
        var children = [];
        for (let elem of xml_parser_1.default.elements(element)) {
            switch (elem.localName) {
                case "p":
                    children.push(this.parseParagraph(elem));
                    break;
                case "tbl":
                    children.push(this.parseTable(elem));
                    break;
                case "sdt":
                    children.push(...this.parseSdt(elem, e => this.parseBodyElements(e)));
                    break;
                case "commentRangeEnd":
                    children.push((0, comments_1.parseCommentRangeEnd)(elem, xml_parser_1.default));
                    break;
            }
        }
        return children;
    }
    parseStylesFile(xstyles) {
        var result = [];
        xmlUtil.foreach(xstyles, n => {
            switch (n.localName) {
                case "style":
                    result.push(this.parseStyle(n));
                    break;
                case "docDefaults":
                    result.push(this.parseDefaultStyles(n));
                    break;
            }
        });
        return result;
    }
    parseDefaultStyles(node) {
        var result = {
            id: null,
            name: null,
            target: null,
            basedOn: null,
            styles: []
        };
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "rPrDefault":
                    var rPr = xml_parser_1.default.element(c, "rPr");
                    if (rPr)
                        result.styles.push({
                            target: "span",
                            values: this.parseDefaultProperties(rPr, {})
                        });
                    break;
                case "pPrDefault":
                    var pPr = xml_parser_1.default.element(c, "pPr");
                    if (pPr)
                        result.styles.push({
                            target: "p",
                            values: this.parseDefaultProperties(pPr, {})
                        });
                    break;
            }
        });
        return result;
    }
    parseStyle(node) {
        var result = {
            id: xml_parser_1.default.attr(node, "styleId"),
            isDefault: xml_parser_1.default.boolAttr(node, "default"),
            name: null,
            target: null,
            basedOn: null,
            styles: [],
            linked: null
        };
        switch (xml_parser_1.default.attr(node, "type")) {
            case "paragraph":
                result.target = "p";
                break;
            case "table":
                result.target = "table";
                break;
            case "character":
                result.target = "span";
                break;
        }
        xmlUtil.foreach(node, n => {
            switch (n.localName) {
                case "basedOn":
                    result.basedOn = xml_parser_1.default.attr(n, "val");
                    break;
                case "name":
                    result.name = xml_parser_1.default.attr(n, "val");
                    break;
                case "link":
                    result.linked = xml_parser_1.default.attr(n, "val");
                    break;
                case "next":
                    result.next = xml_parser_1.default.attr(n, "val");
                    break;
                case "aliases":
                    result.aliases = xml_parser_1.default.attr(n, "val").split(",");
                    break;
                case "pPr":
                    result.styles.push({
                        target: "p",
                        values: this.parseDefaultProperties(n, {})
                    });
                    result.paragraphProps = (0, paragraph_1.parseParagraphProperties)(n, xml_parser_1.default);
                    break;
                case "rPr":
                    result.styles.push({
                        target: "span",
                        values: this.parseDefaultProperties(n, {})
                    });
                    result.runProps = (0, run_1.parseRunProperties)(n, xml_parser_1.default);
                    break;
                case "tblPr":
                case "tcPr":
                    result.styles.push({
                        target: "td",
                        values: this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblStylePr":
                    for (let s of this.parseTableStyle(n))
                        result.styles.push(s);
                    break;
                case "rsid":
                case "qFormat":
                case "hidden":
                case "semiHidden":
                case "unhideWhenUsed":
                case "autoRedefine":
                case "uiPriority":
                    break;
                default:
                    this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
            }
        });
        return result;
    }
    parseTableStyle(node) {
        var result = [];
        var type = xml_parser_1.default.attr(node, "type");
        var selector = "";
        var modificator = "";
        switch (type) {
            case "firstRow":
                modificator = ".first-row";
                selector = "tr.first-row td";
                break;
            case "lastRow":
                modificator = ".last-row";
                selector = "tr.last-row td";
                break;
            case "firstCol":
                modificator = ".first-col";
                selector = "td.first-col";
                break;
            case "lastCol":
                modificator = ".last-col";
                selector = "td.last-col";
                break;
            case "band1Vert":
                modificator = ":not(.no-vband)";
                selector = "td.odd-col";
                break;
            case "band2Vert":
                modificator = ":not(.no-vband)";
                selector = "td.even-col";
                break;
            case "band1Horz":
                modificator = ":not(.no-hband)";
                selector = "tr.odd-row";
                break;
            case "band2Horz":
                modificator = ":not(.no-hband)";
                selector = "tr.even-row";
                break;
            default: return [];
        }
        xmlUtil.foreach(node, n => {
            switch (n.localName) {
                case "pPr":
                    result.push({
                        target: `${selector} p`,
                        mod: modificator,
                        values: this.parseDefaultProperties(n, {})
                    });
                    break;
                case "rPr":
                    result.push({
                        target: `${selector} span`,
                        mod: modificator,
                        values: this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblPr":
                case "tcPr":
                    result.push({
                        target: selector,
                        mod: modificator,
                        values: this.parseDefaultProperties(n, {})
                    });
                    break;
            }
        });
        return result;
    }
    processAbstractNumbering(abstractNumberings, numberings) {
        for (let numbering of numberings) {
            if (numbering.overrides.length > 0) {
                let abstractNumberingLevels = abstractNumberings.filter(item => {
                    return item.id === numbering.abstractId;
                });
                for (let override of numbering.overrides) {
                    let originAbstractNumberingLevels = abstractNumberings.filter(item => {
                        return item.id === String(override.start);
                    });
                    let originLevel = originAbstractNumberingLevels?.find(item => {
                        return item.level === override.level;
                    });
                    let level = abstractNumberingLevels?.find(item => {
                        return item.level === override.level;
                    });
                    if (originLevel) {
                        level.format = originLevel.format;
                        level.justification = originLevel.justification;
                        level.paragraphProps = originLevel.paragraphProps;
                        level.runProps = originLevel.runProps;
                        level.text = originLevel.text;
                    }
                }
            }
        }
    }
    parseNumberingFile(xnums, numberings) {
        var result = [];
        var mapping = {};
        var bullets = [];
        xmlUtil.foreach(xnums, n => {
            switch (n.localName) {
                case "abstractNum":
                    this.parseAbstractNumbering(n, bullets)
                        .forEach(x => result.push(x));
                    break;
                case "numPicBullet":
                    bullets.push(this.parseNumberingPicBullet(n));
                    break;
                case "num":
                    var numId = xml_parser_1.default.attr(n, "numId");
                    var abstractNumId = xml_parser_1.default.elementAttr(n, "abstractNumId", "val");
                    mapping[numId] = abstractNumId;
                    break;
            }
        });
        let myResult = [];
        for (let key in mapping) {
            let absId = mapping[key];
            let absList = result.filter(item => {
                return item.id === absId;
            });
            let myAbsList = JSON.parse(JSON.stringify(absList));
            for (let abs of myAbsList) {
                abs.numId = key;
            }
            myResult = myResult.concat(myAbsList);
        }
        return myResult;
    }
    parseNumberingPicBullet(elem) {
        var pict = xml_parser_1.default.element(elem, "pict");
        var shape = pict && xml_parser_1.default.element(pict, "shape");
        var imagedata = shape && xml_parser_1.default.element(shape, "imagedata");
        return imagedata ? {
            id: xml_parser_1.default.intAttr(elem, "numPicBulletId"),
            src: xml_parser_1.default.attr(imagedata, "id"),
            style: xml_parser_1.default.attr(shape, "style")
        } : null;
    }
    parseAbstractNumbering(node, bullets) {
        var result = [];
        var id = xml_parser_1.default.attr(node, "abstractNumId");
        xmlUtil.foreach(node, n => {
            switch (n.localName) {
                case "lvl":
                    result.push(this.parseNumberingLevel(id, n, bullets));
                    break;
            }
        });
        return result;
    }
    parseNumberingLevel(id, node, bullets) {
        var result = {
            id: id,
            level: xml_parser_1.default.intAttr(node, "ilvl"),
            start: 1,
            pStyleName: undefined,
            pStyle: {},
            rStyle: {},
            suff: "tab"
        };
        xmlUtil.foreach(node, n => {
            switch (n.localName) {
                case "start":
                    result.start = xml_parser_1.default.intAttr(n, "val");
                    break;
                case "pPr":
                    this.parseDefaultProperties(n, result.pStyle);
                    break;
                case "rPr":
                    this.parseDefaultProperties(n, result.rStyle);
                    break;
                case "lvlPicBulletId":
                    var id = xml_parser_1.default.intAttr(n, "val");
                    result.bullet = bullets.find(x => x.id == id);
                    break;
                case "lvlText":
                    result.levelText = xml_parser_1.default.attr(n, "val");
                    break;
                case "pStyle":
                    result.pStyleName = xml_parser_1.default.attr(n, "val");
                    break;
                case "numFmt":
                    result.format = xml_parser_1.default.attr(n, "val");
                    break;
                case "suff":
                    result.suff = xml_parser_1.default.attr(n, "val");
                    break;
            }
        });
        return result;
    }
    parseSdt(node, parser) {
        const sdtContent = xml_parser_1.default.element(node, "sdtContent");
        return sdtContent ? parser(sdtContent) : [];
    }
    parseInserted(node, parentParser) {
        return {
            type: dom_1.DomType.Inserted,
            children: parentParser(node)?.children ?? []
        };
    }
    parseDeleted(node, parentParser) {
        return {
            type: dom_1.DomType.Deleted,
            children: parentParser(node)?.children ?? []
        };
    }
    parseParagraph(node) {
        var result = { type: dom_1.DomType.Paragraph, children: [] };
        for (let el of xml_parser_1.default.elements(node)) {
            switch (el.localName) {
                case "pPr":
                    this.parseParagraphProperties(el, result);
                    break;
                case "r":
                    result.children.push(this.parseRun(el, result));
                    break;
                case "hyperlink":
                    result.children.push(this.parseHyperlink(el, result));
                    break;
                case "bookmarkStart":
                    result.children.push((0, bookmarks_1.parseBookmarkStart)(el, xml_parser_1.default));
                    break;
                case "bookmarkEnd":
                    result.children.push((0, bookmarks_1.parseBookmarkEnd)(el, xml_parser_1.default));
                    break;
                case "oMath":
                case "oMathPara":
                    result.children.push(this.parseMathElement(el));
                    break;
                case "sdt":
                    result.children.push(...this.parseSdt(el, e => this.parseParagraph(e).children));
                    break;
                case "ins":
                    result.children.push(this.parseInserted(el, e => this.parseParagraph(e)));
                    break;
                case "del":
                    result.children.push(this.parseDeleted(el, e => this.parseParagraph(e)));
                    break;
                case "commentRangeStart":
                    result.children.push((0, comments_1.parseCommentRangeStart)(el, xml_parser_1.default));
                    break;
                case "commentRangeEnd":
                    result.children.push((0, comments_1.parseCommentRangeEnd)(el, xml_parser_1.default));
                    break;
            }
        }
        return result;
    }
    parseParagraphProperties(elem, paragraph) {
        this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
            if ((0, paragraph_1.parseParagraphProperty)(c, paragraph, xml_parser_1.default))
                return true;
            switch (c.localName) {
                case "pStyle":
                    paragraph.styleName = xml_parser_1.default.attr(c, "val");
                    break;
                case "cnfStyle":
                    paragraph.className = values.classNameOfCnfStyle(c);
                    break;
                case "framePr":
                    this.parseFrame(c, paragraph);
                    break;
                case "rPr":
                    break;
                default:
                    return false;
            }
            return true;
        });
    }
    parseFrame(node, paragraph) {
        var dropCap = xml_parser_1.default.attr(node, "dropCap");
        if (dropCap == "drop")
            paragraph.cssStyle["float"] = "left";
    }
    parseHyperlink(node, parent) {
        var result = { type: dom_1.DomType.Hyperlink, parent: parent, children: [] };
        var anchor = xml_parser_1.default.attr(node, "anchor");
        var relId = xml_parser_1.default.attr(node, "id");
        if (anchor)
            result.href = "#" + anchor;
        if (relId)
            result.id = relId;
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "r":
                    result.children.push(this.parseRun(c, result));
                    break;
            }
        });
        return result;
    }
    parseRun(node, parent) {
        var result = { type: dom_1.DomType.Run, parent: parent, children: [] };
        xmlUtil.foreach(node, c => {
            c = this.checkAlternateContent(c);
            switch (c.localName) {
                case "t":
                    result.children.push({
                        type: dom_1.DomType.Text,
                        text: c.textContent
                    });
                    break;
                case "delText":
                    result.children.push({
                        type: dom_1.DomType.DeletedText,
                        text: c.textContent
                    });
                    break;
                case "fldSimple":
                    result.children.push({
                        type: dom_1.DomType.SimpleField,
                        instruction: xml_parser_1.default.attr(c, "instr"),
                        lock: xml_parser_1.default.boolAttr(c, "lock", false),
                        dirty: xml_parser_1.default.boolAttr(c, "dirty", false)
                    });
                    break;
                case "instrText":
                    result.fieldRun = true;
                    result.children.push({
                        type: dom_1.DomType.Instruction,
                        text: c.textContent
                    });
                    break;
                case "fldChar":
                    result.fieldRun = true;
                    result.children.push({
                        type: dom_1.DomType.ComplexField,
                        charType: xml_parser_1.default.attr(c, "fldCharType"),
                        lock: xml_parser_1.default.boolAttr(c, "lock", false),
                        dirty: xml_parser_1.default.boolAttr(c, "dirty", false)
                    });
                    break;
                case "noBreakHyphen":
                    result.children.push({ type: dom_1.DomType.NoBreakHyphen });
                    break;
                case "br":
                    result.children.push({
                        type: dom_1.DomType.Break,
                        break: xml_parser_1.default.attr(c, "type") || "textWrapping"
                    });
                    break;
                case "lastRenderedPageBreak":
                    result.children.push({
                        type: dom_1.DomType.Break,
                        break: "lastRenderedPageBreak"
                    });
                    break;
                case "sym":
                    result.children.push({
                        type: dom_1.DomType.Symbol,
                        font: xml_parser_1.default.attr(c, "font"),
                        char: xml_parser_1.default.attr(c, "char")
                    });
                    break;
                case "tab":
                    result.children.push({ type: dom_1.DomType.Tab });
                    break;
                case "footnoteReference":
                    result.children.push({
                        type: dom_1.DomType.FootnoteReference,
                        id: xml_parser_1.default.attr(c, "id")
                    });
                    break;
                case "endnoteReference":
                    result.children.push({
                        type: dom_1.DomType.EndnoteReference,
                        id: xml_parser_1.default.attr(c, "id")
                    });
                    break;
                case "drawing":
                    let d = this.parseDrawing(c);
                    if (d)
                        result.children = [d];
                    break;
                case "pict":
                    result.children.push(this.parseVmlPicture(c));
                    break;
                case "rPr":
                    this.parseRunProperties(c, result);
                    break;
                case "commentReference":
                    let id = xml_parser_1.default.attr(c, "id");
                    result.children.push({
                        type: dom_1.DomType.CommentReference,
                        id: id
                    });
                    break;
            }
        });
        return result;
    }
    parseMathElement(elem) {
        const propsTag = `${elem.localName}Pr`;
        const result = { type: mmlTagMap[elem.localName], children: [] };
        for (const el of xml_parser_1.default.elements(elem)) {
            const childType = mmlTagMap[el.localName];
            if (childType) {
                result.children.push(this.parseMathElement(el));
            }
            else if (el.localName == "r") {
                var run = this.parseRun(el);
                run.type = dom_1.DomType.MmlRun;
                result.children.push(run);
            }
            else if (el.localName == propsTag) {
                result.props = this.parseMathProperies(el);
            }
        }
        return result;
    }
    parseMathProperies(elem) {
        const result = {};
        for (const el of xml_parser_1.default.elements(elem)) {
            switch (el.localName) {
                case "chr":
                    result.char = xml_parser_1.default.attr(el, "val");
                    break;
                case "vertJc":
                    result.verticalJustification = xml_parser_1.default.attr(el, "val");
                    break;
                case "pos":
                    result.position = xml_parser_1.default.attr(el, "val");
                    break;
                case "degHide":
                    result.hideDegree = xml_parser_1.default.boolAttr(el, "val");
                    break;
                case "begChr":
                    result.beginChar = xml_parser_1.default.attr(el, "val");
                    break;
                case "endChr":
                    result.endChar = xml_parser_1.default.attr(el, "val");
                    break;
            }
        }
        return result;
    }
    parseRunProperties(elem, run) {
        this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
            switch (c.localName) {
                case "rStyle":
                    run.styleName = xml_parser_1.default.attr(c, "val");
                    break;
                case "vertAlign":
                    run.verticalAlign = values.valueOfVertAlign(c, true);
                    break;
                default:
                    return false;
            }
            return true;
        });
    }
    parseVmlPicture(elem) {
        const result = { type: dom_1.DomType.VmlPicture, children: [] };
        for (const el of xml_parser_1.default.elements(elem)) {
            const child = (0, vml_1.parseVmlElement)(el, this);
            child && result.children.push(child);
        }
        return result;
    }
    checkAlternateContent(elem) {
        if (elem.localName != 'AlternateContent')
            return elem;
        var choice = xml_parser_1.default.element(elem, "Choice");
        if (choice) {
            var requires = xml_parser_1.default.attr(choice, "Requires");
            var namespaceURI = elem.lookupNamespaceURI(requires);
            if (supportedNamespaceURIs.includes(namespaceURI))
                return choice.firstElementChild;
        }
        return xml_parser_1.default.element(elem, "Fallback")?.firstElementChild;
    }
    parseDrawing(node) {
        for (var n of xml_parser_1.default.elements(node)) {
            switch (n.localName) {
                case "inline":
                case "anchor":
                    return this.parseDrawingWrapper(n);
            }
        }
    }
    parseDrawingWrapper(node) {
        var result = { type: dom_1.DomType.Drawing, children: [], cssStyle: {} };
        var isAnchor = node.localName == "anchor";
        let wrapType = null;
        let simplePos = xml_parser_1.default.boolAttr(node, "simplePos");
        let posX = { relative: "page", align: "left", offset: "0" };
        let posY = { relative: "page", align: "top", offset: "0" };
        for (var n of xml_parser_1.default.elements(node)) {
            switch (n.localName) {
                case "simplePos":
                    if (simplePos) {
                        posX.offset = xml_parser_1.default.lengthAttr(n, "x", common_1.LengthUsage.Emu);
                        posY.offset = xml_parser_1.default.lengthAttr(n, "y", common_1.LengthUsage.Emu);
                    }
                    break;
                case "extent":
                    result.cssStyle["width"] = xml_parser_1.default.lengthAttr(n, "cx", common_1.LengthUsage.Emu);
                    result.cssStyle["height"] = xml_parser_1.default.lengthAttr(n, "cy", common_1.LengthUsage.Emu);
                    break;
                case "positionH":
                case "positionV":
                    if (!simplePos) {
                        let pos = n.localName == "positionH" ? posX : posY;
                        var alignNode = xml_parser_1.default.element(n, "align");
                        var offsetNode = xml_parser_1.default.element(n, "posOffset");
                        pos.relative = xml_parser_1.default.attr(n, "relativeFrom") ?? pos.relative;
                        if (alignNode)
                            pos.align = alignNode.textContent;
                        if (offsetNode)
                            pos.offset = xmlUtil.sizeValue(offsetNode, common_1.LengthUsage.Emu);
                    }
                    break;
                case "wrapTopAndBottom":
                    wrapType = "wrapTopAndBottom";
                    break;
                case "wrapNone":
                    wrapType = "wrapNone";
                    break;
                case "graphic":
                    var g = this.parseGraphic(n);
                    if (g)
                        result.children.push(g);
                    break;
            }
        }
        if (wrapType == "wrapTopAndBottom") {
            result.cssStyle['display'] = 'block';
            if (posX.align) {
                result.cssStyle['text-align'] = posX.align;
                result.cssStyle['width'] = "100%";
            }
        }
        else if (wrapType == "wrapNone") {
            result.cssStyle['display'] = 'block';
            result.cssStyle['position'] = 'relative';
            result.cssStyle["width"] = "0px";
            result.cssStyle["height"] = "0px";
            if (posX.offset)
                result.cssStyle["left"] = posX.offset;
            if (posY.offset)
                result.cssStyle["top"] = posY.offset;
        }
        else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
            result.cssStyle["float"] = posX.align;
        }
        return result;
    }
    parseGraphic(elem) {
        var graphicData = xml_parser_1.default.element(elem, "graphicData");
        for (let n of xml_parser_1.default.elements(graphicData)) {
            switch (n.localName) {
                case "pic":
                    return this.parsePicture(n);
            }
        }
        return null;
    }
    parsePicture(elem) {
        var result = { type: dom_1.DomType.Image, src: "", cssStyle: {} };
        var blipFill = xml_parser_1.default.element(elem, "blipFill");
        var blip = xml_parser_1.default.element(blipFill, "blip");
        result.src = xml_parser_1.default.attr(blip, "embed");
        var spPr = xml_parser_1.default.element(elem, "spPr");
        var xfrm = xml_parser_1.default.element(spPr, "xfrm");
        result.cssStyle["position"] = "relative";
        for (var n of xml_parser_1.default.elements(xfrm)) {
            switch (n.localName) {
                case "ext":
                    result.cssStyle["width"] = xml_parser_1.default.lengthAttr(n, "cx", common_1.LengthUsage.Emu);
                    result.cssStyle["height"] = xml_parser_1.default.lengthAttr(n, "cy", common_1.LengthUsage.Emu);
                    break;
                case "off":
                    result.cssStyle["left"] = xml_parser_1.default.lengthAttr(n, "x", common_1.LengthUsage.Emu);
                    result.cssStyle["top"] = xml_parser_1.default.lengthAttr(n, "y", common_1.LengthUsage.Emu);
                    break;
            }
        }
        return result;
    }
    parseTable(node) {
        var result = { type: dom_1.DomType.Table, children: [] };
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "tr":
                    result.children.push(this.parseTableRow(c));
                    break;
                case "tblGrid":
                    result.columns = this.parseTableColumns(c);
                    break;
                case "tblPr":
                    this.parseTableProperties(c, result);
                    break;
            }
        });
        return result;
    }
    parseTableColumns(node) {
        var result = [];
        xmlUtil.foreach(node, n => {
            switch (n.localName) {
                case "gridCol":
                    result.push({ width: xml_parser_1.default.lengthAttr(n, "w") });
                    break;
            }
        });
        return result;
    }
    parseTableProperties(elem, table) {
        table.cssStyle = {};
        table.cellStyle = {};
        this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
            switch (c.localName) {
                case "tblStyle":
                    table.styleName = xml_parser_1.default.attr(c, "val");
                    break;
                case "tblLook":
                    table.className = values.classNameOftblLook(c);
                    break;
                case "tblpPr":
                    this.parseTablePosition(c, table);
                    break;
                case "tblStyleColBandSize":
                    table.colBandSize = xml_parser_1.default.intAttr(c, "val");
                    break;
                case "tblStyleRowBandSize":
                    table.rowBandSize = xml_parser_1.default.intAttr(c, "val");
                    break;
                default:
                    return false;
            }
            return true;
        });
        switch (table.cssStyle["text-align"]) {
            case "center":
                delete table.cssStyle["text-align"];
                table.cssStyle["margin-left"] = "auto";
                table.cssStyle["margin-right"] = "auto";
                break;
            case "right":
                delete table.cssStyle["text-align"];
                table.cssStyle["margin-left"] = "auto";
                break;
        }
    }
    parseTablePosition(node, table) {
        var topFromText = xml_parser_1.default.lengthAttr(node, "topFromText");
        var bottomFromText = xml_parser_1.default.lengthAttr(node, "bottomFromText");
        var rightFromText = xml_parser_1.default.lengthAttr(node, "rightFromText");
        var leftFromText = xml_parser_1.default.lengthAttr(node, "leftFromText");
        table.cssStyle["float"] = 'left';
        table.cssStyle["tblpXSpec"] = xml_parser_1.default.attr(node, "tblpXSpec");
        table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
        table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
        table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
        table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
    }
    parseTableRow(node) {
        var result = { type: dom_1.DomType.Row, children: [] };
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "tc":
                    result.children.push(this.parseTableCell(c));
                    break;
                case "trPr":
                    this.parseTableRowProperties(c, result);
                    break;
            }
        });
        return result;
    }
    parseTableRowProperties(elem, row) {
        row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
            switch (c.localName) {
                case "cnfStyle":
                    row.className = values.classNameOfCnfStyle(c);
                    break;
                case "tblHeader":
                    row.isHeader = xml_parser_1.default.boolAttr(c, "val");
                    break;
                default:
                    return false;
            }
            return true;
        });
    }
    parseTableCell(node) {
        var result = { type: dom_1.DomType.Cell, children: [] };
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "tbl":
                    result.children.push(this.parseTable(c));
                    break;
                case "p":
                    result.children.push(this.parseParagraph(c));
                    break;
                case "tcPr":
                    this.parseTableCellProperties(c, result);
                    break;
            }
        });
        return result;
    }
    parseTableCellProperties(elem, cell) {
        cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
            switch (c.localName) {
                case "gridSpan":
                    cell.span = xml_parser_1.default.intAttr(c, "val", null);
                    break;
                case "vMerge":
                    cell.verticalMerge = xml_parser_1.default.attr(c, "val") ?? "continue";
                    break;
                case "cnfStyle":
                    cell.className = values.classNameOfCnfStyle(c);
                    break;
                default:
                    return false;
            }
            return true;
        });
    }
    parseDefaultProperties(elem, style = null, childStyle = null, handler = null) {
        style = style || {};
        xmlUtil.foreach(elem, c => {
            if (handler?.(c))
                return;
            switch (c.localName) {
                case "jc":
                    style["text-align"] = values.valueOfJc(c);
                    break;
                case "textAlignment":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;
                case "color":
                    style["color"] = xmlUtil.colorAttr(c, "val", null, exports.autos.color);
                    break;
                case "sz":
                    style["font-size"] = style["min-height"] = xml_parser_1.default.lengthAttr(c, "val", common_1.LengthUsage.FontSize);
                    break;
                case "shd":
                    style["background-color"] = xmlUtil.colorAttr(c, "fill", null, exports.autos.shd);
                    break;
                case "highlight":
                    style["background-color"] = xmlUtil.colorAttr(c, "val", null, exports.autos.highlight);
                    break;
                case "vertAlign":
                    break;
                case "position":
                    style.verticalAlign = xml_parser_1.default.lengthAttr(c, "val", common_1.LengthUsage.FontSize);
                    break;
                case "tcW":
                    if (this.options.ignoreWidth)
                        break;
                case "tblW":
                    style["width"] = values.valueOfSize(c, "w");
                    break;
                case "trHeight":
                    this.parseTrHeight(c, style);
                    break;
                case "strike":
                    style["text-decoration"] = xml_parser_1.default.boolAttr(c, "val", true) ? "line-through" : "none";
                    break;
                case "b":
                    style["font-weight"] = xml_parser_1.default.boolAttr(c, "val", true) ? "bold" : "normal";
                    break;
                case "i":
                    style["font-style"] = xml_parser_1.default.boolAttr(c, "val", true) ? "italic" : "normal";
                    break;
                case "caps":
                    style["text-transform"] = xml_parser_1.default.boolAttr(c, "val", true) ? "uppercase" : "none";
                    break;
                case "smallCaps":
                    style["text-transform"] = xml_parser_1.default.boolAttr(c, "val", true) ? "lowercase" : "none";
                    break;
                case "u":
                    this.parseUnderline(c, style);
                    break;
                case "ind":
                case "tblInd":
                    this.parseIndentation(c, style);
                    break;
                case "rFonts":
                    this.parseFont(c, style);
                    break;
                case "tblBorders":
                    this.parseBorderProperties(c, childStyle || style);
                    break;
                case "tblCellSpacing":
                    style["border-spacing"] = values.valueOfMargin(c);
                    style["border-collapse"] = "separate";
                    break;
                case "pBdr":
                    this.parseBorderProperties(c, style);
                    break;
                case "bdr":
                    style["border"] = values.valueOfBorder(c);
                    break;
                case "tcBorders":
                    this.parseBorderProperties(c, style);
                    break;
                case "vanish":
                    if (xml_parser_1.default.boolAttr(c, "val", true))
                        style["display"] = "none";
                    break;
                case "kern":
                    break;
                case "noWrap":
                    break;
                case "tblCellMar":
                case "tcMar":
                    this.parseMarginProperties(c, childStyle || style);
                    break;
                case "tblLayout":
                    style["table-layout"] = values.valueOfTblLayout(c);
                    break;
                case "vAlign":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;
                case "spacing":
                    if (elem.localName == "pPr")
                        this.parseSpacing(c, style);
                    break;
                case "wordWrap":
                    if (xml_parser_1.default.boolAttr(c, "val"))
                        style["overflow-wrap"] = "break-word";
                    break;
                case "suppressAutoHyphens":
                    style["hyphens"] = xml_parser_1.default.boolAttr(c, "val", true) ? "none" : "auto";
                    break;
                case "lang":
                    style["$lang"] = xml_parser_1.default.attr(c, "val");
                    break;
                case "bCs":
                case "iCs":
                case "szCs":
                case "tabs":
                case "outlineLvl":
                case "contextualSpacing":
                case "tblStyleColBandSize":
                case "tblStyleRowBandSize":
                case "webHidden":
                case "pageBreakBefore":
                case "suppressLineNumbers":
                case "keepLines":
                case "keepNext":
                case "widowControl":
                case "bidi":
                case "rtl":
                case "noProof":
                    break;
                default:
                    if (this.options.debug)
                        console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
                    break;
            }
        });
        return style;
    }
    parseUnderline(node, style) {
        var val = xml_parser_1.default.attr(node, "val");
        if (val == null)
            return;
        switch (val) {
            case "dash":
            case "dashDotDotHeavy":
            case "dashDotHeavy":
            case "dashedHeavy":
            case "dashLong":
            case "dashLongHeavy":
            case "dotDash":
            case "dotDotDash":
                style["text-decoration-style"] = "dashed";
                break;
            case "dotted":
            case "dottedHeavy":
                style["text-decoration-style"] = "dotted";
                break;
            case "double":
                style["text-decoration-style"] = "double";
                break;
            case "single":
            case "thick":
                style["text-decoration"] = "underline";
                break;
            case "wave":
            case "wavyDouble":
            case "wavyHeavy":
                style["text-decoration-style"] = "wavy";
                break;
            case "words":
                style["text-decoration"] = "underline";
                break;
            case "none":
                style["text-decoration"] = "none";
                break;
        }
        var col = xmlUtil.colorAttr(node, "color");
        if (col)
            style["text-decoration-color"] = col;
    }
    parseFont(node, style) {
        var ascii = xml_parser_1.default.attr(node, "ascii");
        var asciiTheme = values.themeValue(node, "asciiTheme");
        var fonts = [ascii, asciiTheme].filter(x => x).join(', ');
        if (fonts.length > 0)
            style["font-family"] = fonts;
    }
    parseIndentation(node, style) {
        var firstLine = xml_parser_1.default.lengthAttr(node, "firstLine");
        var hanging = xml_parser_1.default.lengthAttr(node, "hanging");
        var left = xml_parser_1.default.lengthAttr(node, "left");
        var start = xml_parser_1.default.lengthAttr(node, "start");
        var right = xml_parser_1.default.lengthAttr(node, "right");
        var end = xml_parser_1.default.lengthAttr(node, "end");
        if (firstLine)
            style["text-indent"] = firstLine;
        if (hanging)
            style["text-indent"] = `-${hanging}`;
        if (left || start)
            style["margin-left"] = left || start;
        if (right || end)
            style["margin-right"] = right || end;
    }
    parseSpacing(node, style) {
        var before = xml_parser_1.default.lengthAttr(node, "before");
        var after = xml_parser_1.default.lengthAttr(node, "after");
        var line = xml_parser_1.default.intAttr(node, "line", null);
        var lineRule = xml_parser_1.default.attr(node, "lineRule");
        if (before)
            style["margin-top"] = before;
        if (after)
            style["margin-bottom"] = after;
        if (line !== null) {
            switch (lineRule) {
                case "auto":
                    style["line-height"] = `${(line / 240).toFixed(2)}`;
                    break;
                case "atLeast":
                    style["line-height"] = `calc(100% + ${line / 20}pt)`;
                    break;
                default:
                    style["line-height"] = style["min-height"] = `${line / 20}pt`;
                    break;
            }
        }
    }
    parseMarginProperties(node, output) {
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "left":
                    output["padding-left"] = values.valueOfMargin(c);
                    break;
                case "right":
                    output["padding-right"] = values.valueOfMargin(c);
                    break;
                case "top":
                    output["padding-top"] = values.valueOfMargin(c);
                    break;
                case "bottom":
                    output["padding-bottom"] = values.valueOfMargin(c);
                    break;
            }
        });
    }
    parseTrHeight(node, output) {
        switch (xml_parser_1.default.attr(node, "hRule")) {
            case "exact":
                output["height"] = xml_parser_1.default.lengthAttr(node, "val");
                break;
            case "atLeast":
            default:
                output["height"] = xml_parser_1.default.lengthAttr(node, "val");
                break;
        }
    }
    parseBorderProperties(node, output) {
        xmlUtil.foreach(node, c => {
            switch (c.localName) {
                case "start":
                case "left":
                    output["border-left"] = values.valueOfBorder(c);
                    break;
                case "end":
                case "right":
                    output["border-right"] = values.valueOfBorder(c);
                    break;
                case "top":
                    output["border-top"] = values.valueOfBorder(c);
                    break;
                case "bottom":
                    output["border-bottom"] = values.valueOfBorder(c);
                    break;
            }
        });
    }
}
exports.DocumentParser = DocumentParser;
const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];
class xmlUtil {
    static foreach(node, cb) {
        for (var i = 0; i < node.childNodes.length; i++) {
            let n = node.childNodes[i];
            if (n.nodeType == Node.ELEMENT_NODE)
                cb(n);
        }
    }
    static colorAttr(node, attrName, defValue = null, autoColor = 'black') {
        var v = xml_parser_1.default.attr(node, attrName);
        if (v) {
            if (v == "auto") {
                return autoColor;
            }
            else if (knownColors.includes(v)) {
                return v;
            }
            return `#${v}`;
        }
        var themeColor = xml_parser_1.default.attr(node, "themeColor");
        return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
    }
    static sizeValue(node, type = common_1.LengthUsage.Dxa) {
        return (0, common_1.convertLength)(node.textContent, type);
    }
}
class values {
    static themeValue(c, attr) {
        var val = xml_parser_1.default.attr(c, attr);
        return val ? `var(--docx-${val}-font)` : null;
    }
    static valueOfSize(c, attr) {
        var type = common_1.LengthUsage.Dxa;
        switch (xml_parser_1.default.attr(c, "type")) {
            case "dxa": break;
            case "pct":
                type = common_1.LengthUsage.Percent;
                break;
            case "auto": return "auto";
        }
        return xml_parser_1.default.lengthAttr(c, attr, type);
    }
    static valueOfMargin(c) {
        return xml_parser_1.default.lengthAttr(c, "w");
    }
    static valueOfBorder(c) {
        var type = xml_parser_1.default.attr(c, "val");
        if (type == "nil")
            return "none";
        var color = xmlUtil.colorAttr(c, "color");
        var size = xml_parser_1.default.lengthAttr(c, "sz", common_1.LengthUsage.Border);
        return `${size} solid ${color == "auto" ? exports.autos.borderColor : color}`;
    }
    static valueOfTblLayout(c) {
        var type = xml_parser_1.default.attr(c, "val");
        return type == "fixed" ? "fixed" : "auto";
    }
    static classNameOfCnfStyle(c) {
        const val = xml_parser_1.default.attr(c, "val");
        const classes = [
            'first-row', 'last-row', 'first-col', 'last-col',
            'odd-col', 'even-col', 'odd-row', 'even-row',
            'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
        ];
        return classes.filter((_, i) => val[i] == '1').join(' ');
    }
    static valueOfJc(c) {
        var type = xml_parser_1.default.attr(c, "val");
        switch (type) {
            case "start":
            case "left": return "left";
            case "center": return "center";
            case "end":
            case "right": return "right";
            case "both": return "justify";
        }
        return type;
    }
    static valueOfVertAlign(c, asTagName = false) {
        var type = xml_parser_1.default.attr(c, "val");
        switch (type) {
            case "subscript": return "sub";
            case "superscript": return asTagName ? "sup" : "super";
        }
        return asTagName ? null : type;
    }
    static valueOfTextAlignment(c) {
        var type = xml_parser_1.default.attr(c, "val");
        switch (type) {
            case "auto":
            case "baseline": return "baseline";
            case "top": return "top";
            case "center": return "middle";
            case "bottom": return "bottom";
        }
        return type;
    }
    static addSize(a, b) {
        if (a == null)
            return b;
        if (b == null)
            return a;
        return `calc(${a} + ${b})`;
    }
    static classNameOftblLook(c) {
        const val = xml_parser_1.default.hexAttr(c, "val", 0);
        let className = "";
        if (xml_parser_1.default.boolAttr(c, "firstRow") || (val & 0x0020))
            className += " first-row";
        if (xml_parser_1.default.boolAttr(c, "lastRow") || (val & 0x0040))
            className += " last-row";
        if (xml_parser_1.default.boolAttr(c, "firstColumn") || (val & 0x0080))
            className += " first-col";
        if (xml_parser_1.default.boolAttr(c, "lastColumn") || (val & 0x0100))
            className += " last-col";
        if (xml_parser_1.default.boolAttr(c, "noHBand") || (val & 0x0200))
            className += " no-hband";
        if (xml_parser_1.default.boolAttr(c, "noVBand") || (val & 0x0400))
            className += " no-vband";
        return className.trim();
    }
}


/***/ }),

/***/ "./src/document-props/core-props-part.ts":
/*!***********************************************!*\
  !*** ./src/document-props/core-props-part.ts ***!
  \***********************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.CorePropsPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const core_props_1 = __webpack_require__(/*! ./core-props */ "./src/document-props/core-props.ts");
class CorePropsPart extends part_1.Part {
    parseXml(root) {
        this.props = (0, core_props_1.parseCoreProps)(root, this._package.xmlParser);
    }
}
exports.CorePropsPart = CorePropsPart;


/***/ }),

/***/ "./src/document-props/core-props.ts":
/*!******************************************!*\
  !*** ./src/document-props/core-props.ts ***!
  \******************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCoreProps = void 0;
function parseCoreProps(root, xmlParser) {
    const result = {};
    for (let el of xmlParser.elements(root)) {
        switch (el.localName) {
            case "title":
                result.title = el.textContent;
                break;
            case "description":
                result.description = el.textContent;
                break;
            case "subject":
                result.subject = el.textContent;
                break;
            case "creator":
                result.creator = el.textContent;
                break;
            case "keywords":
                result.keywords = el.textContent;
                break;
            case "language":
                result.language = el.textContent;
                break;
            case "lastModifiedBy":
                result.lastModifiedBy = el.textContent;
                break;
            case "revision":
                el.textContent && (result.revision = parseInt(el.textContent));
                break;
        }
    }
    return result;
}
exports.parseCoreProps = parseCoreProps;


/***/ }),

/***/ "./src/document-props/custom-props-part.ts":
/*!*************************************************!*\
  !*** ./src/document-props/custom-props-part.ts ***!
  \*************************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.CustomPropsPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const custom_props_1 = __webpack_require__(/*! ./custom-props */ "./src/document-props/custom-props.ts");
class CustomPropsPart extends part_1.Part {
    parseXml(root) {
        this.props = (0, custom_props_1.parseCustomProps)(root, this._package.xmlParser);
    }
}
exports.CustomPropsPart = CustomPropsPart;


/***/ }),

/***/ "./src/document-props/custom-props.ts":
/*!********************************************!*\
  !*** ./src/document-props/custom-props.ts ***!
  \********************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCustomProps = void 0;
function parseCustomProps(root, xml) {
    return xml.elements(root, "property").map(e => {
        const firstChild = e.firstChild;
        return {
            formatId: xml.attr(e, "fmtid"),
            name: xml.attr(e, "name"),
            type: firstChild.nodeName,
            value: firstChild.textContent
        };
    });
}
exports.parseCustomProps = parseCustomProps;


/***/ }),

/***/ "./src/document-props/extended-props-part.ts":
/*!***************************************************!*\
  !*** ./src/document-props/extended-props-part.ts ***!
  \***************************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ExtendedPropsPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const extended_props_1 = __webpack_require__(/*! ./extended-props */ "./src/document-props/extended-props.ts");
class ExtendedPropsPart extends part_1.Part {
    parseXml(root) {
        this.props = (0, extended_props_1.parseExtendedProps)(root, this._package.xmlParser);
    }
}
exports.ExtendedPropsPart = ExtendedPropsPart;


/***/ }),

/***/ "./src/document-props/extended-props.ts":
/*!**********************************************!*\
  !*** ./src/document-props/extended-props.ts ***!
  \**********************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseExtendedProps = void 0;
function parseExtendedProps(root, xmlParser) {
    const result = {};
    for (let el of xmlParser.elements(root)) {
        switch (el.localName) {
            case "Template":
                result.template = el.textContent;
                break;
            case "Pages":
                result.pages = safeParseToInt(el.textContent);
                break;
            case "Words":
                result.words = safeParseToInt(el.textContent);
                break;
            case "Characters":
                result.characters = safeParseToInt(el.textContent);
                break;
            case "Application":
                result.application = el.textContent;
                break;
            case "Lines":
                result.lines = safeParseToInt(el.textContent);
                break;
            case "Paragraphs":
                result.paragraphs = safeParseToInt(el.textContent);
                break;
            case "Company":
                result.company = el.textContent;
                break;
            case "AppVersion":
                result.appVersion = el.textContent;
                break;
        }
    }
    return result;
}
exports.parseExtendedProps = parseExtendedProps;
function safeParseToInt(value) {
    if (typeof value === 'undefined')
        return;
    return parseInt(value);
}


/***/ }),

/***/ "./src/document/bookmarks.ts":
/*!***********************************!*\
  !*** ./src/document/bookmarks.ts ***!
  \***********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseBookmarkEnd = exports.parseBookmarkStart = void 0;
const dom_1 = __webpack_require__(/*! ./dom */ "./src/document/dom.ts");
function parseBookmarkStart(elem, xml) {
    return {
        type: dom_1.DomType.BookmarkStart,
        id: xml.attr(elem, "id"),
        name: xml.attr(elem, "name"),
        colFirst: xml.intAttr(elem, "colFirst"),
        colLast: xml.intAttr(elem, "colLast"),
        displacedByCustomXml: xml.attr(elem, "displacedByCustomXml"),
    };
}
exports.parseBookmarkStart = parseBookmarkStart;
function parseBookmarkEnd(elem, xml) {
    return {
        type: dom_1.DomType.BookmarkEnd,
        id: xml.attr(elem, "id")
    };
}
exports.parseBookmarkEnd = parseBookmarkEnd;


/***/ }),

/***/ "./src/document/border.ts":
/*!********************************!*\
  !*** ./src/document/border.ts ***!
  \********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseBorders = exports.parseBorder = void 0;
const common_1 = __webpack_require__(/*! ./common */ "./src/document/common.ts");
function parseBorder(elem, xml) {
    return {
        type: xml.attr(elem, "val"),
        color: xml.attr(elem, "color"),
        size: xml.lengthAttr(elem, "sz", common_1.LengthUsage.Border),
        offset: xml.lengthAttr(elem, "space", common_1.LengthUsage.Point),
        frame: xml.boolAttr(elem, 'frame'),
        shadow: xml.boolAttr(elem, 'shadow')
    };
}
exports.parseBorder = parseBorder;
function parseBorders(elem, xml) {
    var result = {};
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "left":
                result.left = parseBorder(e, xml);
                break;
            case "top":
                result.top = parseBorder(e, xml);
                break;
            case "right":
                result.right = parseBorder(e, xml);
                break;
            case "bottom":
                result.bottom = parseBorder(e, xml);
                break;
        }
    }
    return result;
}
exports.parseBorders = parseBorders;


/***/ }),

/***/ "./src/document/comments.ts":
/*!**********************************!*\
  !*** ./src/document/comments.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCommentRangeEnd = exports.parseCommentRangeStart = void 0;
const dom_1 = __webpack_require__(/*! ./dom */ "./src/document/dom.ts");
function parseCommentRangeStart(elem, xml) {
    return {
        type: dom_1.DomType.CommentRangeStart,
        id: xml.attr(elem, "id"),
    };
}
exports.parseCommentRangeStart = parseCommentRangeStart;
function parseCommentRangeEnd(elem, xml) {
    return {
        type: dom_1.DomType.CommentRangeEnd,
        id: xml.attr(elem, "id")
    };
}
exports.parseCommentRangeEnd = parseCommentRangeEnd;


/***/ }),

/***/ "./src/document/common.ts":
/*!********************************!*\
  !*** ./src/document/common.ts ***!
  \********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCommonProperty = exports.convertPercentage = exports.convertBoolean = exports.convertLength = exports.LengthUsage = exports.ns = void 0;
exports.ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    drawingml: "http://schemas.openxmlformats.org/drawingml/2006/main",
    picture: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    compatibility: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    math: "http://schemas.openxmlformats.org/officeDocument/2006/math"
};
exports.LengthUsage = {
    Dxa: { mul: 0.05, unit: "pt" },
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt" },
    Point: { mul: 1, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: "" },
    VmlEmu: { mul: 1 / 12700, unit: "" },
};
function convertLength(val, usage = exports.LengthUsage.Dxa) {
    if (val == null || /.+(p[xt]|[%])$/.test(val)) {
        return val;
    }
    return `${(parseInt(val) * usage.mul).toFixed(2)}${usage.unit}`;
}
exports.convertLength = convertLength;
function convertBoolean(v, defaultValue = false) {
    switch (v) {
        case "1": return true;
        case "0": return false;
        case "on": return true;
        case "off": return false;
        case "true": return true;
        case "false": return false;
        default: return defaultValue;
    }
}
exports.convertBoolean = convertBoolean;
function convertPercentage(val) {
    return val ? parseInt(val) / 100 : null;
}
exports.convertPercentage = convertPercentage;
function parseCommonProperty(elem, props, xml) {
    if (elem.namespaceURI != exports.ns.wordml)
        return false;
    switch (elem.localName) {
        case "color":
            props.color = xml.attr(elem, "val");
            break;
        case "sz":
            props.fontSize = xml.lengthAttr(elem, "val", exports.LengthUsage.FontSize);
            break;
        default:
            return false;
    }
    return true;
}
exports.parseCommonProperty = parseCommonProperty;


/***/ }),

/***/ "./src/document/document-part.ts":
/*!***************************************!*\
  !*** ./src/document/document-part.ts ***!
  \***************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.DocumentPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
class DocumentPart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    parseXml(root) {
        this.body = this._documentParser.parseDocumentFile(root);
    }
}
exports.DocumentPart = DocumentPart;


/***/ }),

/***/ "./src/document/dom.ts":
/*!*****************************!*\
  !*** ./src/document/dom.ts ***!
  \*****************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.OpenXmlElementBase = exports.DomType = void 0;
var DomType;
(function (DomType) {
    DomType["Document"] = "document";
    DomType["Paragraph"] = "paragraph";
    DomType["Run"] = "run";
    DomType["Break"] = "break";
    DomType["NoBreakHyphen"] = "noBreakHyphen";
    DomType["Table"] = "table";
    DomType["Row"] = "row";
    DomType["Cell"] = "cell";
    DomType["Hyperlink"] = "hyperlink";
    DomType["Drawing"] = "drawing";
    DomType["Image"] = "image";
    DomType["Text"] = "text";
    DomType["Tab"] = "tab";
    DomType["Symbol"] = "symbol";
    DomType["BookmarkStart"] = "bookmarkStart";
    DomType["BookmarkEnd"] = "bookmarkEnd";
    DomType["Footer"] = "footer";
    DomType["Header"] = "header";
    DomType["FootnoteReference"] = "footnoteReference";
    DomType["EndnoteReference"] = "endnoteReference";
    DomType["Footnote"] = "footnote";
    DomType["Endnote"] = "endnote";
    DomType["SimpleField"] = "simpleField";
    DomType["ComplexField"] = "complexField";
    DomType["Instruction"] = "instruction";
    DomType["VmlPicture"] = "vmlPicture";
    DomType["MmlMath"] = "mmlMath";
    DomType["MmlMathParagraph"] = "mmlMathParagraph";
    DomType["MmlFraction"] = "mmlFraction";
    DomType["MmlFunction"] = "mmlFunction";
    DomType["MmlFunctionName"] = "mmlFunctionName";
    DomType["MmlNumerator"] = "mmlNumerator";
    DomType["MmlDenominator"] = "mmlDenominator";
    DomType["MmlRadical"] = "mmlRadical";
    DomType["MmlBase"] = "mmlBase";
    DomType["MmlDegree"] = "mmlDegree";
    DomType["MmlSuperscript"] = "mmlSuperscript";
    DomType["MmlSubscript"] = "mmlSubscript";
    DomType["MmlPreSubSuper"] = "mmlPreSubSuper";
    DomType["MmlSubArgument"] = "mmlSubArgument";
    DomType["MmlSuperArgument"] = "mmlSuperArgument";
    DomType["MmlNary"] = "mmlNary";
    DomType["MmlDelimiter"] = "mmlDelimiter";
    DomType["MmlRun"] = "mmlRun";
    DomType["MmlEquationArray"] = "mmlEquationArray";
    DomType["MmlLimit"] = "mmlLimit";
    DomType["MmlLimitLower"] = "mmlLimitLower";
    DomType["MmlMatrix"] = "mmlMatrix";
    DomType["MmlMatrixRow"] = "mmlMatrixRow";
    DomType["MmlBox"] = "mmlBox";
    DomType["MmlBar"] = "mmlBar";
    DomType["MmlGroupChar"] = "mmlGroupChar";
    DomType["VmlElement"] = "vmlElement";
    DomType["Inserted"] = "inserted";
    DomType["Deleted"] = "deleted";
    DomType["DeletedText"] = "deletedText";
    DomType["CommentReference"] = "commentReference";
    DomType["CommentRangeStart"] = "commentRangeStart";
    DomType["CommentRangeEnd"] = "commentRangeEnd";
})(DomType || (exports.DomType = DomType = {}));
class OpenXmlElementBase {
    constructor() {
        this.children = [];
        this.cssStyle = {};
    }
}
exports.OpenXmlElementBase = OpenXmlElementBase;


/***/ }),

/***/ "./src/document/line-spacing.ts":
/*!**************************************!*\
  !*** ./src/document/line-spacing.ts ***!
  \**************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseLineSpacing = void 0;
function parseLineSpacing(elem, xml) {
    return {
        before: xml.lengthAttr(elem, "before"),
        after: xml.lengthAttr(elem, "after"),
        line: xml.intAttr(elem, "line"),
        lineRule: xml.attr(elem, "lineRule")
    };
}
exports.parseLineSpacing = parseLineSpacing;


/***/ }),

/***/ "./src/document/paragraph.ts":
/*!***********************************!*\
  !*** ./src/document/paragraph.ts ***!
  \***********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseNumbering = exports.parseTabs = exports.parseParagraphProperty = exports.parseParagraphProperties = void 0;
const common_1 = __webpack_require__(/*! ./common */ "./src/document/common.ts");
const section_1 = __webpack_require__(/*! ./section */ "./src/document/section.ts");
const line_spacing_1 = __webpack_require__(/*! ./line-spacing */ "./src/document/line-spacing.ts");
const run_1 = __webpack_require__(/*! ./run */ "./src/document/run.ts");
function parseParagraphProperties(elem, xml) {
    let result = {};
    for (let el of xml.elements(elem)) {
        parseParagraphProperty(el, result, xml);
    }
    return result;
}
exports.parseParagraphProperties = parseParagraphProperties;
function parseParagraphProperty(elem, props, xml) {
    if (elem.namespaceURI != common_1.ns.wordml)
        return false;
    if ((0, common_1.parseCommonProperty)(elem, props, xml))
        return true;
    switch (elem.localName) {
        case "tabs":
            props.tabs = parseTabs(elem, xml);
            break;
        case "sectPr":
            props.sectionProps = (0, section_1.parseSectionProperties)(elem, xml);
            break;
        case "numPr":
            props.numbering = parseNumbering(elem, xml);
            break;
        case "spacing":
            props.lineSpacing = (0, line_spacing_1.parseLineSpacing)(elem, xml);
            return false;
            break;
        case "textAlignment":
            props.textAlignment = xml.attr(elem, "val");
            return false;
            break;
        case "keepLines":
            props.keepLines = xml.boolAttr(elem, "val", true);
            break;
        case "keepNext":
            props.keepNext = xml.boolAttr(elem, "val", true);
            break;
        case "pageBreakBefore":
            props.pageBreakBefore = xml.boolAttr(elem, "val", true);
            break;
        case "outlineLvl":
            props.outlineLevel = xml.intAttr(elem, "val");
            break;
        case "pStyle":
            props.styleName = xml.attr(elem, "val");
            break;
        case "rPr":
            props.runProps = (0, run_1.parseRunProperties)(elem, xml);
            break;
        default:
            return false;
    }
    return true;
}
exports.parseParagraphProperty = parseParagraphProperty;
function parseTabs(elem, xml) {
    return xml.elements(elem, "tab")
        .map(e => ({
        position: xml.lengthAttr(e, "pos"),
        leader: xml.attr(e, "leader"),
        style: xml.attr(e, "val")
    }));
}
exports.parseTabs = parseTabs;
function parseNumbering(elem, xml) {
    var result = {};
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "numId":
                result.id = xml.attr(e, "val");
                break;
            case "ilvl":
                result.level = xml.intAttr(e, "val");
                break;
        }
    }
    if (!result.level) {
        result.level = 0;
    }
    return result;
}
exports.parseNumbering = parseNumbering;


/***/ }),

/***/ "./src/document/run.ts":
/*!*****************************!*\
  !*** ./src/document/run.ts ***!
  \*****************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseRunProperty = exports.parseRunProperties = void 0;
const common_1 = __webpack_require__(/*! ./common */ "./src/document/common.ts");
function parseRunProperties(elem, xml) {
    let result = {};
    for (let el of xml.elements(elem)) {
        parseRunProperty(el, result, xml);
    }
    return result;
}
exports.parseRunProperties = parseRunProperties;
function parseRunProperty(elem, props, xml) {
    if ((0, common_1.parseCommonProperty)(elem, props, xml))
        return true;
    return false;
}
exports.parseRunProperty = parseRunProperty;


/***/ }),

/***/ "./src/document/section.ts":
/*!*********************************!*\
  !*** ./src/document/section.ts ***!
  \*********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseSectionProperties = exports.SectionType = void 0;
const xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
const border_1 = __webpack_require__(/*! ./border */ "./src/document/border.ts");
var SectionType;
(function (SectionType) {
    SectionType["Continuous"] = "continuous";
    SectionType["NextPage"] = "nextPage";
    SectionType["NextColumn"] = "nextColumn";
    SectionType["EvenPage"] = "evenPage";
    SectionType["OddPage"] = "oddPage";
})(SectionType || (exports.SectionType = SectionType = {}));
function parseSectionProperties(elem, xml = xml_parser_1.default) {
    var section = {};
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "pgSz":
                section.pageSize = {
                    width: xml.lengthAttr(e, "w"),
                    height: xml.lengthAttr(e, "h"),
                    orientation: xml.attr(e, "orient")
                };
                break;
            case "type":
                section.type = xml.attr(e, "val");
                break;
            case "pgMar":
                section.pageMargins = {
                    left: xml.lengthAttr(e, "left"),
                    right: xml.lengthAttr(e, "right"),
                    top: xml.lengthAttr(e, "top"),
                    bottom: xml.lengthAttr(e, "bottom"),
                    header: xml.lengthAttr(e, "header"),
                    footer: xml.lengthAttr(e, "footer"),
                    gutter: xml.lengthAttr(e, "gutter"),
                };
                break;
            case "cols":
                section.columns = parseColumns(e, xml);
                break;
            case "headerReference":
                (section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
                break;
            case "footerReference":
                (section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
                break;
            case "titlePg":
                section.titlePage = xml.boolAttr(e, "val", true);
                break;
            case "pgBorders":
                section.pageBorders = (0, border_1.parseBorders)(e, xml);
                break;
            case "pgNumType":
                section.pageNumber = parsePageNumber(e, xml);
                break;
        }
    }
    return section;
}
exports.parseSectionProperties = parseSectionProperties;
function parseColumns(elem, xml) {
    return {
        numberOfColumns: xml.intAttr(elem, "num"),
        space: xml.lengthAttr(elem, "space"),
        separator: xml.boolAttr(elem, "sep"),
        equalWidth: xml.boolAttr(elem, "equalWidth", true),
        columns: xml.elements(elem, "col")
            .map(e => ({
            width: xml.lengthAttr(e, "w"),
            space: xml.lengthAttr(e, "space")
        }))
    };
}
function parsePageNumber(elem, xml) {
    return {
        chapSep: xml.attr(elem, "chapSep"),
        chapStyle: xml.attr(elem, "chapStyle"),
        format: xml.attr(elem, "fmt"),
        start: xml.intAttr(elem, "start")
    };
}
function parseFooterHeaderReference(elem, xml) {
    return {
        id: xml.attr(elem, "id"),
        type: xml.attr(elem, "type"),
    };
}


/***/ }),

/***/ "./src/font-table/font-table.ts":
/*!**************************************!*\
  !*** ./src/font-table/font-table.ts ***!
  \**************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.FontTablePart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const fonts_1 = __webpack_require__(/*! ./fonts */ "./src/font-table/fonts.ts");
class FontTablePart extends part_1.Part {
    parseXml(root) {
        this.fonts = (0, fonts_1.parseFonts)(root, this._package.xmlParser);
    }
}
exports.FontTablePart = FontTablePart;


/***/ }),

/***/ "./src/font-table/fonts.ts":
/*!*********************************!*\
  !*** ./src/font-table/fonts.ts ***!
  \*********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseEmbedFontRef = exports.parseFont = exports.parseFonts = void 0;
const embedFontTypeMap = {
    embedRegular: 'regular',
    embedBold: 'bold',
    embedItalic: 'italic',
    embedBoldItalic: 'boldItalic',
};
function parseFonts(root, xml) {
    return xml.elements(root).map(el => parseFont(el, xml));
}
exports.parseFonts = parseFonts;
function parseFont(elem, xml) {
    let result = {
        name: xml.attr(elem, "name"),
        embedFontRefs: []
    };
    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "family":
                result.family = xml.attr(el, "val");
                break;
            case "altName":
                result.altName = xml.attr(el, "val");
                break;
            case "embedRegular":
            case "embedBold":
            case "embedItalic":
            case "embedBoldItalic":
                result.embedFontRefs.push(parseEmbedFontRef(el, xml));
                break;
        }
    }
    return result;
}
exports.parseFont = parseFont;
function parseEmbedFontRef(elem, xml) {
    return {
        id: xml.attr(elem, "id"),
        key: xml.attr(elem, "fontKey"),
        type: embedFontTypeMap[elem.localName]
    };
}
exports.parseEmbedFontRef = parseEmbedFontRef;


/***/ }),

/***/ "./src/header-footer/elements.ts":
/*!***************************************!*\
  !*** ./src/header-footer/elements.ts ***!
  \***************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.WmlFooter = exports.WmlHeader = void 0;
const dom_1 = __webpack_require__(/*! ../document/dom */ "./src/document/dom.ts");
class WmlHeader extends dom_1.OpenXmlElementBase {
    constructor() {
        super(...arguments);
        this.type = dom_1.DomType.Header;
    }
}
exports.WmlHeader = WmlHeader;
class WmlFooter extends dom_1.OpenXmlElementBase {
    constructor() {
        super(...arguments);
        this.type = dom_1.DomType.Footer;
    }
}
exports.WmlFooter = WmlFooter;


/***/ }),

/***/ "./src/header-footer/parts.ts":
/*!************************************!*\
  !*** ./src/header-footer/parts.ts ***!
  \************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.FooterPart = exports.HeaderPart = exports.BaseHeaderFooterPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const elements_1 = __webpack_require__(/*! ./elements */ "./src/header-footer/elements.ts");
class BaseHeaderFooterPart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    parseXml(root) {
        this.rootElement = this.createRootElement();
        this.rootElement.children = this._documentParser.parseBodyElements(root);
    }
}
exports.BaseHeaderFooterPart = BaseHeaderFooterPart;
class HeaderPart extends BaseHeaderFooterPart {
    createRootElement() {
        return new elements_1.WmlHeader();
    }
}
exports.HeaderPart = HeaderPart;
class FooterPart extends BaseHeaderFooterPart {
    createRootElement() {
        return new elements_1.WmlFooter();
    }
}
exports.FooterPart = FooterPart;


/***/ }),

/***/ "./src/html-renderer.ts":
/*!******************************!*\
  !*** ./src/html-renderer.ts ***!
  \******************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.HtmlRenderer = void 0;
const dom_1 = __webpack_require__(/*! ./document/dom */ "./src/document/dom.ts");
const utils_1 = __webpack_require__(/*! ./utils */ "./src/utils.ts");
const javascript_1 = __webpack_require__(/*! ./javascript */ "./src/javascript.ts");
const index_1 = __webpack_require__(/*! ./assets/index */ "./src/assets/index.ts");
const ns = {
    svg: "http://www.w3.org/2000/svg",
    mathML: "http://www.w3.org/1998/Math/MathML"
};
class HtmlRenderer {
    constructor(htmlDocument) {
        this.htmlDocument = htmlDocument;
        this.className = "docx";
        this.styleMap = {};
        this.currentPart = null;
        this.tableVerticalMerges = [];
        this.currentVerticalMerge = null;
        this.tableCellPositions = [];
        this.currentCellPosition = null;
        this.footnoteMap = {};
        this.endnoteMap = {};
        this.currentEndnoteIds = [];
        this.usedHederFooterParts = [];
        this.currentTabs = [];
        this.tabsTimeout = 0;
        this.currentComment = '';
        this.allComments = [];
        this.createElement = createElement;
    }
    processCommentPart(document, id, commentReference) {
        let comments = document.commentsPart.comments;
        let commentsEx = document.commentsExtendedPart.commentsEx;
        let result = Object.assign(comments.find(item => {
            return item.id === id;
        }), commentReference);
        result.children = [];
        let children = commentsEx.filter(item => {
            return item.paraIdParent === result?.paraId;
        });
        for (let child of children) {
            let comment = comments.find(item => {
                return item.paraId === child.paraId;
            });
            result.children.push(comment);
        }
        if (commentsEx.find(item => { return item.paraId === result.paraId; })?.paraIdParent) {
            result.noRender = true;
        }
        return result;
    }
    joinRangeText(rArr) {
        let result = '';
        for (let i = 0; i < rArr.length; i++) {
            let rChildren = rArr[i]?.children || [];
            for (let child of rChildren) {
                if (child && child?.type === dom_1.DomType.Text) {
                    result += child.text;
                }
            }
        }
        return result;
    }
    flatChildren(document, flatedDocument) {
        let body = document.documentPart.body;
        function myFlat(elem, flatedDocument) {
            flatedDocument.push(elem);
            if (elem.children && elem.children.length > 0) {
                for (let child of elem.children) {
                    myFlat(child, flatedDocument);
                }
            }
        }
        myFlat(body, flatedDocument);
    }
    render(document, bodyContainer, styleContainer = null, options) {
        this.document = document;
        this.options = options;
        this.className = options.className;
        this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
        this.styleMap = null;
        this.flatedDocument = [];
        this.commentIndex = 1;
        this.commentTranslateY = 0;
        styleContainer = styleContainer || bodyContainer;
        this.flatChildren(document, this.flatedDocument);
        removeAllElements(styleContainer);
        removeAllElements(bodyContainer);
        this.processAllComments();
        appendComment(styleContainer, "docxjs library predefined styles");
        styleContainer.appendChild(this.renderDefaultStyle());
        if (document.themePart) {
            appendComment(styleContainer, "docxjs document theme values");
            this.renderTheme(document.themePart, styleContainer);
        }
        if (document.stylesPart != null) {
            this.styleMap = this.processStyles(document.stylesPart.styles);
            appendComment(styleContainer, "docxjs document styles");
            styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
        }
        if (document.numberingPart) {
            this.prodessNumberings(document.numberingPart.domNumberings);
            appendComment(styleContainer, "docxjs document numbering styles");
            styleContainer.appendChild(this.renderTitleNumbering(document.documentPart.body.children, document.numberingPart.domNumberings, styleContainer));
        }
        if (document.footnotesPart) {
            this.footnoteMap = (0, utils_1.keyBy)(document.footnotesPart.notes, x => x.id);
        }
        if (document.endnotesPart) {
            this.endnoteMap = (0, utils_1.keyBy)(document.endnotesPart.notes, x => x.id);
        }
        if (document.settingsPart) {
            this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
        }
        if (!options.ignoreFonts && document.fontTablePart)
            this.renderFontTable(document.fontTablePart, styleContainer);
        var sectionElements = this.renderSections(document.documentPart.body);
        if (this.options.inWrapper) {
            bodyContainer.appendChild(this.renderOutlineWrapper(this.renderWrapper(sectionElements)));
        }
        else {
            appendChildren(bodyContainer, sectionElements);
        }
        this.refreshTabStops();
        (this.options.inWrapper && this.options.renderComments) && this.renderCommentElement();
    }
    processAllComments() {
        let commentsEx = this.document.commentsExtendedPart?.commentsEx;
        let comments = this.document.commentsPart?.comments;
        if (comments && commentsEx) {
            for (let commentEx of commentsEx) {
                if (commentEx.paraIdParent === null) {
                    this.allComments.push(commentEx.paraId);
                }
            }
        }
    }
    renderCommentElement() {
        if (this.document.commentsPart?.comments) {
            let wrap = document.getElementsByClassName(`${this.className}-comment-wrap`)[0];
            for (let paraId of this.allComments) {
                let commentSupElem = document.getElementById(`${this.className}-comment-start-${paraId}`);
                let rect = commentSupElem?.getBoundingClientRect();
                let commentElem = this.createElement("div");
                commentElem.id = `${this.className}-comment-element-${paraId}`;
                commentElem.className = `${this.className}-comment-content-wrap`;
                let that = this;
                commentElem.addEventListener("mouseenter", e => {
                    for (let paraId of that.allComments) {
                        that.inactiveCommentElement(paraId);
                    }
                    that.activeCommentElement(paraId);
                    that.currentComment = paraId;
                    that.setCurrentCommentText();
                });
                commentElem.addEventListener("mouseleave", e => {
                    for (let paraId of that.allComments) {
                        that.inactiveCommentElement(paraId);
                    }
                });
                let commentContent = this.renderCommentContent(paraId);
                if (commentContent !== null) {
                    commentElem.appendChild(commentContent);
                }
                wrap.appendChild(commentElem);
                if (rect) {
                    if (rect.y > this.commentTranslateY) {
                        commentElem.style.top = `${rect.y}px`;
                    }
                    else {
                        commentElem.style.top = `${this.commentTranslateY}px`;
                    }
                    let elemRect = commentElem.getBoundingClientRect();
                    this.commentTranslateY = elemRect.height + elemRect.y + 10;
                }
            }
        }
    }
    activeCommentElement(paraId) {
        let commentElem = document.getElementById(`${this.className}-comment-element-${paraId}`);
        let commentStart = document.getElementById(`${this.className}-comment-start-${paraId}`);
        let commentEnd = document.getElementById(`${this.className}-comment-end-${paraId}`);
        if (commentStart) {
            commentStart.style.borderRight = "4px solid rgb(252, 13, 27)";
            commentStart.style.marginRight = "3px";
        }
        if (commentEnd) {
            commentEnd.style.borderLeft = "4px solid rgb(252, 13, 27)";
            commentEnd.style.marginLeft = "3px";
            commentEnd.classList.add("active");
        }
        if (commentElem) {
            commentElem.style.borderLeft = "5px solid rgb(252, 13, 27)";
            commentElem.style.boxShadow = "3px 3px 6px #ccc";
            commentElem.style.left = "-20px";
            commentElem.classList.add("active");
        }
    }
    inactiveCommentElement(paraId) {
        let commentElem = document.getElementById(`${this.className}-comment-element-${paraId}`);
        let commentStart = document.getElementById(`${this.className}-comment-start-${paraId}`);
        let commentEnd = document.getElementById(`${this.className}-comment-end-${paraId}`);
        if (commentStart) {
            commentStart.style.borderRight = "none";
            commentStart.style.marginRight = "0";
        }
        if (commentEnd) {
            commentEnd.style.borderLeft = "none";
            commentEnd.style.marginLeft = "0";
            commentEnd.classList.remove("active");
        }
        if (commentElem) {
            commentElem.style.borderLeft = "none";
            commentElem.style.boxShadow = "none";
            commentElem.style.left = "0";
            commentElem.classList.remove("active");
        }
    }
    renderCommentContent(paraId) {
        let comments = this.document.commentsPart?.comments || [];
        let commentsEx = this.document.commentsExtendedPart?.commentsEx || [];
        let comment = comments.find(item => {
            return item.paraId === paraId;
        });
        let id = comment?.id;
        let msg = '';
        let commentRangeStartIndex = this.flatedDocument.findIndex(item => {
            return item.type === "commentRangeStart" && item.id === id;
        });
        for (let i = commentRangeStartIndex; i < this.flatedDocument.length; i++) {
            if (this.flatedDocument[i].type === "commentRangeEnd" && this.flatedDocument[i].id === id) {
                break;
            }
            if (this.flatedDocument[i] && this.flatedDocument[i].type === dom_1.DomType.Text) {
                msg = msg + this.flatedDocument[i].text;
            }
        }
        if (comment) {
            comment.msg = msg;
            for (let commentEx of commentsEx) {
                if (commentEx.paraIdParent === paraId) {
                    let child = comments.find(item => {
                        return item.paraId === commentEx.paraId;
                    });
                    if (child) {
                        comment.children.push(child);
                    }
                }
            }
        }
        return this.createCommentContentNode(comment);
    }
    createCommentContentNode(elem) {
        let commentsContainer = this.createElement("div");
        let commentContainer = this.createCommentNode(elem);
        if (commentContainer) {
            commentsContainer.appendChild(commentContainer);
        }
        return commentsContainer;
    }
    createCommentNode(elem) {
        if (!elem) {
            return null;
        }
        let commentContainer = this.createElement("div");
        let author = this.createElement("span");
        author.style.marginRight = "10px";
        author.style.fontSize = "13px";
        author.textContent = elem.author;
        let date = this.createElement("span");
        date.style.fontSize = "13px";
        date.textContent = elem.date?.replace('T', ' ')?.replace('Z', '');
        let text = this.createElement("div");
        text.style.width = "270px";
        text.style.margin = "3px 0 5px";
        text.style.fontSize = "14px";
        text.textContent = elem.text;
        commentContainer.appendChild(author);
        commentContainer.appendChild(date);
        commentContainer.appendChild(text);
        if (elem?.children?.length > 0) {
            let childCommentContainer = this.createElement("div");
            childCommentContainer.style.borderLeft = "2px solid #bbbfc4";
            childCommentContainer.style.paddingLeft = "7px";
            childCommentContainer.style.marginLeft = "20px";
            for (let i = 0; i < elem?.children?.length; i++) {
                childCommentContainer.appendChild(this.createCommentNode(elem?.children?.[i]));
            }
            commentContainer.appendChild(childCommentContainer);
        }
        return commentContainer;
    }
    renderTheme(themePart, styleContainer) {
        const variables = {};
        const fontScheme = themePart.theme?.fontScheme;
        if (fontScheme) {
            if (fontScheme.majorFont) {
                variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
            }
            if (fontScheme.minorFont) {
                variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
            }
        }
        const colorScheme = themePart.theme?.colorScheme;
        if (colorScheme) {
            for (let [k, v] of Object.entries(colorScheme.colors)) {
                variables[`--docx-${k}-color`] = `#${v}`;
            }
        }
        const cssText = this.styleToString(`.${this.className}`, variables);
        styleContainer.appendChild(createStyleElement(cssText));
    }
    renderFontTable(fontsPart, styleContainer) {
        for (let f of fontsPart.fonts) {
            for (let ref of f.embedFontRefs) {
                this.document.loadFont(ref.id, ref.key).then(fontData => {
                    const cssValues = {
                        'font-family': f.name,
                        'src': `url(${fontData})`
                    };
                    if (ref.type == "bold" || ref.type == "boldItalic") {
                        cssValues['font-weight'] = 'bold';
                    }
                    if (ref.type == "italic" || ref.type == "boldItalic") {
                        cssValues['font-style'] = 'italic';
                    }
                    appendComment(styleContainer, `docxjs ${f.name} font`);
                    const cssText = this.styleToString("@font-face", cssValues);
                    styleContainer.appendChild(createStyleElement(cssText));
                    this.refreshTabStops();
                });
            }
        }
    }
    processStyleName(className) {
        return className ? `${this.className}_${(0, utils_1.escapeClassName)(className)}` : this.className;
    }
    processStyles(styles) {
        const stylesMap = (0, utils_1.keyBy)(styles.filter(x => x.id != null), x => x.id);
        for (const style of styles.filter(x => x.basedOn)) {
            var baseStyle = stylesMap[style.basedOn];
            if (baseStyle) {
                style.paragraphProps = (0, utils_1.mergeDeep)(style.paragraphProps, baseStyle.paragraphProps);
                style.runProps = (0, utils_1.mergeDeep)(style.runProps, baseStyle.runProps);
                for (const baseValues of baseStyle.styles) {
                    const styleValues = style.styles.find(x => x.target == baseValues.target);
                    if (styleValues) {
                        this.copyStyleProperties(baseValues.values, styleValues.values);
                    }
                    else {
                        style.styles.push({ ...baseValues, values: { ...baseValues.values } });
                    }
                }
            }
            else if (this.options.debug)
                console.warn(`Can't find base style ${style.basedOn}`);
        }
        for (let style of styles) {
            style.cssName = this.processStyleName(style.id);
        }
        return stylesMap;
    }
    prodessNumberings(numberings) {
        for (let num of numberings.filter(n => n.pStyleName)) {
            const style = this.findStyle(num.pStyleName);
            if (style?.paragraphProps?.numbering) {
                style.paragraphProps.numbering.level = num.level;
            }
        }
    }
    processElement(element) {
        if (element.children) {
            for (var e of element.children) {
                e.parent = element;
                if (e.type == dom_1.DomType.Table) {
                    this.processTable(e);
                }
                else {
                    this.processElement(e);
                }
            }
        }
    }
    processTable(table) {
        for (var r of table.children) {
            for (var c of r.children) {
                c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
                    "border-left", "border-right", "border-top", "border-bottom",
                    "padding-left", "padding-right", "padding-top", "padding-bottom"
                ]);
                this.processElement(c);
            }
        }
    }
    copyStyleProperties(input, output, attrs = null) {
        if (!input)
            return output;
        if (output == null)
            output = {};
        if (attrs == null)
            attrs = Object.getOwnPropertyNames(input);
        for (var key of attrs) {
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                output[key] = input[key];
        }
        return output;
    }
    createSection(className, props) {
        var elem = this.createElement("section", { className });
        if (props) {
            if (props.pageMargins) {
                elem.style.paddingLeft = props.pageMargins.left;
                elem.style.paddingRight = props.pageMargins.right;
                elem.style.paddingTop = props.pageMargins.top;
                elem.style.paddingBottom = props.pageMargins.bottom;
            }
            if (props.pageSize) {
                if (!this.options.ignoreWidth)
                    elem.style.width = props.pageSize.width;
                if (!this.options.ignoreHeight)
                    elem.style.minHeight = props.pageSize.height;
            }
            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = props.columns.space;
                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }
        return elem;
    }
    renderSections(document) {
        const result = [];
        this.processElement(document);
        const sections = this.splitBySection(document.children);
        let prevProps = null;
        for (let i = 0, l = sections.length; i < l; i++) {
            this.currentFootnoteIds = [];
            const section = sections[i];
            const props = section.sectProps || document.props;
            const sectionElement = this.createSection(this.className, props);
            this.renderStyleValues(document.cssStyle, sectionElement);
            this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props, result.length, prevProps != props, sectionElement);
            var contentElement = this.createElement("article");
            this.renderElements(section.elements, contentElement);
            sectionElement.appendChild(contentElement);
            if (this.options.renderFootnotes) {
                this.renderNotes(this.currentFootnoteIds, this.footnoteMap, sectionElement);
            }
            if (this.options.renderEndnotes && i == l - 1) {
                this.renderNotes(this.currentEndnoteIds, this.endnoteMap, sectionElement);
            }
            this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props, result.length, prevProps != props, sectionElement);
            result.push(sectionElement);
            prevProps = props;
        }
        return result;
    }
    renderHeaderFooter(refs, props, page, firstOfSection, into) {
        if (!refs)
            return;
        var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
            ?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
            ?? refs.find(x => x.type == "default");
        var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart);
        if (part) {
            this.currentPart = part;
            if (!this.usedHederFooterParts.includes(part.path)) {
                this.processElement(part.rootElement);
                this.usedHederFooterParts.push(part.path);
            }
            this.renderElements([part.rootElement], into);
            this.currentPart = null;
        }
    }
    isPageBreakElement(elem) {
        if (elem.type != dom_1.DomType.Break)
            return false;
        if (elem.break == "lastRenderedPageBreak")
            return !this.options.ignoreLastRenderedPageBreak;
        return elem.break == "page";
    }
    splitBySection(elements) {
        var current = { sectProps: null, elements: [] };
        var result = [current];
        for (let elem of elements) {
            if (elem.type == dom_1.DomType.Paragraph) {
                const s = this.findStyle(elem.styleName);
                if (s?.paragraphProps?.pageBreakBefore) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
            }
            current.elements.push(elem);
            if (elem.type == dom_1.DomType.Paragraph) {
                const p = elem;
                var sectProps = p.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                if (this.options.breakPages && p.children) {
                    pBreakIndex = p.children.findIndex(r => {
                        rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
                        return rBreakIndex != -1;
                    });
                }
                if (sectProps || pBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
                if (pBreakIndex != -1 && pBreakIndex != 0) {
                    let breakRun = p.children[pBreakIndex];
                    let splitRun = false;
                    if (pBreakIndex < p.children.length - 1 || splitRun) {
                        var children = elem.children;
                        var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
                        elem.children = children.slice(0, pBreakIndex);
                        newParagraph.noRenderNumbering = true;
                        current.elements.push(newParagraph);
                        if (splitRun) {
                            let runChildren = breakRun.children;
                            let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
                            elem.children.push(newRun);
                            breakRun.children = runChildren.slice(rBreakIndex);
                            console.log(breakRun);
                        }
                    }
                }
            }
        }
        let currentSectProps = null;
        for (let i = result.length - 1; i >= 0; i--) {
            if (result[i].sectProps == null) {
                result[i].sectProps = currentSectProps;
            }
            else {
                currentSectProps = result[i].sectProps;
            }
        }
        return result;
    }
    renderWrapper(children) {
        return this.createElement("div", { className: `${this.className}-wrapper` }, children);
    }
    renderOutlineWrapper(child) {
        let outline = this.createOutlineElement();
        let documentContainer = this.createElement("div", { className: `${this.className}-document-container-wrapper` });
        documentContainer.appendChild(child);
        let documentCommentContainer = this.createElement("div", { className: `${this.className}-document-comment-container-wrapper` });
        documentCommentContainer.appendChild(documentContainer);
        let collapseBtn = this.createElement("div", { className: `${this.className}-collapse-btn` });
        collapseBtn.appendChild((0, index_1.createIconCollapse)());
        collapseBtn.addEventListener("click", function () {
            if (!outline.className.includes("close")) {
                outline.classList.add("close");
                collapseBtn.classList.add("close");
            }
            else {
                outline.classList.remove("close");
                collapseBtn.classList.remove("close");
            }
        });
        let collapseBtnWrapper = this.createElement("div", { className: `${this.className}-collapse-btn-wrapper` });
        collapseBtnWrapper.appendChild(collapseBtn);
        let commentControler = this.createElement("div", { className: `${this.className}-comment-controler` });
        let currentCommentElem = this.createElement("span", { className: `${this.className}-current-comment` });
        currentCommentElem.id = `${this.className}-comment-current-index`;
        this.setCurrentCommentText(currentCommentElem);
        let allCommentElem = this.createElement("span", { className: `${this.className}-all-comment` });
        this.setAllCommentText(allCommentElem);
        let that = this;
        let prevComment = this.createElement("span", { className: `${this.className}-prev-comment` });
        prevComment.appendChild((0, index_1.createIconPrev)());
        prevComment.title = "ctrl + ↑";
        prevComment.addEventListener("click", function (e) {
            e.stopPropagation();
            that.handleClickPrevComment();
        });
        document.addEventListener("keydown", function (e) {
            if (e.keyCode === 38 && (navigator.platform.match("Mac") ? e.metaKey : e.ctrlKey)) {
                e.preventDefault();
                that.handleClickPrevComment();
            }
        });
        let nextComment = this.createElement("span", { className: `${this.className}-next-comment` });
        nextComment.appendChild((0, index_1.createIconPrev)());
        nextComment.title = "ctrl + ↓";
        nextComment.addEventListener("click", function (e) {
            e.stopPropagation();
            that.handleClickNextComment();
        });
        document.addEventListener("keydown", function (e) {
            if (e.keyCode === 40 && (navigator.platform.match("Mac") ? e.metaKey : e.ctrlKey)) {
                e.preventDefault();
                that.handleClickNextComment();
            }
        });
        commentControler.appendChild(currentCommentElem);
        commentControler.appendChild(allCommentElem);
        commentControler.appendChild(prevComment);
        commentControler.appendChild(nextComment);
        let commentWrap = this.createElement("div", { className: `${this.className}-comment-wrap` });
        commentWrap.appendChild(commentControler);
        let renderers = [documentCommentContainer];
        if (this.options.renderOutline) {
            renderers.unshift(collapseBtnWrapper);
            renderers.unshift(outline);
        }
        if (this.options.renderComments) {
            documentCommentContainer.appendChild(commentWrap);
        }
        let outlineWrapper = this.createElement("div", { className: `${this.className}-outline-wrapper` }, renderers);
        return outlineWrapper;
    }
    handleClickPrevComment() {
        let index = this.allComments.findIndex(item => { return item === this.currentComment; });
        if (index > 0) {
            this.currentComment = this.allComments[index - 1];
        }
        else if (index === 0) {
            this.currentComment = this.allComments[this.allComments.length - 1];
        }
        else {
            return;
        }
        let wrap = document.getElementsByClassName(`${this.className}-comment-wrap`)[0];
        let anchor = wrap.querySelector(`#${this.className}-comment-element-${this.currentComment}`);
        anchor?.scrollIntoView({ behavior: "auto", block: "center", inline: "nearest" });
        anchor.click();
        this.setCurrentCommentText();
        for (let paraId of this.allComments) {
            this.inactiveCommentElement(paraId);
        }
        this.activeCommentElement(this.currentComment);
    }
    handleClickNextComment() {
        let index = this.allComments.findIndex(item => { return item === this.currentComment; });
        if (index < this.allComments.length - 1 && this.allComments.length > 0) {
            this.currentComment = this.allComments[index + 1];
        }
        else if (index === this.allComments.length - 1 && this.allComments.length > 0) {
            this.currentComment = this.allComments[0];
        }
        else {
            return;
        }
        let wrap = document.getElementsByClassName(`${this.className}-comment-wrap`)[0];
        let anchor = wrap.querySelector(`#${this.className}-comment-element-${this.currentComment}`);
        anchor?.scrollIntoView({ behavior: "auto", block: "center", inline: "nearest" });
        anchor.click();
        this.setCurrentCommentText();
        for (let paraId of this.allComments) {
            this.inactiveCommentElement(paraId);
        }
        this.activeCommentElement(this.currentComment);
    }
    setCurrentCommentText(currentElem) {
        let elem = document.getElementById(`${this.className}-comment-current-index`);
        let index = this.allComments.findIndex(item => { return item === this.currentComment; });
        if (currentElem) {
            currentElem.textContent = `评论（${index + 1}/`;
        }
        else {
            if (elem) {
                elem.textContent = `评论（${index + 1}/`;
            }
        }
    }
    setAllCommentText(elem) {
        elem.textContent = `${this.allComments.length}）`;
    }
    createOutlineElement() {
        let outlineContainer = this.createElement("div", { className: `${this.className}-outline-container` });
        let outlineContent = this.createElement("div", { className: `${this.className}-outline-content` });
        let pArr = this.document.documentPart.body.children;
        for (let p of pArr) {
            if (p.type === dom_1.DomType.Paragraph) {
                let pChildren = p.children ?? [];
                for (let i = 0; i < pChildren.length; i++) {
                    let child = pChildren[i];
                    if (child.type === dom_1.DomType.BookmarkStart && child.name?.startsWith("_Toc") && !child.displacedByCustomXml) {
                        let endIndex = pChildren.findIndex(item => {
                            return item.type === dom_1.DomType.BookmarkEnd && item.id === child.id;
                        });
                        if (endIndex > -1) {
                            let title = this.joinRangeText(pChildren);
                            if (title.length > 30) {
                                break;
                            }
                            let result = this.createElement("p");
                            result.style.marginLeft = `${16}px`;
                            result.style.fontSize = '14px';
                            result.style.lineHeight = '20px';
                            result.textContent = title;
                            let that = this;
                            result.addEventListener("click", function () {
                                let wrap = document.getElementsByClassName(`${that.className}-wrapper`)[0];
                                let anchor = wrap.querySelector(`#${child.name}`);
                                anchor?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
                            });
                            outlineContent.appendChild(result);
                            break;
                        }
                    }
                }
            }
        }
        outlineContainer.appendChild(outlineContent);
        return outlineContainer;
    }
    renderDefaultStyle() {
        var c = this.className;
        var styleText = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; flex: 1; } 
.${c}-outline-wrapper { display: flex; height: 100%; }
.${c}-document-comment-container-wrapper { flex-grow: 1; height: 100%; overflow: auto; display: flex; background: gray; counter-reset: ${c}-comment-range-end-counter 0; }
.${c}-document-container-wrapper { flex-grow: 1; }
.${c}-collapse-comment-btn-wrapper { position: sticky; top: 0; }
.${c}-comment-wrap { position: relative; width: 340px; counter-reset: ${c}-comment-counter 0;}
.${c}-comment-range-end { counter-increment: ${c}-comment-range-end-counter; }
.${c}-comment-range-end.active::after { content: ""counter(${c}-comment-range-end-counter, decimal)""; color: rgb(252, 13, 27); font-size: 16px; font-weight: 600; margin-left: 3px;}
.${c}-comment-content-wrap { counter-increment: ${c}-comment-counter; position: absolute; left: 0; width: 300px; padding: 15px; margin-right: 5px; z-index: 998; background-color: #fff; border-radius: 5px; text-align: start; cursor: pointer;}
.${c}-comment-content-wrap.active::before {content: ""counter(${c}-comment-counter, decimal)""; position: absolute; left: -26px; top: -2px; color: #fff; font-weight: 600; font-size: 20px;}
.${c}-comment-controler { position: fixed; top: 5px; right: 25px; background: #fff; min-width: 200px; 
	height: 35px; display: flex; align-items: center; border-radius: 4px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); z-index: 999; }
.${c}-outline-container { background: white; width: 300px; height: 100%; overflow: auto; transition: width .5s ease; }
.${c}-outline-container.close { width: 0 }
.${c}-collapse-btn-wrapper { width: 1.5em; height: 100%; background: gray; }
.${c}-collapse-btn { transform: rotate(90deg); width: 1.5em; height: 1.5em; background: gray; object-fit: contain; cursor: pointer; transition: transform .8s ease-out; }
.${c}-collapse-btn.close { transform: rotate(0deg); transition: transform .8s ease-out; }
.${c}-outline-content { background: white; }
.${c}-outline-content > p { text-align: left; margin-top: 3px; margin-bottom: 3px; margin-right: 30px; cursor: pointer; }
.${c}-outline-content > p span { font-size: 12px !important; color: #333; }
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
.${c} { color: black; hyphens: auto; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c}-prev-comment { margin-left: 10px; transform: rotate(-90deg); cursor: pointer; }
.${c}-next-comment { margin-right: 10px; transform: rotate(90deg); cursor: pointer; }
.${c}-prev-comment:hover svg path { fill: #2A66FF; }
.${c}-next-comment:hover svg path { fill: #2A66FF; }
.${c}-all-comment { height: 22px; line-height: 19px; padding-right: 10px; border-right: 1px solid #ddd; color: gray; }
.${c}-current-comment { height: 22px; line-height: 19px; margin-left: 20px; color: gray; }
`;
        return createStyleElement(styleText);
    }
    renderNumbering(numberings, styleContainer) {
        var styleText = "";
        var resetCounters = [];
        for (var num of numberings) {
            var selector = `p.${this.numberingClass(num.id, num.level)}`;
            var listStyleType = "none";
            if (num.bullet) {
                let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
                styleText += this.styleToString(`${selector}:before`, {
                    "content": "' '",
                    "display": "inline-block",
                    "background": `var(${valiable})`
                }, num.bullet.style);
                this.document.loadNumberingImage(num.bullet.src).then(data => {
                    var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else if (num.levelText) {
                let counter = this.numberingCounter(num.id, num.level);
                const counterReset = counter + " " + (num.start - 1);
                if (num.level > 0) {
                    styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                        "counter-reset": counterReset
                    });
                }
                resetCounters.push(counterReset);
                styleText += this.styleToString(`${selector}:before`, {
                    "content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
                    "counter-increment": counter,
                    ...num.rStyle,
                });
            }
            else {
                listStyleType = this.numFormatToCssValue(num.format);
            }
            styleText += this.styleToString(selector, {
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": listStyleType,
                ...num.pStyle
            });
        }
        if (resetCounters.length > 0) {
            styleText += this.styleToString(this.rootSelector, {
                "counter-reset": resetCounters.join(" ")
            });
        }
        return createStyleElement(styleText);
    }
    renderTitleNumbering(pArr, domNumberings, styleContainer) {
        let styleText = "";
        let resetCounters = [];
        for (let p of pArr) {
            if (!isNaN(Number(p.styleName)) && Number(p.styleName) !== 0 && p?.numbering?.id !== '0') {
                let titleLevel = p.styleName;
                const style = this.findStyle(p.styleName);
                let numbering = p.numbering;
                if (numbering && numbering.level !== undefined && numbering.level !== null && (numbering.level > (Number(titleLevel) - 1))) {
                    titleLevel = numbering.level;
                }
                else {
                    titleLevel = Number(titleLevel - 1);
                }
                numbering = numbering ?? style?.paragraphProps?.numbering;
                if (numbering === undefined || numbering === null) {
                    continue;
                }
                let currentDomNumbering = domNumberings.find(item => {
                    return item.numId === numbering.id && item.level === numbering.level;
                });
                let selector = `p.${this.numberingTitleClass(titleLevel)}`;
                let listStyleType = "none";
                if (currentDomNumbering.bullet) {
                    let valiable = `--${this.className}-${currentDomNumbering.bullet.src}`.toLowerCase();
                    styleText += this.styleToString(`${selector}:before`, {
                        "content": "' '",
                        "display": "inline-block",
                        "background": `var(${valiable})`
                    }, currentDomNumbering.bullet.style);
                    this.document.loadNumberingImage(currentDomNumbering.bullet.src).then(data => {
                        var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
                        styleContainer.appendChild(createStyleElement(text));
                    });
                }
                else if (currentDomNumbering.levelText) {
                    let counter = this.numberingTitleCounter(titleLevel);
                    const counterReset = counter + " " + 0;
                    if (Number(titleLevel) > 0) {
                        styleText += this.styleToString(`p.${this.numberingTitleClass(Number(titleLevel) - 1)}`, {
                            "counter-reset": counterReset
                        });
                    }
                    resetCounters.push(counterReset);
                    styleText += this.styleToString(`${selector}:before`, {
                        "content": this.levelTextToContentOfTitle(currentDomNumbering.levelText, currentDomNumbering.suff, this.numFormatToCssValue(currentDomNumbering.format)),
                        "counter-increment": counter,
                        ...currentDomNumbering.rStyle,
                    });
                }
                else {
                    listStyleType = this.numFormatToCssValue(currentDomNumbering.format);
                }
                styleText += this.styleToString(selector, {
                    "display": "list-item",
                    "list-style-position": "inside",
                    "list-style-type": listStyleType,
                    ...currentDomNumbering.pStyle
                });
            }
        }
        for (var num of domNumberings) {
            var selector = `p.${this.numberingClass(num.id, num.level)}`;
            var listStyleType = "none";
            if (num.bullet) {
                let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
                styleText += this.styleToString(`${selector}:before`, {
                    "content": "' '",
                    "display": "inline-block",
                    "background": `var(${valiable})`
                }, num.bullet.style);
                this.document.loadNumberingImage(num.bullet.src).then(data => {
                    var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else if (num.levelText) {
                let counter = this.numberingCounter(num.id, num.level);
                const counterReset = counter + " " + (num.start - 1);
                if (num.level > 0) {
                    styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                        "counter-reset": counterReset
                    });
                }
                resetCounters.push(counterReset);
                styleText += this.styleToString(`${selector}:before`, {
                    "content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
                    "counter-increment": counter,
                    ...num.rStyle,
                });
            }
            else {
                listStyleType = this.numFormatToCssValue(num.format);
            }
            styleText += this.styleToString(selector, {
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": listStyleType,
                ...num.pStyle
            });
        }
        if (resetCounters.length > 0) {
            styleText += this.styleToString(this.rootSelector, {
                "counter-reset": resetCounters.join(" ")
            });
        }
        return createStyleElement(styleText);
    }
    renderMyNumbering(numberings, overNumberings, styleContainer) {
        var styleText = "";
        var resetCounters = [];
        for (var num of numberings) {
            var selector = `p.${this.numberingClass(num.id, num.level)}`;
            var listStyleType = "none";
            if (num.bullet) {
                let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
                styleText += this.styleToString(`${selector}:before`, {
                    "content": "' '",
                    "display": "inline-block",
                    "background": `var(${valiable})`
                }, num.bullet.style);
                this.document.loadNumberingImage(num.bullet.src).then(data => {
                    var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else if (num.levelText) {
                let counter = this.numberingCounter(num.id, num.level);
                const counterReset = counter + " " + (num.start - 1);
                if (num.level > 0) {
                    styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                        "counter-reset": counterReset
                    });
                }
                resetCounters.push(counterReset);
                styleText += this.styleToString(`${selector}:before`, {
                    "content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
                    "counter-increment": counter,
                    ...num.rStyle,
                });
            }
            else {
                listStyleType = this.numFormatToCssValue(num.format);
            }
            styleText += this.styleToString(selector, {
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": listStyleType,
                ...num.pStyle
            });
        }
        if (resetCounters.length > 0) {
            styleText += this.styleToString(this.rootSelector, {
                "counter-reset": resetCounters.join(" ")
            });
        }
        return createStyleElement(styleText);
    }
    renderStyles(styles) {
        var styleText = "";
        const stylesMap = this.styleMap;
        const defautStyles = (0, utils_1.keyBy)(styles.filter(s => s.isDefault), s => s.target);
        for (const style of styles) {
            var subStyles = style.styles;
            if (style.linked) {
                var linkedStyle = style.linked && stylesMap[style.linked];
                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if (this.options.debug)
                    console.warn(`Can't find linked style ${style.linked}`);
            }
            for (const subStyle of subStyles) {
                var selector = `${style.target ?? ''}.${style.cssName}`;
                if (style.target != subStyle.target)
                    selector += ` ${subStyle.target}`;
                if (defautStyles[style.target] == style)
                    selector = `.${this.className} ${style.target}, ` + selector;
                styleText += this.styleToString(selector, subStyle.values);
            }
        }
        return createStyleElement(styleText);
    }
    renderNotes(noteIds, notesMap, into) {
        var notes = noteIds.map(id => notesMap[id]).filter(x => x);
        if (notes.length > 0) {
            var result = this.createElement("ol", null, this.renderElements(notes));
            into.appendChild(result);
        }
    }
    renderElement(elem) {
        switch (elem.type) {
            case dom_1.DomType.Paragraph:
                return this.renderParagraph(elem);
            case dom_1.DomType.BookmarkStart:
                return this.renderBookmarkStart(elem);
            case dom_1.DomType.BookmarkEnd:
                return null;
            case dom_1.DomType.Run:
                return this.renderRun(elem);
            case dom_1.DomType.Table:
                return this.renderTable(elem);
            case dom_1.DomType.Row:
                return this.renderTableRow(elem);
            case dom_1.DomType.Cell:
                return this.renderTableCell(elem);
            case dom_1.DomType.Hyperlink:
                return this.renderHyperlink(elem);
            case dom_1.DomType.Drawing:
                return this.renderDrawing(elem);
            case dom_1.DomType.Image:
                return this.renderImage(elem);
            case dom_1.DomType.Text:
                return this.renderText(elem);
            case dom_1.DomType.Text:
                return this.renderText(elem);
            case dom_1.DomType.DeletedText:
                return this.renderDeletedText(elem);
            case dom_1.DomType.Tab:
                return this.renderTab(elem);
            case dom_1.DomType.Symbol:
                return this.renderSymbol(elem);
            case dom_1.DomType.Break:
                return this.renderBreak(elem);
            case dom_1.DomType.Footer:
                return this.renderContainer(elem, "footer");
            case dom_1.DomType.Header:
                return this.renderContainer(elem, "header");
            case dom_1.DomType.Footnote:
            case dom_1.DomType.Endnote:
                let li = this.renderContainer(elem, "li");
                li.style.fontSize = "10pt";
                return li;
            case dom_1.DomType.FootnoteReference:
                return this.renderFootnoteReference(elem);
            case dom_1.DomType.EndnoteReference:
                return this.renderEndnoteReference(elem);
            case dom_1.DomType.NoBreakHyphen:
                return this.createElement("wbr");
            case dom_1.DomType.VmlPicture:
                return this.renderVmlPicture(elem);
            case dom_1.DomType.VmlElement:
                return this.renderVmlElement(elem);
            case dom_1.DomType.MmlMath:
                return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
            case dom_1.DomType.MmlMathParagraph:
                return this.renderContainer(elem, "span");
            case dom_1.DomType.MmlFraction:
                return this.renderContainerNS(elem, ns.mathML, "mfrac");
            case dom_1.DomType.MmlBase:
                return this.renderContainerNS(elem, ns.mathML, elem.parent.type == dom_1.DomType.MmlMatrixRow ? "mtd" : "mrow");
            case dom_1.DomType.MmlNumerator:
            case dom_1.DomType.MmlDenominator:
            case dom_1.DomType.MmlFunction:
            case dom_1.DomType.MmlLimit:
            case dom_1.DomType.MmlBox:
                return this.renderContainerNS(elem, ns.mathML, "mrow");
            case dom_1.DomType.MmlGroupChar:
                return this.renderMmlGroupChar(elem);
            case dom_1.DomType.MmlLimitLower:
                return this.renderContainerNS(elem, ns.mathML, "munder");
            case dom_1.DomType.MmlMatrix:
                return this.renderContainerNS(elem, ns.mathML, "mtable");
            case dom_1.DomType.MmlMatrixRow:
                return this.renderContainerNS(elem, ns.mathML, "mtr");
            case dom_1.DomType.MmlRadical:
                return this.renderMmlRadical(elem);
            case dom_1.DomType.MmlSuperscript:
                return this.renderContainerNS(elem, ns.mathML, "msup");
            case dom_1.DomType.MmlSubscript:
                return this.renderContainerNS(elem, ns.mathML, "msub");
            case dom_1.DomType.MmlDegree:
            case dom_1.DomType.MmlSuperArgument:
            case dom_1.DomType.MmlSubArgument:
                return this.renderContainerNS(elem, ns.mathML, "mn");
            case dom_1.DomType.MmlFunctionName:
                return this.renderContainerNS(elem, ns.mathML, "ms");
            case dom_1.DomType.MmlDelimiter:
                return this.renderMmlDelimiter(elem);
            case dom_1.DomType.MmlRun:
                return this.renderMmlRun(elem);
            case dom_1.DomType.MmlNary:
                return this.renderMmlNary(elem);
            case dom_1.DomType.MmlPreSubSuper:
                return this.renderMmlPreSubSuper(elem);
            case dom_1.DomType.MmlBar:
                return this.renderMmlBar(elem);
            case dom_1.DomType.MmlEquationArray:
                return this.renderMllList(elem);
            case dom_1.DomType.Inserted:
                return this.renderInserted(elem);
            case dom_1.DomType.Deleted:
                return this.renderDeleted(elem);
            case dom_1.DomType.CommentRangeStart:
                return this.renderCommentRangeStart(elem);
            case dom_1.DomType.CommentRangeEnd:
                return this.renderCommentRangeEnd(elem);
        }
        return null;
    }
    renderFieldRun(elem) {
        for (let child of elem.children ?? []) {
            switch (child.type) {
                case dom_1.DomType.Instruction:
                    return this.renderInstrText(child);
            }
            return null;
        }
        return null;
    }
    renderInstrText(elem) {
        let text = elem.text;
        if (text.split(' ')[0] === "REF") {
            let id = Math.random();
            elem.instrTextId = id;
            let result = this.createElement("span");
            let paragraph = elem.parent?.parent;
            let textArr = [];
            let rArr = paragraph?.children || [];
            let index = rArr?.findIndex(item => {
                let rChildren = item.children;
                let c = rChildren?.find(c => {
                    return c?.instrTextId === id;
                });
                if (c !== undefined)
                    return true;
                return false;
            });
            for (let i = index; i < rArr?.length || 0; i++) {
                let rChildren = rArr?.[i]?.children;
                let child = rChildren?.find(c => {
                    return c?.type === 'complexField' && c?.charType === 'end';
                });
                if (child === undefined) {
                    if (rArr?.[i].fieldRun !== true) {
                        textArr.push(rArr?.[i]);
                        if (rArr?.[i] !== undefined) {
                            rArr[i].fieldRun = true;
                        }
                    }
                }
                else {
                    break;
                }
            }
            result.textContent = this.joinRangeText(textArr);
            result.style.cursor = "pointer";
            let that = this;
            result.addEventListener("click", function () {
                let ref = text.split(' ')[1];
                let wrap = document.getElementsByClassName(`${that.className}-wrapper`)[0];
                let anchor = wrap.querySelector(`#${ref}`);
                anchor?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
            });
            return result;
        }
        return null;
    }
    renderCommentRangeStart(elem) {
        let comments = this.document.commentsPart.comments;
        let commentsEx = this.document.commentsExtendedPart.commentsEx;
        let paraId = comments.find(item => { return item.id === elem.id; })?.paraId;
        let paraIdParent = commentsEx.find(item => { return item.paraId === paraId; })?.paraIdParent;
        if (paraIdParent) {
            return null;
        }
        return this.createCommentSupNode('start', paraId);
    }
    renderCommentRangeEnd(elem) {
        let comments = this.document.commentsPart.comments;
        let commentsEx = this.document.commentsExtendedPart.commentsEx;
        let paraId = comments.find(item => { return item.id === elem.id; })?.paraId;
        let paraIdParent = commentsEx.find(item => { return item.paraId === paraId; })?.paraIdParent;
        if (paraIdParent) {
            return null;
        }
        let endElem = this.createCommentSupNode('end', paraId);
        endElem.className = `${this.className}-comment-range-end`;
        return endElem;
    }
    createCommentSupNode(type, paraId) {
        let elem = this.createElement("span");
        elem.id = `${this.className}-comment-${type}-${paraId}`;
        this.commentIndex += 1;
        return elem;
    }
    renderChildren(elem, into) {
        return this.renderElements(elem.children, into);
    }
    renderElements(elems, into) {
        if (elems == null)
            return null;
        var result = elems.flatMap(e => this.renderElement(e)).filter(e => e != null);
        if (into)
            appendChildren(into, result);
        return result;
    }
    renderContainer(elem, tagName, props) {
        return this.createElement(tagName, props, this.renderChildren(elem));
    }
    renderContainerNS(elem, ns, tagName, props) {
        return createElementNS(ns, tagName, props, this.renderChildren(elem));
    }
    renderParagraph(elem) {
        var result = this.createElement("p");
        const style = this.findStyle(elem.styleName);
        elem.tabs ?? (elem.tabs = style?.paragraphProps?.tabs);
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        this.renderCommonProperties(result.style, elem);
        const numbering = elem.numbering ?? style?.paragraphProps?.numbering;
        if (this.options.renderNumbering && numbering && !isNaN(Number(elem.styleName)) && Number(elem.styleName) !== 0 && Number(elem.styleName) < 4 && elem?.numbering?.id !== '0') {
            let titleLevel = elem.styleName;
            if (elem.numbering && elem.numbering.level !== undefined && elem.numbering.level !== null && (elem.numbering.level > (Number(titleLevel) - 1))) {
                titleLevel = `${elem.numbering.level}`;
            }
            else {
                titleLevel = `${(Number(titleLevel) - 1)}`;
            }
            result.classList.add(this.numberingTitleClass(Number(titleLevel)));
            return result;
        }
        if (this.options.renderNumbering && numbering && !elem.noRenderNumbering) {
            let numberingPart = this.document.numberingPart.numberings;
            let currentNumbering = numberingPart.find(item => {
                return item.id === numbering.id;
            });
            result.classList.add(this.numberingClass(currentNumbering?.abstractId, numbering.level));
            let currentOverride = currentNumbering?.overrides?.find(item => { return item.level === numbering.level; });
            if (currentOverride) {
                let overrideCounter = this.numberingCounter(currentNumbering?.abstractId, numbering.level);
                let overrideCounterReset = overrideCounter + " " + (currentOverride.start ? currentOverride.start - 1 : 0);
                result.style.counterReset = overrideCounterReset;
            }
        }
        return result;
    }
    renderRunProperties(style, props) {
        this.renderCommonProperties(style, props);
    }
    renderCommonProperties(style, props) {
        if (props == null)
            return;
        if (props.color) {
            style["color"] = props.color;
        }
        if (props.fontSize) {
            style["font-size"] = props.fontSize;
        }
    }
    renderHyperlink(elem) {
        var result = this.createElement("a");
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (elem.href) {
            result.style.cursor = "pointer";
            let that = this;
            result.addEventListener("click", function (e) {
                let wrap = document.getElementsByClassName(`${that.className}-wrapper`)[0];
                let anchor = wrap.querySelector(`#${elem?.href?.replace("#", '')}`);
                anchor?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
            });
        }
        else if (elem.id) {
            const rel = this.document.documentPart.rels
                .find(it => it.id == elem.id && it.targetMode === "External");
            result.href = rel?.target;
        }
        return result;
    }
    renderDrawing(elem) {
        var result = this.createElement("div");
        result.style.display = "inline-block";
        result.style.position = "relative";
        result.style.textIndent = "0px";
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        return result;
    }
    renderImage(elem) {
        let result = this.createElement("img");
        this.renderStyleValues(elem.cssStyle, result);
        if (this.document) {
            this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
                result.src = x;
            });
        }
        return result;
    }
    renderText(elem) {
        return this.htmlDocument.createTextNode(elem.text);
    }
    renderDeletedText(elem) {
        return this.options.renderEndnotes ? this.htmlDocument.createTextNode(elem.text) : null;
    }
    renderBreak(elem) {
        if (elem.break == "textWrapping") {
            return this.createElement("br");
        }
        return null;
    }
    renderInserted(elem) {
        if (this.options.renderChanges)
            return this.renderContainer(elem, "ins");
        return this.renderChildren(elem);
    }
    renderDeleted(elem) {
        if (this.options.renderChanges)
            return this.renderContainer(elem, "del");
        return null;
    }
    renderSymbol(elem) {
        var span = this.createElement("span");
        span.style.fontFamily = elem.font;
        span.innerHTML = `&#x${elem.char};`;
        return span;
    }
    renderFootnoteReference(elem) {
        var result = this.createElement("sup");
        this.currentFootnoteIds.push(elem.id);
        result.textContent = `${this.currentFootnoteIds.length}`;
        let notes = this.document?.footnotesPart?.notes;
        result.title = this.renderNotePreview(elem.id, notes);
        return result;
    }
    renderEndnoteReference(elem) {
        var result = this.createElement("sup");
        this.currentEndnoteIds.push(elem.id);
        result.textContent = `${this.currentFootnoteIds.length}`;
        let notes = this.document?.endnotesPart?.notes;
        result.title = this.renderNotePreview(elem.id, notes);
        return result;
    }
    renderNotePreview(id, notes) {
        let note = notes.find(item => {
            return item.id === id;
        });
        if (note) {
            let rArr = [];
            let pArr = note.children;
            for (let p of pArr) {
                rArr = rArr.concat(p?.children || []);
            }
            let result = this.joinRangeText(rArr);
            return result;
        }
        return '';
    }
    renderTab(elem) {
        var tabSpan = this.createElement("span");
        tabSpan.innerHTML = "&emsp;";
        if (this.options.experimental) {
            tabSpan.className = this.tabStopClass();
            var stops = findParent(elem, dom_1.DomType.Paragraph)?.tabs;
            this.currentTabs.push({ stops, span: tabSpan });
        }
        return tabSpan;
    }
    renderBookmarkStart(elem) {
        var result = this.createElement("span");
        result.id = elem.name;
        return result;
    }
    renderRun(elem) {
        if (elem.fieldRun)
            return this.renderFieldRun(elem);
        const result = this.createElement("span");
        if (elem.id)
            result.id = elem.id;
        this.renderClass(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (elem.verticalAlign) {
            const wrapper = this.createElement(elem.verticalAlign);
            this.renderChildren(elem, wrapper);
            result.appendChild(wrapper);
        }
        else {
            this.renderChildren(elem, result);
        }
        return result;
    }
    renderTable(elem) {
        let result = this.createElement("table");
        this.tableCellPositions.push(this.currentCellPosition);
        this.tableVerticalMerges.push(this.currentVerticalMerge);
        this.currentVerticalMerge = {};
        this.currentCellPosition = { col: 0, row: 0 };
        if (elem.columns)
            result.appendChild(this.renderTableColumns(elem.columns));
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        let tblpXSpec = elem.cssStyle?.tblpXSpec;
        if (tblpXSpec === 'center') {
            let num = parseFloat(elem.cssStyle?.width || '0');
            let unit = elem.cssStyle?.width.replace(/[0-9]*/g, '').replace(/\./, '');
            let left = ((100 - Number(num)) / 2).toFixed(2);
            elem.cssStyle['margin-left'] = `${left}${unit}`;
        }
        else if (tblpXSpec) {
            elem.cssStyle.float = tblpXSpec;
        }
        this.renderStyleValues(elem.cssStyle, result);
        this.currentVerticalMerge = this.tableVerticalMerges.pop();
        this.currentCellPosition = this.tableCellPositions.pop();
        return result;
    }
    renderTableColumns(columns) {
        let result = this.createElement("colgroup");
        for (let col of columns) {
            let colElem = this.createElement("col");
            if (col.width)
                colElem.style.width = col.width;
            result.appendChild(colElem);
        }
        return result;
    }
    renderTableRow(elem) {
        let result = this.createElement("tr");
        if (this.currentCellPosition) {
            this.currentCellPosition.col = 0;
        }
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (this.currentCellPosition) {
            this.currentCellPosition.row++;
        }
        return result;
    }
    renderTableCell(elem) {
        let result = this.createElement("td");
        const key = this.currentCellPosition?.col;
        if (elem.verticalMerge) {
            if (elem.verticalMerge == "restart") {
                this.currentVerticalMerge[key] = result;
                result.rowSpan = 1;
            }
            else if (this.currentVerticalMerge[key]) {
                this.currentVerticalMerge[key].rowSpan += 1;
                result.style.display = "none";
            }
        }
        else {
            if (this.currentVerticalMerge) {
                this.currentVerticalMerge[key] = null;
            }
        }
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (elem.span)
            result.colSpan = elem.span;
        if (this.currentCellPosition) {
            this.currentCellPosition.col += result.colSpan;
        }
        return result;
    }
    renderVmlPicture(elem) {
        var result = createElement("div");
        this.renderChildren(elem, result);
        return result;
    }
    renderVmlElement(elem) {
        var container = createSvgElement("svg");
        container.setAttribute("style", elem.cssStyleText);
        const result = this.renderVmlChildElement(elem);
        if (elem.imageHref?.id) {
            this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
                .then(x => result.setAttribute("href", x));
        }
        container.appendChild(result);
        requestAnimationFrame(() => {
            const bb = container.firstElementChild.getBBox();
            container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
            container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
        });
        return container;
    }
    renderVmlChildElement(elem) {
        const result = createSvgElement(elem.tagName);
        Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));
        for (let child of elem.children) {
            if (child.type == dom_1.DomType.VmlElement) {
                result.appendChild(this.renderVmlChildElement(child));
            }
            else {
                result.appendChild(...(0, utils_1.asArray)(this.renderElement(child)));
            }
        }
        return result;
    }
    renderMmlRadical(elem) {
        const base = elem.children.find(el => el.type == dom_1.DomType.MmlBase);
        if (elem.props?.hideDegree) {
            return createElementNS(ns.mathML, "msqrt", null, this.renderElements([base]));
        }
        const degree = elem.children.find(el => el.type == dom_1.DomType.MmlDegree);
        return createElementNS(ns.mathML, "mroot", null, this.renderElements([base, degree]));
    }
    renderMmlDelimiter(elem) {
        const children = [];
        children.push(createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
        children.push(...this.renderElements(elem.children));
        children.push(createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));
        return createElementNS(ns.mathML, "mrow", null, children);
    }
    renderMmlNary(elem) {
        const children = [];
        const grouped = (0, utils_1.keyBy)(elem.children, x => x.type);
        const sup = grouped[dom_1.DomType.MmlSuperArgument];
        const sub = grouped[dom_1.DomType.MmlSubArgument];
        const supElem = sup ? createElementNS(ns.mathML, "mo", null, (0, utils_1.asArray)(this.renderElement(sup))) : null;
        const subElem = sub ? createElementNS(ns.mathML, "mo", null, (0, utils_1.asArray)(this.renderElement(sub))) : null;
        const charElem = createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);
        if (supElem || subElem) {
            children.push(createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
        }
        else if (supElem) {
            children.push(createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
        }
        else if (subElem) {
            children.push(createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
        }
        else {
            children.push(charElem);
        }
        children.push(...this.renderElements(grouped[dom_1.DomType.MmlBase].children));
        return createElementNS(ns.mathML, "mrow", null, children);
    }
    renderMmlPreSubSuper(elem) {
        const children = [];
        const grouped = (0, utils_1.keyBy)(elem.children, x => x.type);
        const sup = grouped[dom_1.DomType.MmlSuperArgument];
        const sub = grouped[dom_1.DomType.MmlSubArgument];
        const supElem = sup ? createElementNS(ns.mathML, "mo", null, (0, utils_1.asArray)(this.renderElement(sup))) : null;
        const subElem = sub ? createElementNS(ns.mathML, "mo", null, (0, utils_1.asArray)(this.renderElement(sub))) : null;
        const stubElem = createElementNS(ns.mathML, "mo", null);
        children.push(createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));
        children.push(...this.renderElements(grouped[dom_1.DomType.MmlBase].children));
        return createElementNS(ns.mathML, "mrow", null, children);
    }
    renderMmlGroupChar(elem) {
        const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
        const result = this.renderContainerNS(elem, ns.mathML, tagName);
        if (elem.props.char) {
            result.appendChild(createElementNS(ns.mathML, "mo", null, [elem.props.char]));
        }
        return result;
    }
    renderMmlBar(elem) {
        const result = this.renderContainerNS(elem, ns.mathML, "mrow");
        switch (elem.props.position) {
            case "top":
                result.style.textDecoration = "overline";
                break;
            case "bottom":
                result.style.textDecoration = "underline";
                break;
        }
        return result;
    }
    renderMmlRun(elem) {
        const result = createElementNS(ns.mathML, "ms");
        this.renderClass(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        this.renderChildren(elem, result);
        return result;
    }
    renderMllList(elem) {
        const result = createElementNS(ns.mathML, "mtable");
        this.renderClass(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        const childern = this.renderChildren(elem);
        for (let child of this.renderChildren(elem)) {
            result.appendChild(createElementNS(ns.mathML, "mtr", null, [
                createElementNS(ns.mathML, "mtd", null, [child])
            ]));
        }
        return result;
    }
    renderStyleValues(style, ouput) {
        for (let k in style) {
            if (k.startsWith("$")) {
                ouput.setAttribute(k.slice(1), style[k]);
            }
            else {
                ouput.style[k] = style[k];
            }
        }
    }
    renderClass(input, ouput) {
        if (input.className)
            ouput.className = input.className;
        if (input.styleName)
            ouput.classList.add(this.processStyleName(input.styleName));
    }
    findStyle(styleName) {
        return styleName && this.styleMap?.[styleName];
    }
    numberingClass(id, lvl) {
        return `${this.className}-num-${id}-${lvl}`;
    }
    numberingTitleClass(lvl) {
        return `${this.className}-title-${lvl}`;
    }
    tabStopClass() {
        return `${this.className}-tab-stop`;
    }
    styleToString(selectors, values, cssText = null) {
        let result = `${selectors} {\r\n`;
        for (const key in values) {
            if (key.startsWith('$'))
                continue;
            result += `  ${key}: ${values[key]};\r\n`;
        }
        if (cssText)
            result += cssText;
        return result + "}\r\n";
    }
    numberingCounter(id, lvl) {
        return `${this.className}-num-${id}-${lvl}`;
    }
    numberingTitleCounter(lvl) {
        return `${this.className}-title-${lvl}`;
    }
    levelTextToContent(text, suff, id, numformat) {
        const suffMap = {
            "tab": "\\9",
            "space": "\\a0",
        };
        var result = text.replace(/%\d*/g, s => {
            let lvl = parseInt(s.substring(1), 10) - 1;
            return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
        });
        return `"${result}${suffMap[suff] ?? ""}"`;
    }
    levelTextToContentOfTitle(text, suff, numformat) {
        const suffMap = {
            "tab": "\\9",
            "space": "\\a0",
        };
        var result = text.replace(/%\d*/g, s => {
            let lvl = parseInt(s.substring(1), 10) - 1;
            return `"counter(${this.numberingTitleCounter(lvl)}, ${numformat})"`;
        });
        return `"${result}${suffMap[suff] ?? ""}"`;
    }
    numFormatToCssValue(format) {
        var mapping = {
            none: "none",
            bullet: "disc",
            decimal: "decimal",
            lowerLetter: "lower-alpha",
            upperLetter: "upper-alpha",
            lowerRoman: "lower-roman",
            upperRoman: "upper-roman",
            decimalZero: "decimal-leading-zero",
            aiueo: "katakana",
            aiueoFullWidth: "katakana",
            chineseCounting: "simp-chinese-informal",
            chineseCountingThousand: "simp-chinese-informal",
            chineseLegalSimplified: "simp-chinese-formal",
            chosung: "hangul-consonant",
            ideographDigital: "cjk-ideographic",
            ideographTraditional: "cjk-heavenly-stem",
            ideographLegalTraditional: "trad-chinese-formal",
            ideographZodiac: "cjk-earthly-branch",
            iroha: "katakana-iroha",
            irohaFullWidth: "katakana-iroha",
            japaneseCounting: "japanese-informal",
            japaneseDigitalTenThousand: "cjk-decimal",
            japaneseLegal: "japanese-formal",
            thaiNumbers: "thai",
            koreanCounting: "korean-hangul-formal",
            koreanDigital: "korean-hangul-formal",
            koreanDigital2: "korean-hanja-informal",
            hebrew1: "hebrew",
            hebrew2: "hebrew",
            hindiNumbers: "devanagari",
            ganada: "hangul",
            taiwaneseCounting: "cjk-ideographic",
            taiwaneseCountingThousand: "cjk-ideographic",
            taiwaneseDigital: "cjk-decimal",
        };
        return mapping[format] ?? format;
    }
    refreshTabStops() {
        if (!this.options.experimental)
            return;
        clearTimeout(this.tabsTimeout);
        this.tabsTimeout = setTimeout(() => {
            const pixelToPoint = (0, javascript_1.computePixelToPoint)();
            for (let tab of this.currentTabs) {
                (0, javascript_1.updateTabStop)(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
            }
        }, 500);
    }
}
exports.HtmlRenderer = HtmlRenderer;
function createElement(tagName, props, children) {
    return createElementNS(undefined, tagName, props, children);
}
function createSvgElement(tagName, props, children) {
    return createElementNS(ns.svg, tagName, props, children);
}
function createElementNS(ns, tagName, props, children) {
    var result = ns ? document.createElementNS(ns, tagName) : document.createElement(tagName);
    Object.assign(result, props);
    children && appendChildren(result, children);
    return result;
}
function removeAllElements(elem) {
    elem.innerHTML = '';
}
function appendChildren(elem, children) {
    children.forEach(c => elem.appendChild((0, utils_1.isString)(c) ? document.createTextNode(c) : c));
}
function createStyleElement(cssText) {
    return createElement("style", { innerHTML: cssText });
}
function createScriptElement(jsText) {
    return createElement("script", { innerHTML: jsText });
}
function appendComment(elem, comment) {
    elem.appendChild(document.createComment(comment));
}
function findParent(elem, type) {
    var parent = elem.parent;
    while (parent != null && parent.type != type)
        parent = parent.parent;
    return parent;
}


/***/ }),

/***/ "./src/javascript.ts":
/*!***************************!*\
  !*** ./src/javascript.ts ***!
  \***************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.updateTabStop = exports.computePixelToPoint = void 0;
const defaultTab = { pos: 0, leader: "none", style: "left" };
const maxTabs = 50;
function computePixelToPoint(container = document.body) {
    const temp = document.createElement("div");
    temp.style.width = '100pt';
    container.appendChild(temp);
    const result = 100 / temp.offsetWidth;
    container.removeChild(temp);
    return result;
}
exports.computePixelToPoint = computePixelToPoint;
function updateTabStop(elem, tabs, defaultTabSize, pixelToPoint = 72 / 96) {
    const p = elem.closest("p");
    const ebb = elem.getBoundingClientRect();
    const pbb = p.getBoundingClientRect();
    const pcs = getComputedStyle(p);
    const tabStops = tabs?.length > 0 ? tabs.map(t => ({
        pos: lengthToPoint(t.position),
        leader: t.leader,
        style: t.style
    })).sort((a, b) => a.pos - b.pos) : [defaultTab];
    const lastTab = tabStops[tabStops.length - 1];
    const pWidthPt = pbb.width * pixelToPoint;
    const size = lengthToPoint(defaultTabSize);
    let pos = lastTab.pos + size;
    if (pos < pWidthPt) {
        for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) {
            tabStops.push({ ...defaultTab, pos: pos });
        }
    }
    const marginLeft = parseFloat(pcs.marginLeft);
    const pOffset = pbb.left + marginLeft;
    const left = (ebb.left - pOffset) * pixelToPoint;
    const tab = tabStops.find(t => t.style != "clear" && t.pos > left);
    if (tab == null)
        return;
    let width = 1;
    if (tab.style == "right" || tab.style == "center") {
        const tabStops = Array.from(p.querySelectorAll(`.${elem.className}`));
        const nextIdx = tabStops.indexOf(elem) + 1;
        const range = document.createRange();
        range.setStart(elem, 1);
        if (nextIdx < tabStops.length) {
            range.setEndBefore(tabStops[nextIdx]);
        }
        else {
            range.setEndAfter(p);
        }
        const mul = tab.style == "center" ? 0.5 : 1;
        const nextBB = range.getBoundingClientRect();
        const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);
        width = tab.pos - offset * pixelToPoint;
    }
    else {
        width = tab.pos - left;
    }
    elem.innerHTML = "&nbsp;";
    elem.style.textDecoration = "inherit";
    elem.style.wordSpacing = `${width.toFixed(0)}pt`;
    switch (tab.leader) {
        case "dot":
        case "middleDot":
            elem.style.textDecoration = "underline";
            elem.style.textDecorationStyle = "dotted";
            break;
        case "hyphen":
        case "heavy":
        case "underscore":
            elem.style.textDecoration = "underline";
            break;
    }
}
exports.updateTabStop = updateTabStop;
function lengthToPoint(length) {
    return parseFloat(length);
}


/***/ }),

/***/ "./src/notes/elements.ts":
/*!*******************************!*\
  !*** ./src/notes/elements.ts ***!
  \*******************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.WmlEndnote = exports.WmlFootnote = exports.WmlBaseNote = void 0;
const dom_1 = __webpack_require__(/*! ../document/dom */ "./src/document/dom.ts");
class WmlBaseNote {
}
exports.WmlBaseNote = WmlBaseNote;
class WmlFootnote extends WmlBaseNote {
    constructor() {
        super(...arguments);
        this.type = dom_1.DomType.Footnote;
    }
}
exports.WmlFootnote = WmlFootnote;
class WmlEndnote extends WmlBaseNote {
    constructor() {
        super(...arguments);
        this.type = dom_1.DomType.Endnote;
    }
}
exports.WmlEndnote = WmlEndnote;


/***/ }),

/***/ "./src/notes/parts.ts":
/*!****************************!*\
  !*** ./src/notes/parts.ts ***!
  \****************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.EndnotesPart = exports.FootnotesPart = exports.BaseNotePart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const elements_1 = __webpack_require__(/*! ./elements */ "./src/notes/elements.ts");
class BaseNotePart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
}
exports.BaseNotePart = BaseNotePart;
class FootnotesPart extends BaseNotePart {
    constructor(pkg, path, parser) {
        super(pkg, path, parser);
    }
    parseXml(root) {
        this.notes = this._documentParser.parseNotes(root, "footnote", elements_1.WmlFootnote);
    }
}
exports.FootnotesPart = FootnotesPart;
class EndnotesPart extends BaseNotePart {
    constructor(pkg, path, parser) {
        super(pkg, path, parser);
    }
    parseXml(root) {
        this.notes = this._documentParser.parseNotes(root, "endnote", elements_1.WmlEndnote);
    }
}
exports.EndnotesPart = EndnotesPart;


/***/ }),

/***/ "./src/numbering/numbering-part.ts":
/*!*****************************************!*\
  !*** ./src/numbering/numbering-part.ts ***!
  \*****************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.NumberingPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const numbering_1 = __webpack_require__(/*! ./numbering */ "./src/numbering/numbering.ts");
class NumberingPart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    parseXml(root) {
        Object.assign(this, (0, numbering_1.parseNumberingPart)(root, this._package.xmlParser));
        this.domNumberings = this._documentParser.parseNumberingFile(root, this.numberings);
    }
}
exports.NumberingPart = NumberingPart;


/***/ }),

/***/ "./src/numbering/numbering.ts":
/*!************************************!*\
  !*** ./src/numbering/numbering.ts ***!
  \************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseNumberingBulletPicture = exports.parseNumberingLevelOverrride = exports.parseNumberingLevel = exports.parseAbstractNumbering = exports.parseNumbering = exports.parseNumberingPart = void 0;
const paragraph_1 = __webpack_require__(/*! ../document/paragraph */ "./src/document/paragraph.ts");
const run_1 = __webpack_require__(/*! ../document/run */ "./src/document/run.ts");
function parseNumberingPart(elem, xml) {
    let result = {
        numberings: [],
        abstractNumberings: [],
        bulletPictures: []
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "num":
                result.numberings.push(parseNumbering(e, xml));
                break;
            case "abstractNum":
                result.abstractNumberings.push(parseAbstractNumbering(e, xml));
                break;
            case "numPicBullet":
                result.bulletPictures.push(parseNumberingBulletPicture(e, xml));
                break;
        }
    }
    return result;
}
exports.parseNumberingPart = parseNumberingPart;
function parseNumbering(elem, xml) {
    let result = {
        id: xml.attr(elem, 'numId'),
        overrides: []
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "abstractNumId":
                result.abstractId = xml.attr(e, "val");
                break;
            case "lvlOverride":
                result.overrides.push(parseNumberingLevelOverrride(e, xml));
                break;
        }
    }
    return result;
}
exports.parseNumbering = parseNumbering;
function parseAbstractNumbering(elem, xml) {
    let result = {
        id: xml.attr(elem, 'abstractNumId'),
        levels: []
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "name":
                result.name = xml.attr(e, "val");
                break;
            case "multiLevelType":
                result.multiLevelType = xml.attr(e, "val");
                break;
            case "numStyleLink":
                result.numberingStyleLink = xml.attr(e, "val");
                break;
            case "styleLink":
                result.styleLink = xml.attr(e, "val");
                break;
            case "lvl":
                result.levels.push(parseNumberingLevel(e, xml));
                break;
        }
    }
    return result;
}
exports.parseAbstractNumbering = parseAbstractNumbering;
function parseNumberingLevel(elem, xml) {
    let result = {
        level: xml.intAttr(elem, 'ilvl')
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "start":
                result.start = xml.attr(e, "val");
                break;
            case "lvlRestart":
                result.restart = xml.intAttr(e, "val");
                break;
            case "numFmt":
                result.format = xml.attr(e, "val");
                break;
            case "lvlText":
                result.text = xml.attr(e, "val");
                break;
            case "lvlJc":
                result.justification = xml.attr(e, "val");
                break;
            case "lvlPicBulletId":
                result.bulletPictureId = xml.attr(e, "val");
                break;
            case "pStyle":
                result.paragraphStyle = xml.attr(e, "val");
                break;
            case "pPr":
                result.paragraphProps = (0, paragraph_1.parseParagraphProperties)(e, xml);
                break;
            case "rPr":
                result.runProps = (0, run_1.parseRunProperties)(e, xml);
                break;
        }
    }
    return result;
}
exports.parseNumberingLevel = parseNumberingLevel;
function parseNumberingLevelOverrride(elem, xml) {
    let result = {
        level: xml.intAttr(elem, 'ilvl')
    };
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "startOverride":
                result.start = xml.intAttr(e, "val");
                break;
            case "lvl":
                result.numberingLevel = parseNumberingLevel(e, xml);
                break;
        }
    }
    return result;
}
exports.parseNumberingLevelOverrride = parseNumberingLevelOverrride;
function parseNumberingBulletPicture(elem, xml) {
    var pict = xml.element(elem, "pict");
    var shape = pict && xml.element(pict, "shape");
    var imagedata = shape && xml.element(shape, "imagedata");
    return imagedata ? {
        id: xml.attr(elem, "numPicBulletId"),
        referenceId: xml.attr(imagedata, "id"),
        style: xml.attr(shape, "style")
    } : null;
}
exports.parseNumberingBulletPicture = parseNumberingBulletPicture;


/***/ }),

/***/ "./src/parser/xml-parser.ts":
/*!**********************************!*\
  !*** ./src/parser/xml-parser.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.XmlParser = exports.serializeXmlString = exports.parseXmlString = void 0;
const common_1 = __webpack_require__(/*! ../document/common */ "./src/document/common.ts");
function parseXmlString(xmlString, trimXmlDeclaration = false) {
    if (trimXmlDeclaration)
        xmlString = xmlString.replace(/<[?].*[?]>/, "");
    xmlString = removeUTF8BOM(xmlString);
    const result = new DOMParser().parseFromString(xmlString, "application/xml");
    const errorText = hasXmlParserError(result);
    if (errorText)
        throw new Error(errorText);
    return result;
}
exports.parseXmlString = parseXmlString;
function hasXmlParserError(doc) {
    return doc.getElementsByTagName("parsererror")[0]?.textContent;
}
function removeUTF8BOM(data) {
    return data.charCodeAt(0) === 0xFEFF ? data.substring(1) : data;
}
function serializeXmlString(elem) {
    return new XMLSerializer().serializeToString(elem);
}
exports.serializeXmlString = serializeXmlString;
class XmlParser {
    elements(elem, localName = null) {
        const result = [];
        for (let i = 0, l = elem?.childNodes.length || 0; i < l; i++) {
            let c = elem.childNodes.item(i);
            if (c.nodeType == 1 && (localName == null || c.localName == localName))
                result.push(c);
        }
        return result;
    }
    element(elem, localName) {
        for (let i = 0, l = elem?.childNodes.length || 0; i < l; i++) {
            let c = elem.childNodes.item(i);
            if (c.nodeType == 1 && c.localName == localName)
                return c;
        }
        return null;
    }
    elementAttr(elem, localName, attrLocalName) {
        var el = this.element(elem, localName);
        return el ? this.attr(el, attrLocalName) : undefined;
    }
    attrs(elem) {
        return Array.from(elem.attributes);
    }
    attr(elem, localName) {
        for (let i = 0, l = elem.attributes.length; i < l; i++) {
            let a = elem.attributes.item(i);
            if (a.localName == localName)
                return a.value;
        }
        return null;
    }
    intAttr(node, attrName, defaultValue = null) {
        var val = this.attr(node, attrName);
        return val ? parseInt(val) : defaultValue;
    }
    hexAttr(node, attrName, defaultValue = null) {
        var val = this.attr(node, attrName);
        return val ? parseInt(val, 16) : defaultValue;
    }
    floatAttr(node, attrName, defaultValue = null) {
        var val = this.attr(node, attrName);
        return val ? parseFloat(val) : defaultValue;
    }
    boolAttr(node, attrName, defaultValue = null) {
        return (0, common_1.convertBoolean)(this.attr(node, attrName), defaultValue);
    }
    lengthAttr(node, attrName, usage = common_1.LengthUsage.Dxa) {
        return (0, common_1.convertLength)(this.attr(node, attrName), usage);
    }
}
exports.XmlParser = XmlParser;
const globalXmlParser = new XmlParser();
exports["default"] = globalXmlParser;


/***/ }),

/***/ "./src/settings/settings-part.ts":
/*!***************************************!*\
  !*** ./src/settings/settings-part.ts ***!
  \***************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.SettingsPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const settings_1 = __webpack_require__(/*! ./settings */ "./src/settings/settings.ts");
class SettingsPart extends part_1.Part {
    constructor(pkg, path) {
        super(pkg, path);
    }
    parseXml(root) {
        this.settings = (0, settings_1.parseSettings)(root, this._package.xmlParser);
    }
}
exports.SettingsPart = SettingsPart;


/***/ }),

/***/ "./src/settings/settings.ts":
/*!**********************************!*\
  !*** ./src/settings/settings.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseNoteProperties = exports.parseSettings = void 0;
function parseSettings(elem, xml) {
    var result = {};
    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "defaultTabStop":
                result.defaultTabStop = xml.lengthAttr(el, "val");
                break;
            case "footnotePr":
                result.footnoteProps = parseNoteProperties(el, xml);
                break;
            case "endnotePr":
                result.endnoteProps = parseNoteProperties(el, xml);
                break;
            case "autoHyphenation":
                result.autoHyphenation = xml.boolAttr(el, "val");
                break;
        }
    }
    return result;
}
exports.parseSettings = parseSettings;
function parseNoteProperties(elem, xml) {
    var result = {
        defaultNoteIds: []
    };
    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "numFmt":
                result.nummeringFormat = xml.attr(el, "val");
                break;
            case "footnote":
            case "endnote":
                result.defaultNoteIds.push(xml.attr(el, "id"));
                break;
        }
    }
    return result;
}
exports.parseNoteProperties = parseNoteProperties;


/***/ }),

/***/ "./src/styles/styles-part.ts":
/*!***********************************!*\
  !*** ./src/styles/styles-part.ts ***!
  \***********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.StylesPart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
class StylesPart extends part_1.Part {
    constructor(pkg, path, parser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    parseXml(root) {
        this.styles = this._documentParser.parseStylesFile(root);
    }
}
exports.StylesPart = StylesPart;


/***/ }),

/***/ "./src/theme/theme-part.ts":
/*!*********************************!*\
  !*** ./src/theme/theme-part.ts ***!
  \*********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ThemePart = void 0;
const part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
const theme_1 = __webpack_require__(/*! ./theme */ "./src/theme/theme.ts");
class ThemePart extends part_1.Part {
    constructor(pkg, path) {
        super(pkg, path);
    }
    parseXml(root) {
        this.theme = (0, theme_1.parseTheme)(root, this._package.xmlParser);
    }
}
exports.ThemePart = ThemePart;


/***/ }),

/***/ "./src/theme/theme.ts":
/*!****************************!*\
  !*** ./src/theme/theme.ts ***!
  \****************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseFontInfo = exports.parseFontScheme = exports.parseColorScheme = exports.parseTheme = exports.DmlTheme = void 0;
class DmlTheme {
}
exports.DmlTheme = DmlTheme;
function parseTheme(elem, xml) {
    var result = new DmlTheme();
    var themeElements = xml.element(elem, "themeElements");
    for (let el of xml.elements(themeElements)) {
        switch (el.localName) {
            case "clrScheme":
                result.colorScheme = parseColorScheme(el, xml);
                break;
            case "fontScheme":
                result.fontScheme = parseFontScheme(el, xml);
                break;
        }
    }
    return result;
}
exports.parseTheme = parseTheme;
function parseColorScheme(elem, xml) {
    var result = {
        name: xml.attr(elem, "name"),
        colors: {}
    };
    for (let el of xml.elements(elem)) {
        var srgbClr = xml.element(el, "srgbClr");
        var sysClr = xml.element(el, "sysClr");
        if (srgbClr) {
            result.colors[el.localName] = xml.attr(srgbClr, "val");
        }
        else if (sysClr) {
            result.colors[el.localName] = xml.attr(sysClr, "lastClr");
        }
    }
    return result;
}
exports.parseColorScheme = parseColorScheme;
function parseFontScheme(elem, xml) {
    var result = {
        name: xml.attr(elem, "name"),
    };
    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "majorFont":
                result.majorFont = parseFontInfo(el, xml);
                break;
            case "minorFont":
                result.minorFont = parseFontInfo(el, xml);
                break;
        }
    }
    return result;
}
exports.parseFontScheme = parseFontScheme;
function parseFontInfo(elem, xml) {
    return {
        latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
        eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
        csTypeface: xml.elementAttr(elem, "cs", "typeface"),
    };
}
exports.parseFontInfo = parseFontInfo;


/***/ }),

/***/ "./src/utils.ts":
/*!**********************!*\
  !*** ./src/utils.ts ***!
  \**********************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.asArray = exports.formatCssRules = exports.parseCssRules = exports.mergeDeep = exports.isString = exports.isObject = exports.blobToBase64 = exports.keyBy = exports.resolvePath = exports.splitPath = exports.escapeClassName = void 0;
function escapeClassName(className) {
    return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
}
exports.escapeClassName = escapeClassName;
function splitPath(path) {
    let si = path.lastIndexOf('/') + 1;
    let folder = si == 0 ? "" : path.substring(0, si);
    let fileName = si == 0 ? path : path.substring(si);
    return [folder, fileName];
}
exports.splitPath = splitPath;
function resolvePath(path, base) {
    try {
        const prefix = "http://docx/";
        const url = new URL(path, prefix + base).toString();
        return url.substring(prefix.length);
    }
    catch {
        return `${base}${path}`;
    }
}
exports.resolvePath = resolvePath;
function keyBy(array, by) {
    return array.reduce((a, x) => {
        a[by(x)] = x;
        return a;
    }, {});
}
exports.keyBy = keyBy;
function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = () => reject();
        reader.readAsDataURL(blob);
    });
}
exports.blobToBase64 = blobToBase64;
function isObject(item) {
    return item && typeof item === 'object' && !Array.isArray(item);
}
exports.isObject = isObject;
function isString(item) {
    return typeof item === 'string' || item instanceof String;
}
exports.isString = isString;
function mergeDeep(target, ...sources) {
    if (!sources.length)
        return target;
    const source = sources.shift();
    if (isObject(target) && isObject(source)) {
        for (const key in source) {
            if (isObject(source[key])) {
                const val = target[key] ?? (target[key] = {});
                mergeDeep(val, source[key]);
            }
            else {
                target[key] = source[key];
            }
        }
    }
    return mergeDeep(target, ...sources);
}
exports.mergeDeep = mergeDeep;
function parseCssRules(text) {
    const result = {};
    for (const rule of text.split(';')) {
        const [key, val] = rule.split(':');
        result[key] = val;
    }
    return result;
}
exports.parseCssRules = parseCssRules;
function formatCssRules(style) {
    return Object.entries(style).map((k, v) => `${k}: ${v}`).join(';');
}
exports.formatCssRules = formatCssRules;
function asArray(val) {
    return Array.isArray(val) ? val : [val];
}
exports.asArray = asArray;


/***/ }),

/***/ "./src/vml/vml.ts":
/*!************************!*\
  !*** ./src/vml/vml.ts ***!
  \************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseVmlElement = exports.VmlElement = void 0;
const common_1 = __webpack_require__(/*! ../document/common */ "./src/document/common.ts");
const dom_1 = __webpack_require__(/*! ../document/dom */ "./src/document/dom.ts");
const xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
class VmlElement extends dom_1.OpenXmlElementBase {
    constructor() {
        super(...arguments);
        this.type = dom_1.DomType.VmlElement;
        this.attrs = {};
    }
}
exports.VmlElement = VmlElement;
function parseVmlElement(elem, parser) {
    var result = new VmlElement();
    switch (elem.localName) {
        case "rect":
            result.tagName = "rect";
            Object.assign(result.attrs, { width: '100%', height: '100%' });
            break;
        case "oval":
            result.tagName = "ellipse";
            Object.assign(result.attrs, { cx: "50%", cy: "50%", rx: "50%", ry: "50%" });
            break;
        case "line":
            result.tagName = "line";
            break;
        case "shape":
            result.tagName = "g";
            break;
        case "textbox":
            result.tagName = "foreignObject";
            Object.assign(result.attrs, { width: '100%', height: '100%' });
            break;
        default:
            return null;
    }
    for (const at of xml_parser_1.default.attrs(elem)) {
        switch (at.localName) {
            case "style":
                result.cssStyleText = at.value;
                break;
            case "fillcolor":
                result.attrs.fill = at.value;
                break;
            case "from":
                const [x1, y1] = parsePoint(at.value);
                Object.assign(result.attrs, { x1, y1 });
                break;
            case "to":
                const [x2, y2] = parsePoint(at.value);
                Object.assign(result.attrs, { x2, y2 });
                break;
        }
    }
    for (const el of xml_parser_1.default.elements(elem)) {
        switch (el.localName) {
            case "stroke":
                Object.assign(result.attrs, parseStroke(el));
                break;
            case "fill":
                Object.assign(result.attrs, parseFill(el));
                break;
            case "imagedata":
                result.tagName = "image";
                Object.assign(result.attrs, { width: '100%', height: '100%' });
                result.imageHref = {
                    id: xml_parser_1.default.attr(el, "id"),
                    title: xml_parser_1.default.attr(el, "title"),
                };
                break;
            case "txbxContent":
                result.children.push(...parser.parseBodyElements(el));
                break;
            default:
                const child = parseVmlElement(el, parser);
                child && result.children.push(child);
                break;
        }
    }
    return result;
}
exports.parseVmlElement = parseVmlElement;
function parseStroke(el) {
    return {
        'stroke': xml_parser_1.default.attr(el, "color"),
        'stroke-width': xml_parser_1.default.lengthAttr(el, "weight", common_1.LengthUsage.Emu) ?? '1px'
    };
}
function parseFill(el) {
    return {};
}
function parsePoint(val) {
    return val.split(",");
}
function convertPath(path) {
    return path.replace(/([mlxe])|([-\d]+)|([,])/g, (m) => {
        if (/[-\d]/.test(m))
            return (0, common_1.convertLength)(m, common_1.LengthUsage.VmlEmu);
        if (/[ml,]/.test(m))
            return m;
        return '';
    });
}


/***/ }),

/***/ "./src/word-document.ts":
/*!******************************!*\
  !*** ./src/word-document.ts ***!
  \******************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.deobfuscate = exports.WordDocument = void 0;
const relationship_1 = __webpack_require__(/*! ./common/relationship */ "./src/common/relationship.ts");
const font_table_1 = __webpack_require__(/*! ./font-table/font-table */ "./src/font-table/font-table.ts");
const open_xml_package_1 = __webpack_require__(/*! ./common/open-xml-package */ "./src/common/open-xml-package.ts");
const document_part_1 = __webpack_require__(/*! ./document/document-part */ "./src/document/document-part.ts");
const utils_1 = __webpack_require__(/*! ./utils */ "./src/utils.ts");
const numbering_part_1 = __webpack_require__(/*! ./numbering/numbering-part */ "./src/numbering/numbering-part.ts");
const styles_part_1 = __webpack_require__(/*! ./styles/styles-part */ "./src/styles/styles-part.ts");
const parts_1 = __webpack_require__(/*! ./header-footer/parts */ "./src/header-footer/parts.ts");
const extended_props_part_1 = __webpack_require__(/*! ./document-props/extended-props-part */ "./src/document-props/extended-props-part.ts");
const core_props_part_1 = __webpack_require__(/*! ./document-props/core-props-part */ "./src/document-props/core-props-part.ts");
const theme_part_1 = __webpack_require__(/*! ./theme/theme-part */ "./src/theme/theme-part.ts");
const parts_2 = __webpack_require__(/*! ./notes/parts */ "./src/notes/parts.ts");
const settings_part_1 = __webpack_require__(/*! ./settings/settings-part */ "./src/settings/settings-part.ts");
const custom_props_part_1 = __webpack_require__(/*! ./document-props/custom-props-part */ "./src/document-props/custom-props-part.ts");
const comments_part_1 = __webpack_require__(/*! ./comments/comments-part */ "./src/comments/comments-part.ts");
const commentsExtended_part_1 = __webpack_require__(/*! ./comments/commentsExtended-part */ "./src/comments/commentsExtended-part.ts");
const topLevelRels = [
    { type: relationship_1.RelationshipTypes.OfficeDocument, target: "word/document.xml" },
    { type: relationship_1.RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
    { type: relationship_1.RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
    { type: relationship_1.RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
];
class WordDocument {
    constructor() {
        this.parts = [];
        this.partsMap = {};
    }
    static async load(blob, parser, options) {
        var d = new WordDocument();
        d._options = options;
        d._parser = parser;
        d._package = await open_xml_package_1.OpenXmlPackage.load(blob, options);
        d.rels = await d._package.loadRelationships();
        await Promise.all(topLevelRels.map(rel => {
            const r = d.rels.find(x => x.type === rel.type) ?? rel;
            return d.loadRelationshipPart(r.target, r.type);
        }));
        return d;
    }
    save(type = "blob") {
        return this._package.save(type);
    }
    async loadRelationshipPart(path, type) {
        if (this.partsMap[path])
            return this.partsMap[path];
        if (!this._package.get(path))
            return null;
        let part = null;
        switch (type) {
            case relationship_1.RelationshipTypes.OfficeDocument:
                this.documentPart = part = new document_part_1.DocumentPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.FontTable:
                this.fontTablePart = part = new font_table_1.FontTablePart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.Numbering:
                this.numberingPart = part = new numbering_part_1.NumberingPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Styles:
                this.stylesPart = part = new styles_part_1.StylesPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Theme:
                this.themePart = part = new theme_part_1.ThemePart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.Footnotes:
                this.footnotesPart = part = new parts_2.FootnotesPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Endnotes:
                this.endnotesPart = part = new parts_2.EndnotesPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Footer:
                part = new parts_1.FooterPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Header:
                part = new parts_1.HeaderPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.CoreProperties:
                this.corePropsPart = part = new core_props_part_1.CorePropsPart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.ExtendedProperties:
                this.extendedPropsPart = part = new extended_props_part_1.ExtendedPropsPart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.CustomProperties:
                part = new custom_props_part_1.CustomPropsPart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.Settings:
                this.settingsPart = part = new settings_part_1.SettingsPart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.Comments:
                this.commentsPart = part = new comments_part_1.CommentsPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.CommentsExtended:
                this.commentsExtendedPart = part = new commentsExtended_part_1.CommentsExtendedPart(this._package, path, this._parser);
                break;
        }
        if (part == null)
            return Promise.resolve(null);
        this.partsMap[path] = part;
        this.parts.push(part);
        await part.load();
        if (part.rels?.length > 0) {
            const [folder] = (0, utils_1.splitPath)(part.path);
            await Promise.all(part.rels.map(rel => this.loadRelationshipPart((0, utils_1.resolvePath)(rel.target, folder), rel.type)));
        }
        return part;
    }
    async loadDocumentImage(id, part) {
        const x = await this.loadResource(part ?? this.documentPart, id, "blob");
        return this.blobToURL(x);
    }
    async loadNumberingImage(id) {
        const x = await this.loadResource(this.numberingPart, id, "blob");
        return this.blobToURL(x);
    }
    async loadFont(id, key) {
        const x = await this.loadResource(this.fontTablePart, id, "uint8array");
        return x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x;
    }
    blobToURL(blob) {
        if (!blob)
            return null;
        if (this._options.useBase64URL) {
            return (0, utils_1.blobToBase64)(blob);
        }
        return URL.createObjectURL(blob);
    }
    findPartByRelId(id, basePart = null) {
        var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
        const folder = basePart ? (0, utils_1.splitPath)(basePart.path)[0] : '';
        return rel ? this.partsMap[(0, utils_1.resolvePath)(rel.target, folder)] : null;
    }
    getPathById(part, id) {
        const rel = part.rels.find(x => x.id == id);
        const [folder] = (0, utils_1.splitPath)(part.path);
        return rel ? (0, utils_1.resolvePath)(rel.target, folder) : null;
    }
    loadResource(part, id, outputType) {
        const path = this.getPathById(part, id);
        return path ? this._package.load(path, outputType) : Promise.resolve(null);
    }
}
exports.WordDocument = WordDocument;
function deobfuscate(data, guidKey) {
    const len = 16;
    const trimmed = guidKey.replace(/{|}|-/g, "");
    const numbers = new Array(len);
    for (let i = 0; i < len; i++)
        numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);
    for (let i = 0; i < 32; i++)
        data[i] = data[i] ^ numbers[i % len];
    return data;
}
exports.deobfuscate = deobfuscate;


/***/ }),

/***/ "jszip":
/*!*******************************************************************************************************!*\
  !*** external {"root":"JSZip","commonjs":"jszip","commonjs2":"jszip","amd":"jszip","module":"jszip"} ***!
  \*******************************************************************************************************/
/***/ ((module) => {

var x = y => { var x = {}; __webpack_require__.d(x, y); return x; }
var y = x => () => x
module.exports = __WEBPACK_EXTERNAL_MODULE_jszip__;

/***/ })

/******/ });
/************************************************************************/
/******/ // The module cache
/******/ var __webpack_module_cache__ = {};
/******/ 
/******/ // The require function
/******/ function __webpack_require__(moduleId) {
/******/ 	// Check if module is in cache
/******/ 	var cachedModule = __webpack_module_cache__[moduleId];
/******/ 	if (cachedModule !== undefined) {
/******/ 		return cachedModule.exports;
/******/ 	}
/******/ 	// Create a new module (and put it into the cache)
/******/ 	var module = __webpack_module_cache__[moduleId] = {
/******/ 		// no module.id needed
/******/ 		// no module.loaded needed
/******/ 		exports: {}
/******/ 	};
/******/ 
/******/ 	// Execute the module function
/******/ 	__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 
/******/ 	// Return the exports of the module
/******/ 	return module.exports;
/******/ }
/******/ 
/************************************************************************/
/******/ 
/******/ // startup
/******/ // Load entry module and return exports
/******/ // This entry module is referenced by other modules so it can't be inlined
/******/ var __webpack_exports__ = __webpack_require__("./src/cx-docx.ts");
/******/ var __webpack_exports___esModule = __webpack_exports__.__esModule;
/******/ var __webpack_exports__defaultOptions = __webpack_exports__.defaultOptions;
/******/ var __webpack_exports__praseAsync = __webpack_exports__.praseAsync;
/******/ var __webpack_exports__renderAsync = __webpack_exports__.renderAsync;
/******/ export { __webpack_exports___esModule as __esModule, __webpack_exports__defaultOptions as defaultOptions, __webpack_exports__praseAsync as praseAsync, __webpack_exports__renderAsync as renderAsync };
/******/ 

//# sourceMappingURL=cx-docx.js.map