import { OpenXmlPackage } from "../common/open-xml-package.ts";
import { Part } from "../common/part.ts";
import { DocumentParser } from "../document-parser.ts";
import { ICommentExtended } from "../document/dom.ts";
import {CommentsExtendedPartProperties, parseCommentsExtendedPart } from "./commentsExtended.ts";

export class CommentsExtendedPart extends Part implements CommentsExtendedPartProperties {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    commentsEx: ICommentExtended[];

    parseXml(root: Element) {
        Object.assign(this, parseCommentsExtendedPart(root, this._package.xmlParser));
    }
}
