import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { ICommentExtended } from "../document/dom";
import {CommentsExtendedPartProperties, parseCommentsExtendedPart } from "./commentsExtended";

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
