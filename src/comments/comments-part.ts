import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IComment } from "../document/dom";
import {CommentsPartProperties, parseCommentsPart } from "./comments";

export class CommentsPart extends Part implements CommentsPartProperties {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    comments: IComment[];

    parseXml(root: Element) {
        Object.assign(this, parseCommentsPart(root, this._package.xmlParser));
    }
}
