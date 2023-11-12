import { OpenXmlPackage } from "../common/open-xml-package.ts";
import { Part } from "../common/part.ts";
import { DocumentParser } from "../document-parser.ts";
import { IComment } from "../document/dom.ts";
import {CommentsPartProperties, parseCommentsPart } from "./comments.ts";

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
