import { XmlParser } from "../parser/xml-parser";
import { DomType, OpenXmlElement } from "./dom";

export interface WmlBookmarkStart extends OpenXmlElement {
    id: string;
    name: string;
    colFirst: number;
    colLast: number;
    displacedByCustomXml: string;
}

export interface WmlBookmarkEnd extends OpenXmlElement {
    id: string;
}

export function parseBookmarkStart(elem: Element, xml: XmlParser): WmlBookmarkStart {
    return {
        type: DomType.BookmarkStart,
        id: xml.attr(elem, "id"),
        name: xml.attr(elem, "name"),
        colFirst: xml.intAttr(elem, "colFirst"),
        colLast: xml.intAttr(elem, "colLast"),
        displacedByCustomXml: xml.attr(elem, "displacedByCustomXml"),
    }
}

export function parseBookmarkEnd(elem: Element, xml: XmlParser): WmlBookmarkEnd {
    return {
        type: DomType.BookmarkEnd,
        id: xml.attr(elem, "id")
    }
}