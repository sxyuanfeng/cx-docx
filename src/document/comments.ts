import { XmlParser } from "../parser/xml-parser.ts";
import { DomType, OpenXmlElement } from "./dom.ts";

export interface WmlCommentReference extends OpenXmlElement {
  id: string
}

export interface WmlCommentRangeStart extends OpenXmlElement {
  id: string
}

export interface WmlCommentRangeEnd extends OpenXmlElement {
  id: string
}

export function parseCommentRangeStart(elem: Element, xml: XmlParser): WmlCommentRangeStart {
  return {
      type: DomType.CommentRangeStart,
      id: xml.attr(elem, "id"),
  }
}

export function parseCommentRangeEnd(elem: Element, xml: XmlParser): WmlCommentRangeEnd {
  return {
      type: DomType.CommentRangeEnd,
      id: xml.attr(elem, "id")
  }
}
