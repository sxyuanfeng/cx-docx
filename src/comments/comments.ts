import { IComment } from "../document/dom";
import { XmlParser } from "../parser/xml-parser";

export interface CommentsPartProperties {
    comments: IComment[];
}

export function parseCommentsPart(elem: Element, xml: XmlParser): CommentsPartProperties {
    let result: CommentsPartProperties = {
        comments: []
    }
    
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "comment":
                result.comments.push(parseComment(e, xml));
                break;

        }
    }

    return result;
}

export function parseComment(elem: Element, xml: XmlParser): IComment {

  let result: IComment = {
    id: xml.attr(elem, "id"),
    author: xml.attr(elem, "author"),
    date: xml.attr(elem, "date"),
    paraId: xml.elementAttr(elem, "p", "paraId"),
    text: ParseCommentText(xml.element(elem, "p"), xml),
    noRender: false,
    msg: '',
    type: '',
    children: []
  }

  return result;
}

export function ParseCommentText(elem: Element, xml: XmlParser): string {
  let result = '';

  for (let e of xml.elements(elem, "r")) {
    for (let t of xml.elements(e, "t")) {
      result += t.textContent;
    }

  }

  return result;
}
