import { ICommentExtended } from "../document/dom";
import { XmlParser } from "../parser/xml-parser";

export interface CommentsExtendedPartProperties {
    commentsEx: ICommentExtended[];
}

export function parseCommentsExtendedPart(elem: Element, xml: XmlParser): CommentsExtendedPartProperties {
    let result: CommentsExtendedPartProperties = {
        commentsEx: []
    }
    
    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "commentEx":
                result.commentsEx.push(parseCommentEx(e, xml));
                break;

        }
    }

    return result;
}

export function parseCommentEx(elem: Element, xml: XmlParser): ICommentExtended {

  let result: ICommentExtended = {
    paraIdParent: xml.attr(elem, "paraIdParent"),
    paraId: xml.attr(elem, "paraId")
  }

  return result;
}
