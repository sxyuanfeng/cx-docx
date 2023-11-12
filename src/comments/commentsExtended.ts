/*
 * @Author: xujiang
 * @Date: 2023-11-06 10:26:54
 * @LastEditors: xujiang
 * Copyright (c) 2023 by xujiang/cicc, All Rights Reserved.
 */
import { ICommentExtended } from "../document/dom.ts";
import { XmlParser } from "../parser/xml-parser.ts";

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
