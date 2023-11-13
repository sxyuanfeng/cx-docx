import { OpenXmlElementBase, DomType, OpenXmlElement } from "../document/dom";

export abstract class WmlBaseNote implements OpenXmlElementBase {
    type: DomType;
    id: string;
	noteType: string;
	children?: OpenXmlElement[];
}

export class WmlFootnote extends WmlBaseNote {
	type = DomType.Footnote
}

export class WmlEndnote extends WmlBaseNote {
	type = DomType.Endnote
}