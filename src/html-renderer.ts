import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference, IComment
} from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './cx-docx';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, escapeClassName, isString, keyBy, mergeDeep } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { WmlRun, RunProperties } from './document/run';
import { WmlBookmarkEnd, WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { createIconCollapse, createIconPrev } from './assets/index';
import { WmlCommentRangeEnd, WmlCommentRangeStart } from './document/comments';

const ns = {
	svg: "http://www.w3.org/2000/svg",
	mathML: "http://www.w3.org/1998/Math/MathML"
}

interface CellPos {
	col: number;
	row: number;
}

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

export class HtmlRenderer {

	className: string = "docx";
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	defaultTabSize: string;
	currentTabs: any[] = [];
	tabsTimeout: any = 0;

	currentComment: string = '';
	allComments: string[] = [];

	flatedDocument: [];

	commentIndex: number;

	commentTranslateY: number;

	constructor(public htmlDocument: Document) {
	}

	// 解析part中的文本、父子关系、作者等信息
	processCommentPart(document: WordDocument, id: string, commentReference) {
		let comments = document.commentsPart.comments;
		let commentsEx = document.commentsExtendedPart.commentsEx;
		let result = Object.assign(comments.find(item => {
			return item.id === id;
		}), commentReference);
		result.children = [];
		let children = commentsEx.filter(item => {
			return item.paraIdParent === result?.paraId;
		})
		for (let child of children) {
			let comment = comments.find(item => {
				return item.paraId === child.paraId;
			})
			result.children.push(comment);
		}
		if (commentsEx.find(item => {return item.paraId === result.paraId})?.paraIdParent) {
			result.noRender = true;
		}
		
		return result;
	}

	// 将分割在不通段落、连续块中的划线文本进行拼接
	joinRangeText(rArr): string {
		let result = '';
		for (let i = 0; i < rArr.length; i++) {
			let rChildren = rArr[i]?.children || [];
			for (let child of rChildren) {
				if (child && child?.type === DomType.Text) {
					result += child.text;
				}
			}
		}
		return result;
	}

	flatChildren(document: WordDocument, flatedDocument: []) {
		let body = document.documentPart.body;
		function myFlat(elem, flatedDocument) {
			flatedDocument.push(elem);
			if (elem.children && elem.children.length > 0) {
				for (let child of elem.children) {
					myFlat(child, flatedDocument);
				}
			}
		}
		myFlat(body, flatedDocument);
	}

	render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.styleMap = null;
		this.flatedDocument = [];
		this.commentIndex = 1;
		this.commentTranslateY = 0;

		styleContainer = styleContainer || bodyContainer;

		this.flatChildren(document, this.flatedDocument);

		removeAllElements(styleContainer);
		removeAllElements(bodyContainer);

		this.processAllComments();

		appendComment(styleContainer, "docxjs library predefined styles");
		styleContainer.appendChild(this.renderDefaultStyle());

		if (document.themePart) {
			appendComment(styleContainer, "docxjs document theme values");
			this.renderTheme(document.themePart, styleContainer);
		}

		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			appendComment(styleContainer, "docxjs document styles");
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}

		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);

			appendComment(styleContainer, "docxjs document numbering styles");
			// styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			styleContainer.appendChild(this.renderTitleNumbering(document.documentPart.body.children, document.numberingPart.domNumberings, styleContainer));
			// styleContainer.appendChild(this.renderMyNumbering(document.numberingPart.domNumberings, document.numberingPart.numberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}

		if (document.footnotesPart) {
			this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
		}

		if (document.endnotesPart) {
			this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
		}

		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}

		if (!options.ignoreFonts && document.fontTablePart)
			this.renderFontTable(document.fontTablePart, styleContainer);

		var sectionElements = this.renderSections(document.documentPart.body);

		if (this.options.inWrapper) {
			// if (this.options.renderOutline || this.options.renderComments) {
			// 	bodyContainer.appendChild(this.renderOutlineWrapper(this.renderWrapper(sectionElements)));
			// } else {
			// 	bodyContainer.appendChild(this.renderWrapper(sectionElements));
			// }
			bodyContainer.appendChild(this.renderOutlineWrapper(this.renderWrapper(sectionElements)));
		} else {
			appendChildren(bodyContainer, sectionElements);
		}

		this.refreshTabStops();

		(this.options.inWrapper && this.options.renderComments) && this.renderCommentElement();
	}

	processAllComments() {
		let commentsEx = this.document.commentsExtendedPart?.commentsEx;
		let comments = this.document.commentsPart?.comments;
		if (comments && commentsEx) {
			for (let commentEx of commentsEx) {
				if (commentEx.paraIdParent === null) {
					this.allComments.push(commentEx.paraId);
				}
			}
		}
	}

	renderCommentElement() {
		if (this.document.commentsPart?.comments) {
			let wrap = document.getElementsByClassName(`${this.className}-comment-wrap`)[0];
			for (let paraId of this.allComments) {
				let commentSupElem = document.getElementById(`${this.className}-comment-start-${paraId}`);
				let rect = commentSupElem?.getBoundingClientRect();
				let commentElem = this.createElement("div");
				commentElem.id = `${this.className}-comment-element-${paraId}`;
				commentElem.className = `${this.className}-comment-content-wrap`;
				let that = this;
				commentElem.addEventListener("mouseenter", e => {
					for (let paraId of that.allComments) {
						that.inactiveCommentElement(paraId);
					}
					that.activeCommentElement(paraId);
					that.currentComment = paraId;
					that.setCurrentCommentText();
				})
				commentElem.addEventListener("mouseleave", e => {
					for (let paraId of that.allComments) {
						that.inactiveCommentElement(paraId);
					}
				})
				let commentContent = this.renderCommentContent(paraId);
				if (commentContent !== null) {
					commentElem.appendChild(commentContent);
				}
				wrap.appendChild(commentElem);
				if (rect) {
					if (rect.y > this.commentTranslateY) {
						commentElem.style.top = `${rect.y}px`;
					} else {
						commentElem.style.top = `${this.commentTranslateY}px`;
					}
					let elemRect = commentElem.getBoundingClientRect();
					this.commentTranslateY = elemRect.height + elemRect.y + 10;
				}
			}
		}
	}

	activeCommentElement(paraId) {
		let commentElem = document.getElementById(`${this.className}-comment-element-${paraId}`);
		let commentStart = document.getElementById(`${this.className}-comment-start-${paraId}`);
		let commentEnd = document.getElementById(`${this.className}-comment-end-${paraId}`);
		if (commentStart) {
			commentStart.style.borderRight = "4px solid rgb(252, 13, 27)";
			commentStart.style.marginRight = "3px";
		}
		if (commentEnd) {
			commentEnd.style.borderLeft = "4px solid rgb(252, 13, 27)";
			commentEnd.style.marginLeft = "3px";
			commentEnd.classList.add("active");
		}
		if (commentElem) {
			commentElem.style.borderLeft = "5px solid rgb(252, 13, 27)";
			commentElem.style.boxShadow = "3px 3px 6px #ccc";
			commentElem.style.left = "-20px";
			commentElem.classList.add("active");
		}
	}

	inactiveCommentElement(paraId) {
		let commentElem = document.getElementById(`${this.className}-comment-element-${paraId}`);
		let commentStart = document.getElementById(`${this.className}-comment-start-${paraId}`);
		let commentEnd = document.getElementById(`${this.className}-comment-end-${paraId}`);
		if (commentStart) {
			commentStart.style.borderRight = "none";
			commentStart.style.marginRight = "0";
		}
		if (commentEnd) {
			commentEnd.style.borderLeft = "none";
			commentEnd.style.marginLeft = "0";
			commentEnd.classList.remove("active");
		}
		if (commentElem) {
			commentElem.style.borderLeft = "none";
			commentElem.style.boxShadow = "none";
			commentElem.style.left = "0";
			commentElem.classList.remove("active");
		}
	}

	renderCommentContent(paraId): Node | null {
		let comments = this.document.commentsPart?.comments || [];
		let commentsEx = this.document.commentsExtendedPart?.commentsEx || [];
		let comment = comments.find(item => {
			return item.paraId === paraId;
		}) as IComment;
		let id = comment?.id;
		let msg = '';
		let commentRangeStartIndex = this.flatedDocument.findIndex(item => {
			return (item as OpenXmlElement).type === "commentRangeStart" && (item as WmlCommentRangeStart).id === id;
		});
		for (let i = commentRangeStartIndex; i < this.flatedDocument.length; i++) {
			if ((this.flatedDocument[i] as OpenXmlElement).type === "commentRangeEnd" && (this.flatedDocument[i] as WmlCommentRangeEnd).id === id) {
				break;
			}
			if (this.flatedDocument[i] && (this.flatedDocument[i] as OpenXmlElement).type === DomType.Text) {
				msg = msg + (this.flatedDocument[i] as WmlText).text;
			}
		}
		if (comment) {
			comment.msg = msg;
			for (let commentEx of commentsEx) {
				if (commentEx.paraIdParent === paraId) {
					let child = comments.find(item => {
						return item.paraId === commentEx.paraId;
					});
					if (child) {
						comment.children.push(child);
					}
				}
			}
		}
		return this.createCommentContentNode(comment);
	}

	// 渲染批注内容
	createCommentContentNode(elem: IComment): Node {
		let commentsContainer = this.createElement("div");
		let commentContainer = this.createCommentNode(elem);
		if (commentContainer) {
			commentsContainer.appendChild(commentContainer);
		}

		return commentsContainer;
	}

	// 渲染单个批注
	createCommentNode(elem: IComment): Node | null {
		if (!elem) {return null;}
		let commentContainer = this.createElement("div");
		let author = this.createElement("span");
		author.style.marginRight = "10px";
		author.style.fontSize = "13px";
		author.textContent = elem.author;
		let date = this.createElement("span");
		date.style.fontSize = "13px";
		date.textContent = elem.date?.replace('T', ' ')?.replace('Z', '');
		let text = this.createElement("div");
		text.style.width = "270px";
		text.style.margin = "3px 0 5px";
		text.style.fontSize = "14px";
		text.textContent = elem.text;

		commentContainer.appendChild(author);
		commentContainer.appendChild(date);
		commentContainer.appendChild(text);

		if (elem?.children?.length > 0) {
			let childCommentContainer = this.createElement("div");
			childCommentContainer.style.borderLeft = "2px solid #bbbfc4";
			childCommentContainer.style.paddingLeft = "7px";
			childCommentContainer.style.marginLeft = "20px";
			for (let i = 0; i < elem?.children?.length; i++) {
				childCommentContainer.appendChild(this.createCommentNode(elem?.children?.[i]));
			}
			commentContainer.appendChild(childCommentContainer);
		}

		return commentContainer;
	}

	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (let [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		styleContainer.appendChild(createStyleElement(cssText));
	}

	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				this.document.loadFont(ref.id, ref.key).then(fontData => {
					const cssValues = {
						'font-family': f.name,
						'src': `url(${fontData})`
					};

					if (ref.type == "bold" || ref.type == "boldItalic") {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == "italic" || ref.type == "boldItalic") {
						cssValues['font-style'] = 'italic';
					}

					appendComment(styleContainer, `docxjs ${f.name} font`);
					const cssText = this.styleToString("@font-face", cssValues);
					styleContainer.appendChild(createStyleElement(cssText));
					this.refreshTabStops();
				});
			}
		}
	}

	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	processStyles(styles: IDomStyle[]) {
		const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);

		for (const style of styles.filter(x => x.basedOn)) {
			var baseStyle = stylesMap[style.basedOn];

			if (baseStyle) {
				style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
				style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

				for (const baseValues of baseStyle.styles) {
					const styleValues = style.styles.find(x => x.target == baseValues.target);

					if (styleValues) {
						this.copyStyleProperties(baseValues.values, styleValues.values);
					} else {
						style.styles.push({ ...baseValues, values: { ...baseValues.values } });
					}
				}
			}
			else if (this.options.debug)
				console.warn(`Can't find base style ${style.basedOn}`);
		}

		for (let style of styles) {
			style.cssName = this.processStyleName(style.id);
		}

		return stylesMap;
	}

	prodessNumberings(numberings: IDomNumbering[]) {
		for (let num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (var e of element.children) {
				e.parent = element;

				if (e.type == DomType.Table) {
					this.processTable(e);
				}
				else {
					this.processElement(e);
				}
			}
		}
	}

	processTable(table: WmlTable) {
		for (var r of table.children) {
			for (var c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left", "border-right", "border-top", "border-bottom",
					"padding-left", "padding-right", "padding-top", "padding-bottom"
				]);

				this.processElement(c);
			}
		}
	}

	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input)
			return output;

		if (output == null) output = {};
		if (attrs == null) attrs = Object.getOwnPropertyNames(input);

		for (var key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	createSection(className: string, props: SectionProperties) {
		var elem = this.createElement("section", { className });

		if (props) {
			if (props.pageMargins) {
				elem.style.paddingLeft = props.pageMargins.left;
				elem.style.paddingRight = props.pageMargins.right;
				elem.style.paddingTop = props.pageMargins.top;
				elem.style.paddingBottom = props.pageMargins.bottom;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth)
					elem.style.width = props.pageSize.width;
				if (!this.options.ignoreHeight)
					elem.style.minHeight = props.pageSize.height;
			}

			if (props.columns && props.columns.numberOfColumns) {
				elem.style.columnCount = `${props.columns.numberOfColumns}`;
				elem.style.columnGap = props.columns.space;

				if (props.columns.separator) {
					elem.style.columnRule = "1px solid black";
				}
			}
		}

		return elem;
	}

	renderSections(document: DocumentElement): HTMLElement[] {
		const result = [];

		this.processElement(document);
		const sections = this.splitBySection(document.children);
		let prevProps = null;

		for (let i = 0, l = sections.length; i < l; i++) {
			this.currentFootnoteIds = [];

			const section = sections[i];
			const props = section.sectProps || document.props;
			const sectionElement = this.createSection(this.className, props);
			this.renderStyleValues(document.cssStyle, sectionElement);

			this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props,
				result.length, prevProps != props, sectionElement);

			var contentElement = this.createElement("article");
			this.renderElements(section.elements, contentElement);
			sectionElement.appendChild(contentElement);

			if (this.options.renderFootnotes) {
				this.renderNotes(this.currentFootnoteIds, this.footnoteMap, sectionElement);
			}

			if (this.options.renderEndnotes && i == l - 1) {
				this.renderNotes(this.currentEndnoteIds, this.endnoteMap, sectionElement);
			}

			this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props,
				result.length, prevProps != props, sectionElement);

			result.push(sectionElement);
			prevProps = props;
		}

		return result;
	}

	renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
		if (!refs) return;

		var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
			?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
			?? refs.find(x => x.type == "default");

		var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart;

		if (part) {
			this.currentPart = part;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			this.renderElements([part.rootElement], into);
			this.currentPart = null;
		}
	}

	isPageBreakElement(elem: OpenXmlElement): boolean {
		if (elem.type != DomType.Break)
			return false;

		if ((elem as WmlBreak).break == "lastRenderedPageBreak")
			return !this.options.ignoreLastRenderedPageBreak;

		return (elem as WmlBreak).break == "page";
	}

	splitBySection(elements): { sectProps: SectionProperties, elements: OpenXmlElement[] }[] {
		var current = { sectProps: null, elements: [] };
		var result = [current];

		for (let elem of elements) {
			if (elem.type == DomType.Paragraph) {
				const s = this.findStyle((elem as WmlParagraph).styleName);

				if (s?.paragraphProps?.pageBreakBefore) {
					current.sectProps = sectProps;
					current = { sectProps: null, elements: [] };
					result.push(current);
				}
			}

			current.elements.push(elem);

			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;

				var sectProps = p.sectionProps;
				var pBreakIndex = -1;
				var rBreakIndex = -1;

				if (this.options.breakPages && p.children) {
					pBreakIndex = p.children.findIndex(r => {
						rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
						return rBreakIndex != -1;
					});
				}

				if (sectProps || pBreakIndex != -1) {
					current.sectProps = sectProps;
					current = { sectProps: null, elements: [] };
					result.push(current);
				}

				if (pBreakIndex != -1 && pBreakIndex != 0) {
					let breakRun = p.children[pBreakIndex];
					// let splitRun = rBreakIndex < breakRun.children.length - 1;
					let splitRun = false;

					if (pBreakIndex < p.children.length - 1 || splitRun) {
						var children = elem.children;
						var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
						elem.children = children.slice(0, pBreakIndex);
						newParagraph.noRenderNumbering = true;
						current.elements.push(newParagraph);

						if (splitRun) {
							let runChildren = breakRun.children;
							let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
							elem.children.push(newRun);
							breakRun.children = runChildren.slice(rBreakIndex);
							console.log(breakRun)
						}
					}
				}
			}
		}

		let currentSectProps = null;

		for (let i = result.length - 1; i >= 0; i--) {
			if (result[i].sectProps == null) {
				result[i].sectProps = currentSectProps;
			} else {
				currentSectProps = result[i].sectProps
			}
		}

		return result;
	}

	renderWrapper(children: HTMLElement[]) {
		return this.createElement("div", { className: `${this.className}-wrapper` }, children);
	}

	renderOutlineWrapper(child: HTMLElement): HTMLElement {
		let outline = this.createOutlineElement();
		let documentContainer = this.createElement("div", { className: `${this.className}-document-container-wrapper`});
		documentContainer.appendChild(child);
		let documentCommentContainer = this.createElement("div", { className: `${this.className}-document-comment-container-wrapper`});
		documentCommentContainer.appendChild(documentContainer);
		let collapseBtn = this.createElement("div", { className: `${this.className}-collapse-btn` });
		collapseBtn.appendChild(createIconCollapse());
		collapseBtn.addEventListener("click", function() {
			if (!outline.className.includes("close")) {
				outline.classList.add("close");
				collapseBtn.classList.add("close");
			} else {
				outline.classList.remove("close");
				collapseBtn.classList.remove("close");
			}
		})
		let collapseBtnWrapper = this.createElement("div", { className: `${this.className}-collapse-btn-wrapper`});
		collapseBtnWrapper.appendChild(collapseBtn);

		let commentControler = this.createElement("div", { className: `${this.className}-comment-controler`});
		let currentCommentElem = this.createElement("span", { className: `${this.className}-current-comment`});
		currentCommentElem.id = `${this.className}-comment-current-index`;
		this.setCurrentCommentText(currentCommentElem);
		let allCommentElem = this.createElement("span", { className: `${this.className}-all-comment`});
		this.setAllCommentText(allCommentElem);

		let that = this;
		let prevComment = this.createElement("span", { className: `${this.className}-prev-comment`});
		prevComment.appendChild(createIconPrev());
		prevComment.title = "ctrl + ↑";
		prevComment.addEventListener("click", function(e) {
			e.stopPropagation();
			that.handleClickPrevComment();
		})

		document.addEventListener("keydown", function(e) {
			if (e.keyCode === 38 && (navigator.platform.match("Mac") ? e.metaKey : e.ctrlKey)) {
				e.preventDefault();
				that.handleClickPrevComment();
			}
		})

		let nextComment = this.createElement("span", { className: `${this.className}-next-comment`});
		nextComment.appendChild(createIconPrev());
		nextComment.title = "ctrl + ↓";
		nextComment.addEventListener("click", function(e) {
			e.stopPropagation();
			that.handleClickNextComment();
		})

		document.addEventListener("keydown", function(e) {
			if (e.keyCode === 40 && (navigator.platform.match("Mac") ? e.metaKey : e.ctrlKey)) {
				e.preventDefault();
				that.handleClickNextComment();
			}
		})
		
		commentControler.appendChild(currentCommentElem);
		commentControler.appendChild(allCommentElem);
		commentControler.appendChild(prevComment);
		commentControler.appendChild(nextComment);

		let commentWrap = this.createElement("div", { className: `${this.className}-comment-wrap`});
		commentWrap.appendChild(commentControler);
		
		// let collapseCommentBtn = this.createElement("div", { className: `${this.className}-collapse-comment-btn` });
		// collapseCommentBtn.appendChild(createIconCollapse());
		// collapseCommentBtn.addEventListener("click", function() {
		// 	if (!commentWrap.className.includes("close")) {
		// 		commentWrap.classList.add("close");
		// 		collapseCommentBtn.classList.add("close");
		// 	} else {
		// 		commentWrap.classList.remove("close");
		// 		collapseCommentBtn.classList.remove("close");
		// 	}
		// })
		// let collapseCommentBtnWrapper = this.createElement("div", { className: `${this.className}-collapse-comment-btn-wrapper`});
		// collapseCommentBtnWrapper.appendChild(collapseCommentBtn);

		let renderers = [documentCommentContainer];
		if (this.options.renderOutline) {
			renderers.unshift(collapseBtnWrapper)
			renderers.unshift(outline as HTMLDivElement);
		}
		if (this.options.renderComments) {
			// documentCommentContainer.appendChild(collapseCommentBtnWrapper);
			documentCommentContainer.appendChild(commentWrap);
		}
		let outlineWrapper = this.createElement("div", { className: `${this.className}-outline-wrapper` }, renderers);

		return outlineWrapper;
	}

	handleClickPrevComment() {
		let index = this.allComments.findIndex(item => {return item === this.currentComment;});
		if (index > 0) {
			this.currentComment = this.allComments[index - 1];
		} else if (index === 0) {
			this.currentComment = this.allComments[this.allComments.length - 1];
		} else {
			return;
		}
		let wrap = document.getElementsByClassName(`${this.className}-comment-wrap`)[0];
		let anchor = wrap.querySelector(`#${this.className}-comment-element-${this.currentComment}`) as HTMLElement;
		anchor?.scrollIntoView({ behavior: "auto", block: "center", inline: "nearest" });
		anchor.click();
		this.setCurrentCommentText();
		for (let paraId of this.allComments) {
			this.inactiveCommentElement(paraId);
		}
		this.activeCommentElement(this.currentComment);
	}

	handleClickNextComment() {
		let index = this.allComments.findIndex(item => {return item === this.currentComment;});
		if (index < this.allComments.length - 1 && this.allComments.length > 0) {
			this.currentComment = this.allComments[index + 1];
		} else if (index === this.allComments.length - 1  && this.allComments.length > 0) {
			this.currentComment = this.allComments[0];
		} else {
			return;
		}
		let wrap = document.getElementsByClassName(`${this.className}-comment-wrap`)[0];
		let anchor = wrap.querySelector(`#${this.className}-comment-element-${this.currentComment}`) as HTMLElement;
		anchor?.scrollIntoView({ behavior: "auto", block: "center", inline: "nearest" });
		anchor.click();
		this.setCurrentCommentText();
		for (let paraId of this.allComments) {
			this.inactiveCommentElement(paraId);
		}
		this.activeCommentElement(this.currentComment);
	}

	setCurrentCommentText(currentElem?) {
		let elem = document.getElementById(`${this.className}-comment-current-index`);
		let index = this.allComments.findIndex(item => {return item === this.currentComment;})
		if (currentElem) {
			currentElem.textContent = `评论（${index + 1}/`;
		} else {
			if (elem) {
				elem.textContent = `评论（${index + 1}/`;
			}
		}
	}

	setAllCommentText(elem) {
		elem.textContent = `${this.allComments.length}）`;
	}

	createOutlineElement(): HTMLElement {
		let outlineContainer = this.createElement("div", { className: `${this.className}-outline-container`});
		let outlineContent = this.createElement("div", { className: `${this.className}-outline-content`});

		let pArr = this.document.documentPart.body.children;
		for (let p of pArr) {
			if (p.type === DomType.Paragraph) {
				let pChildren = p.children ?? [];
				for (let i = 0; i < pChildren.length; i++) {
					let child = pChildren[i];
					if (child.type === DomType.BookmarkStart && (child as WmlBookmarkStart).name?.startsWith("_Toc") && !(child as WmlBookmarkStart).displacedByCustomXml) {
						let endIndex = pChildren.findIndex(item => {
							return item.type === DomType.BookmarkEnd && (item as WmlBookmarkEnd).id === (child as WmlBookmarkStart).id;
						})
						if (endIndex > -1) {
							let title = this.joinRangeText(pChildren);
							if (title.length > 30) {
								break;
							}
							let result = this.createElement("p");
							result.style.marginLeft = `${16}px`;
							result.style.fontSize = '14px';
							result.style.lineHeight = '20px';
							result.textContent = title;
							let that = this;
							result.addEventListener("click", function() {
								let wrap = document.getElementsByClassName(`${that.className}-wrapper`)[0];
								let anchor = wrap.querySelector(`#${(child as WmlBookmarkStart).name}`);
								anchor?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
							})
							outlineContent.appendChild(result);
							break;
						}
					}
				}
			}
		}
		outlineContainer.appendChild(outlineContent);
		return outlineContainer;
	}

	renderDefaultStyle() {
		var c = this.className;
		var styleText = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; flex: 1; } 
.${c}-outline-wrapper { display: flex; height: 100%; }
.${c}-document-comment-container-wrapper { flex-grow: 1; height: 100%; overflow: auto; display: flex; background: gray; counter-reset: ${c}-comment-range-end-counter 0; }
.${c}-document-container-wrapper { flex-grow: 1; }
.${c}-collapse-comment-btn-wrapper { position: sticky; top: 0; }
.${c}-comment-wrap { position: relative; width: 340px; counter-reset: ${c}-comment-counter 0;}
.${c}-comment-range-end { counter-increment: ${c}-comment-range-end-counter; }
.${c}-comment-range-end.active::after { content: ""counter(${c}-comment-range-end-counter, decimal)""; color: rgb(252, 13, 27); font-size: 16px; font-weight: 600; margin-left: 3px;}
.${c}-comment-content-wrap { counter-increment: ${c}-comment-counter; position: absolute; left: 0; width: 300px; padding: 15px; margin-right: 5px; z-index: 998; background-color: #fff; border-radius: 5px; text-align: start; cursor: pointer;}
.${c}-comment-content-wrap.active::before {content: ""counter(${c}-comment-counter, decimal)""; position: absolute; left: -26px; top: -2px; color: #fff; font-weight: 600; font-size: 20px;}
.${c}-comment-controler { position: fixed; top: 5px; right: 25px; background: #fff; min-width: 200px; 
	height: 35px; display: flex; align-items: center; border-radius: 4px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); z-index: 999; }
.${c}-outline-container { background: white; width: 300px; height: 100%; overflow: auto; transition: width .5s ease; }
.${c}-outline-container.close { width: 0 }
.${c}-collapse-btn-wrapper { width: 1.5em; height: 100%; background: gray; }
.${c}-collapse-btn { transform: rotate(90deg); width: 1.5em; height: 1.5em; background: gray; object-fit: contain; cursor: pointer; transition: transform .8s ease-out; }
.${c}-collapse-btn.close { transform: rotate(0deg); transition: transform .8s ease-out; }
.${c}-outline-content { background: white; }
.${c}-outline-content > p { text-align: left; margin-top: 3px; margin-bottom: 3px; margin-right: 30px; cursor: pointer; }
.${c}-outline-content > p span { font-size: 12px !important; color: #333; }
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
.${c} { color: black; hyphens: auto; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c}-prev-comment { margin-left: 10px; transform: rotate(-90deg); cursor: pointer; }
.${c}-next-comment { margin-right: 10px; transform: rotate(90deg); cursor: pointer; }
.${c}-prev-comment:hover svg path { fill: #2A66FF; }
.${c}-next-comment:hover svg path { fill: #2A66FF; }
.${c}-all-comment { height: 22px; line-height: 19px; padding-right: 10px; border-right: 1px solid #ddd; color: gray; }
.${c}-current-comment { height: 22px; line-height: 19px; margin-left: 20px; color: gray; }
`;

		return createStyleElement(styleText);
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	//     let css = "";
	//     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	//     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	//     const topCounters = [];

	//     for(let num of numberingPart.numberings) {
	//         const absNum = numberingMap[num.abstractId];

	//         for(let lvl of absNum.levels) {
	//             const className = this.numberingClass(num.id, lvl.level);
	//             let listStyleType = "none";

	//             if(lvl.text && lvl.format == 'decimal') {
	//                 const counter = this.numberingCounter(num.id, lvl.level);

	//                 if (lvl.level > 0) {
	//                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	//                         "counter-reset": counter
	//                     });
	//                 } else {
	//                     topCounters.push(counter);
	//                 }

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": this.levelTextToContent(lvl.text, num.id),
	//                     "counter-increment": counter
	//                 });
	//             } else if(lvl.bulletPictureId) {
	//                 let pict = bulletMap[lvl.bulletPictureId];
	//                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": "' '",
	//                     "display": "inline-block",
	//                     "background": `var(${variable})`
	//                 }, pict.style);

	//                 this.document.loadNumberingImage(pict.referenceId).then(data => {
	//                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	//                     container.appendChild(createStyleElement(text));
	//                 });
	//             } else {
	//                 listStyleType = this.numFormatToCssValue(lvl.format);
	//             }

	//             css += this.styleToString(`p.${className}`, {
	//                 "display": "list-item",
	//                 "list-style-position": "inside",
	//                 "list-style-type": listStyleType,
	//                 //TODO
	//                 //...num.style
	//             });
	//         }
	//     }

	//     if (topCounters.length > 0) {
	//         css += this.styleToString(`.${this.className}-wrapper`, {
	//             "counter-reset": topCounters.join(" ")
	//         });
	//     }

	//     return createStyleElement(css);
	// }

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		var styleText = "";
		var resetCounters = [];

		// console.log(numberings)
		for (var num of numberings) {
			// let abstractId = numbering.find(item => {
			// 	return item.id === num.id;
			// })?.abstractId;

			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.document.loadNumberingImage(num.bullet.src).then(data => {
					var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(createStyleElement(text));
				});
			}
			else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-reset": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return createStyleElement(styleText);
	}

	renderTitleNumbering(pArr, domNumberings, styleContainer) {
		let styleText = "";
		let resetCounters = [];
		for (let p of pArr) {
			if (!isNaN(Number(p.styleName)) && Number(p.styleName) !== 0 && p?.numbering?.id !== '0') {
				let titleLevel = p.styleName;
				const style = this.findStyle(p.styleName);
				let numbering = p.numbering;
				if (numbering && numbering.level !== undefined && numbering.level !== null && (numbering.level > (Number(titleLevel) - 1))) {
					titleLevel = numbering.level;
				} else {
					titleLevel = Number(titleLevel - 1);
				}

				numbering = numbering ?? style?.paragraphProps?.numbering;
				if (numbering === undefined || numbering === null) {
					continue;
				}

				let currentDomNumbering = domNumberings.find(item => {
					return item.numId === numbering.id && item.level === numbering.level;
				});

			  let selector = `p.${this.numberingTitleClass(titleLevel)}`;
			  let listStyleType = "none";

			  if (currentDomNumbering.bullet) {
					let valiable = `--${this.className}-${currentDomNumbering.bullet.src}`.toLowerCase();

					styleText += this.styleToString(`${selector}:before`, {
						"content": "' '",
						"display": "inline-block",
						"background": `var(${valiable})`
					}, currentDomNumbering.bullet.style);

					this.document.loadNumberingImage(currentDomNumbering.bullet.src).then(data => {
						var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
						styleContainer.appendChild(createStyleElement(text));
					});
				}
				else if (currentDomNumbering.levelText) {

					let counter = this.numberingTitleCounter(titleLevel);
					const counterReset = counter + " " + 0;
					if (Number(titleLevel) > 0) {
						styleText += this.styleToString(`p.${this.numberingTitleClass(Number(titleLevel) - 1)}`, {
							"counter-reset": counterReset
						});
					}
					// reset all level counters with start value
					resetCounters.push(counterReset);

					styleText += this.styleToString(`${selector}:before`, {
						"content": this.levelTextToContentOfTitle(currentDomNumbering.levelText, currentDomNumbering.suff, this.numFormatToCssValue(currentDomNumbering.format)),
						"counter-increment": counter,
						...currentDomNumbering.rStyle,
					});
				}
				else {
					listStyleType = this.numFormatToCssValue(currentDomNumbering.format);
				}

				styleText += this.styleToString(selector, {
					"display": "list-item",
					"list-style-position": "inside",
					"list-style-type": listStyleType,
					...currentDomNumbering.pStyle
				});
				
			}
		}

		for (var num of domNumberings) {
			// let abstractId = numbering.find(item => {
			// 	return item.id === num.id;
			// })?.abstractId;

			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.document.loadNumberingImage(num.bullet.src).then(data => {
					var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(createStyleElement(text));
				});
			}
			else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-reset": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return createStyleElement(styleText);
	}

	renderMyNumbering(numberings: IDomNumbering[], overNumberings, styleContainer: HTMLElement) {
		var styleText = "";
		var resetCounters = [];

		// console.log(numberings)
		for (var num of numberings) {
			// let abstractId = numbering.find(item => {
			// 	return item.id === num.id;
			// })?.abstractId;

			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.document.loadNumberingImage(num.bullet.src).then(data => {
					var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(createStyleElement(text));
				});
			}
			else if (num.levelText) {

				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-reset": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return createStyleElement(styleText);
	}

	renderStyles(styles: IDomStyle[]): HTMLElement {
		var styleText = "";
		const stylesMap = this.styleMap;
		const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

		for (const style of styles) {
			var subStyles = style.styles;

			if (style.linked) {
				var linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle)
					subStyles = subStyles.concat(linkedStyle.styles);
				else if (this.options.debug)
					console.warn(`Can't find linked style ${style.linked}`);
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				var selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''} 

				if (style.target != subStyle.target)
					selector += ` ${subStyle.target}`;

				if (defautStyles[style.target] == style)
					selector = `.${this.className} ${style.target}, ` + selector;

				styleText += this.styleToString(selector, subStyle.values);
			}
		}

		return createStyleElement(styleText);
	}

	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, into: HTMLElement) {
		var notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			var result = this.createElement("ol", null, this.renderElements(notes));
			into.appendChild(result);
		}
	}

	renderElement(elem: OpenXmlElement): Node | Node[] {
		switch (elem.type) {
			case DomType.Paragraph:
				return this.renderParagraph(elem as WmlParagraph);

			case DomType.BookmarkStart:
				return this.renderBookmarkStart(elem as WmlBookmarkStart);

			case DomType.BookmarkEnd:
				return null; //ignore bookmark end

			case DomType.Run:
				return this.renderRun(elem as WmlRun);

			case DomType.Table:
				return this.renderTable(elem);

			case DomType.Row:
				return this.renderTableRow(elem);

			case DomType.Cell:
				return this.renderTableCell(elem);

			case DomType.Hyperlink:
				return this.renderHyperlink(elem);

			case DomType.Drawing:
				return this.renderDrawing(elem);

			case DomType.Image:
				return this.renderImage(elem as IDomImage);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.DeletedText:
				return this.renderDeletedText(elem as WmlText);
	
			case DomType.Tab:
				return this.renderTab(elem);

			case DomType.Symbol:
				return this.renderSymbol(elem as WmlSymbol);

			case DomType.Break:
				return this.renderBreak(elem as WmlBreak);

			case DomType.Footer:
				return this.renderContainer(elem, "footer");

			case DomType.Header:
				return this.renderContainer(elem, "header");

			case DomType.Footnote:
			case DomType.Endnote:
				let li = this.renderContainer(elem, "li");
				li.style.fontSize = "10pt";
				return li;

			case DomType.FootnoteReference:
				return this.renderFootnoteReference(elem as WmlNoteReference);

			case DomType.EndnoteReference:
				return this.renderEndnoteReference(elem as WmlNoteReference);

			case DomType.NoBreakHyphen:
				return this.createElement("wbr");

			case DomType.VmlPicture:
				return this.renderVmlPicture(elem);

			case DomType.VmlElement:
				return this.renderVmlElement(elem as VmlElement);
	
			case DomType.MmlMath:
				return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
	
			case DomType.MmlMathParagraph:
				return this.renderContainer(elem, "span");

			case DomType.MmlFraction:
				return this.renderContainerNS(elem, ns.mathML, "mfrac");

			case DomType.MmlBase:
				return this.renderContainerNS(elem, ns.mathML, 
					elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				return this.renderContainerNS(elem, ns.mathML, "mrow");

			case DomType.MmlGroupChar:
				return this.renderMmlGroupChar(elem);

			case DomType.MmlLimitLower:
				return this.renderContainerNS(elem, ns.mathML, "munder");

			case DomType.MmlMatrix:
				return this.renderContainerNS(elem, ns.mathML, "mtable");

			case DomType.MmlMatrixRow:
				return this.renderContainerNS(elem, ns.mathML, "mtr");
	
			case DomType.MmlRadical:
				return this.renderMmlRadical(elem);

			case DomType.MmlSuperscript:
				return this.renderContainerNS(elem, ns.mathML, "msup");

			case DomType.MmlSubscript:
				return this.renderContainerNS(elem, ns.mathML, "msub");

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				return this.renderContainerNS(elem, ns.mathML, "mn");

			case DomType.MmlFunctionName:
				return this.renderContainerNS(elem, ns.mathML, "ms");
	
			case DomType.MmlDelimiter:
				return this.renderMmlDelimiter(elem);

			case DomType.MmlRun:
				return this.renderMmlRun(elem);

			case DomType.MmlNary:
				return this.renderMmlNary(elem);

			case DomType.MmlPreSubSuper:
				return this.renderMmlPreSubSuper(elem);

			case DomType.MmlBar:
				return this.renderMmlBar(elem);
	
			case DomType.MmlEquationArray:
				return this.renderMllList(elem);

			case DomType.Inserted:
				return this.renderInserted(elem);

			case DomType.Deleted:
				return this.renderDeleted(elem);

			case DomType.CommentRangeStart:
				return this.renderCommentRangeStart(elem);

			case DomType.CommentRangeEnd:
				return this.renderCommentRangeEnd(elem);
		}

		return null;
	}

	renderFieldRun(elem: WmlRun): Node {
		for (let child of elem.children ?? []) {
			switch(child.type) {
				case DomType.Instruction:
					return this.renderInstrText(child);
			}
			return null;
		}
		return null;
	}

	renderInstrText(elem): Node {
		let text = elem.text;
		if (text.split(' ')[0] === "REF") {
			let id = Math.random();
			elem.instrTextId = id;
			let result = this.createElement("span");
			let paragraph = elem.parent?.parent;

			let textArr = [];
			let rArr = paragraph?.children || [];
			let index = rArr?.findIndex(item => {
				let rChildren = item.children;
				let c = rChildren?.find(c => {
					return c?.instrTextId === id;
				})
				if (c !== undefined) return true;
				return false;
			})
			for (let i = index; i < rArr?.length || 0; i++) {
				let rChildren = rArr?.[i]?.children;
				let child = rChildren?.find(c => {
					return c?.type === 'complexField' && c?.charType === 'end';
				})
				if (child === undefined) {
					if (rArr?.[i].fieldRun !== true) {
						textArr.push(rArr?.[i]);
						if (rArr?.[i] !== undefined) {
							rArr[i].fieldRun = true;
						}
					}
				} else {
					break;
				}
			}
			result.textContent = this.joinRangeText(textArr);
			result.style.cursor = "pointer";
			let that = this;
			result.addEventListener("click" , function() {
				let ref = text.split(' ')[1];
				let wrap = document.getElementsByClassName(`${that.className}-wrapper`)[0];
				let anchor = wrap.querySelector(`#${ref}`);
				anchor?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
			});
			return result;
		}
		return null;
	}

	renderCommentRangeStart(elem): Node | null {
		let comments = this.document.commentsPart.comments;
		let commentsEx = this.document.commentsExtendedPart.commentsEx;
		let paraId = comments.find(item => {return item.id === elem.id})?.paraId;
		let paraIdParent = commentsEx.find(item => {return item.paraId === paraId})?.paraIdParent;
		if (paraIdParent) {
			return null;
		}
		return this.createCommentSupNode('start', paraId);
	}

	renderCommentRangeEnd(elem): Node | null {
		let comments = this.document.commentsPart.comments;
		let commentsEx = this.document.commentsExtendedPart.commentsEx;
		let paraId = comments.find(item => {return item.id === elem.id})?.paraId;
		let paraIdParent = commentsEx.find(item => {return item.paraId === paraId})?.paraIdParent;
		if (paraIdParent) {
			return null;
		}
		let endElem = this.createCommentSupNode('end', paraId);
		endElem.className = `${this.className}-comment-range-end`;

		return endElem;
	}

	// 创建批注图标
	createCommentSupNode(type, paraId): HTMLElement {
		let elem = this.createElement("span");
		elem.id = `${this.className}-comment-${type}-${paraId}`;
		this.commentIndex += 1;
		return elem;
	}

	renderChildren(elem: OpenXmlElement, into?: Element): Node[] {
		return this.renderElements(elem.children, into);
	}

	renderElements(elems: OpenXmlElement[], into?: Element): Node[] {
		if (elems == null)
			return null;

		var result = elems.flatMap(e => this.renderElement(e)).filter(e => e != null);

		if (into)
			appendChildren(into, result);

		return result;
	}

	renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
		return this.createElement(tagName, props, this.renderChildren(elem));
	}

	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		return createElementNS(ns, tagName, props, this.renderChildren(elem));
	}

	renderParagraph(elem: WmlParagraph) {
		var result = this.createElement("p");

		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		this.renderClass(elem, result);
		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);
		this.renderCommonProperties(result.style, elem);

		const numbering = elem.numbering ?? style?.paragraphProps?.numbering;
		// const numbering = style?.paragraphProps?.numbering ?? elem.numbering;
		// console.log(style?.paragraphProps?.numbering, elem.numbering)

		// if (elem.numbering && style?.paragraphProps?.numbering) {
		// 	numbering.id = style?.paragraphProps?.numbering.id;
		// 	numbering.level = elem.numbering.level;
		// }

		if (this.options.renderNumbering && numbering && !isNaN(Number(elem.styleName)) && Number(elem.styleName) !== 0 && Number(elem.styleName) < 4 && elem?.numbering?.id !== '0') {
			let titleLevel = elem.styleName;
			if (elem.numbering && elem.numbering.level !== undefined && elem.numbering.level !== null && (elem.numbering.level > (Number(titleLevel) - 1))) {
				titleLevel = `${elem.numbering.level}`;
			} else {
				titleLevel = `${(Number(titleLevel) - 1)}`;
			}
			result.classList.add(this.numberingTitleClass(Number(titleLevel)));

			return result;
		}

		if (this.options.renderNumbering && numbering && !elem.noRenderNumbering) {
			let numberingPart = this.document.numberingPart.numberings;
			let currentNumbering = numberingPart.find(item => {
				return item.id === numbering.id;
			});
			result.classList.add(this.numberingClass(currentNumbering?.abstractId, numbering.level));

			let currentOverride = currentNumbering?.overrides?.find(item => {return item.level === numbering.level});
			if (currentOverride) {
				let overrideCounter = this.numberingCounter(currentNumbering?.abstractId, numbering.level);
				let overrideCounterReset = overrideCounter + " " + (currentOverride.start ? currentOverride.start - 1 : 0);
				result.style.counterReset = overrideCounterReset;
			}
		}

		return result;
	}

	renderRunProperties(style: any, props: RunProperties) {
		this.renderCommonProperties(style, props);
	}

	renderCommonProperties(style: any, props: CommonProperties) {
		if (props == null)
			return;

		if (props.color) {
			style["color"] = props.color;
		}

		if (props.fontSize) {
			style["font-size"] = props.fontSize;
		}
	}

	renderHyperlink(elem: WmlHyperlink) {
		var result = this.createElement("a");

		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.href) {
			// result.href = elem.href;
			result.style.cursor = "pointer";
			let that = this;
			result.addEventListener("click", function (e) {
				let wrap = document.getElementsByClassName(`${that.className}-wrapper`)[0];
				let anchor = wrap.querySelector(`#${elem?.href?.replace("#", '')}`);
				anchor?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
			})
		} else if(elem.id) {
			const rel = this.document.documentPart.rels
				.find(it => it.id == elem.id && it.targetMode === "External");
			result.href = rel?.target;
		}

		return result;
	}

	renderDrawing(elem: OpenXmlElement) {
		var result = this.createElement("div");

		result.style.display = "inline-block";
		result.style.position = "relative";
		result.style.textIndent = "0px";

		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		return result;
	}

	renderImage(elem: IDomImage) {
		let result = this.createElement("img");

		this.renderStyleValues(elem.cssStyle, result);

		if (this.document) {
			this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
				result.src = x;
			});
		}

		return result;
	}

	renderText(elem: WmlText) {
		return this.htmlDocument.createTextNode(elem.text);
	}

	renderDeletedText(elem: WmlText) {
		return this.options.renderEndnotes ? this.htmlDocument.createTextNode(elem.text) : null;
	}

	renderBreak(elem: WmlBreak) {
		if (elem.break == "textWrapping") {
			return this.createElement("br");
		}

		return null;
	}

	renderInserted(elem: OpenXmlElement): Node | Node[] {
		if (this.options.renderChanges)
			return this.renderContainer(elem, "ins");

		return this.renderChildren(elem);
	}

	renderDeleted(elem: OpenXmlElement): Node {
		if (this.options.renderChanges)
			return this.renderContainer(elem, "del");

		return null;
	}

	renderSymbol(elem: WmlSymbol) {
		var span = this.createElement("span");
		span.style.fontFamily = elem.font;
		span.innerHTML = `&#x${elem.char};`
		return span;
	}

	// renderFootnoteReference(elem: WmlNoteReference) {
	// 	var result = this.createElement("sup");
	// 	this.currentFootnoteIds.push(elem.id);
	// 	result.textContent = `${this.currentFootnoteIds.length}`;
	// 	return result;
	// }

	renderFootnoteReference(elem: WmlNoteReference) {
		var result = this.createElement("sup");
		this.currentFootnoteIds.push(elem.id);
		result.textContent = `${this.currentFootnoteIds.length}`;
		let notes = this.document?.footnotesPart?.notes;
		result.title = this.renderNotePreview(elem.id, notes);
		return result;
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		var result = this.createElement("sup");
		this.currentEndnoteIds.push(elem.id);
		result.textContent = `${this.currentFootnoteIds.length}`;
		let notes = this.document?.endnotesPart?.notes;
		result.title = this.renderNotePreview(elem.id, notes);
		return result;
	}

	renderNotePreview(id: string, notes: WmlBaseNote[]): string {
		let note = notes.find(item => {
			return item.id === id;
		});
		if (note) {
			let rArr = [];
			let pArr = note.children;
			for (let p of pArr) {
				rArr = rArr.concat(p?.children || []);
			}
			let result = this.joinRangeText(rArr);
			return result;
		}
		return '';
	}

	renderTab(elem: OpenXmlElement) {
		var tabSpan = this.createElement("span");

		tabSpan.innerHTML = "&emsp;";//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			var stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
		var result = this.createElement("span");
		result.id = elem.name;
		return result;
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun)
			return this.renderFieldRun(elem);

		const result = this.createElement("span");

		if (elem.id)
			result.id = elem.id;

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.verticalAlign) {
			const wrapper = this.createElement(elem.verticalAlign as any);
			this.renderChildren(elem, wrapper);
			result.appendChild(wrapper);
		}
		else {
			this.renderChildren(elem, result);
		}

		return result;
	}

	renderTable(elem: WmlTable) {
		let result = this.createElement("table");

		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		if (elem.columns)
			result.appendChild(this.renderTableColumns(elem.columns));

		this.renderClass(elem, result);
		this.renderChildren(elem, result);

		let tblpXSpec = elem.cssStyle?.tblpXSpec;
		if (tblpXSpec === 'center') {
			let num = parseFloat(elem.cssStyle?.width || '0');
			let unit = elem.cssStyle?.width.replace(/[0-9]*/g, '').replace(/\./, '');
			let left = ((100 - Number(num)) / 2).toFixed(2);
			elem.cssStyle['margin-left'] = `${left}${unit}`;
		} else if (tblpXSpec) {
			elem.cssStyle.float = tblpXSpec;
		}

		this.renderStyleValues(elem.cssStyle, result);

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();

		return result;
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		let result = this.createElement("colgroup");

		for (let col of columns) {
			let colElem = this.createElement("col");

			if (col.width)
				colElem.style.width = col.width;

			result.appendChild(colElem);
		}

		return result;
	}

	renderTableRow(elem: OpenXmlElement) {
		let result = this.createElement("tr");

		if (this.currentCellPosition) {
			this.currentCellPosition.col = 0;
		}

		this.renderClass(elem, result);
		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (this.currentCellPosition) {
			this.currentCellPosition.row++;
		}

		return result;
	}

	renderTableCell(elem: WmlTableCell) {
		let result = this.createElement("td");

		const key = this.currentCellPosition?.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == "restart") {
				this.currentVerticalMerge[key] = result;
				result.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				result.style.display = "none";
			}
		} else {
			if (this.currentVerticalMerge) {
				this.currentVerticalMerge[key] = null;
			}
		}

		this.renderClass(elem, result);
		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.span)
			result.colSpan = elem.span;

			if (this.currentCellPosition) {
				this.currentCellPosition.col += result.colSpan;
			}

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		var result = createElement("div");
		this.renderChildren(elem, result);
		return result;
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		var container = createSvgElement("svg");

		container.setAttribute("style", elem.cssStyleText);

		const result = this.renderVmlChildElement(elem);

		if (elem.imageHref?.id) {
			this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			const bb = (container.firstElementChild as any).getBBox();

			container.setAttribute("width", `${Math.ceil(bb.x +  bb.width)}`);
			container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement): any {
		const result = createSvgElement(elem.tagName as any);
		Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				result.appendChild(...asArray(this.renderElement(child as any)));
			}
		}

		return result;
	}

	renderMmlRadical(elem: OpenXmlElement): HTMLElement {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return createElementNS(ns.mathML, "msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return createElementNS(ns.mathML, "mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement): HTMLElement {		
		const children = [];

		children.push(createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));

		return createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement): HTMLElement {		
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;

		const charElem = createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
		} else if(supElem) {
			children.push(createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
		} else if(subElem) {
			children.push(createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = createElementNS(ns.mathML, "mo", null);

		children.push(createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(createElementNS(ns.mathML, "mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const result = this.renderContainerNS(elem, ns.mathML, "mrow");

		switch(elem.props.position) {
			case "top": result.style.textDecoration = "overline"; break
			case "bottom": result.style.textDecoration = "underline"; break
		}

		return result;
	}

	renderMmlRun(elem: OpenXmlElement) {
		const result = createElementNS(ns.mathML, "ms");

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);
		this.renderChildren(elem, result);

		return result;
	}

	renderMllList(elem: OpenXmlElement) {
		const result = createElementNS(ns.mathML, "mtable");

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		const childern = this.renderChildren(elem);

		for (let child of this.renderChildren(elem)) {
			result.appendChild(createElementNS(ns.mathML, "mtr", null, [
				createElementNS(ns.mathML, "mtd", null, [child])
			]));
		}

		return result;
	}


	renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
		for (let k in style) {
			if (k.startsWith("$")) {
				ouput.setAttribute(k.slice(1), style[k]);
			} else {
				ouput.style[k] = style[k];
			}
		}
	}

	renderClass(input: OpenXmlElement, ouput: HTMLElement) {
		if (input.className)
			ouput.className = input.className;

		if (input.styleName)
			ouput.classList.add(this.processStyleName(input.styleName));
	}

	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	numberingTitleClass(lvl: number) {
		return `${this.className}-title-${lvl}`;
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$'))
				continue;
			
			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText)
			result += cssText;

		return result + "}\r\n";
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	numberingTitleCounter(lvl: number) {
		return `${this.className}-title-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		var result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
	}

	levelTextToContentOfTitle(text: string, suff: string, numformat: string) {
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		var result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingTitleCounter(lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
	}

	numFormatToCssValue(format: string) {
		var mapping = {
			none: "none",
			bullet: "disc",
			decimal: "decimal",
			lowerLetter: "lower-alpha",
			upperLetter: "upper-alpha",
			lowerRoman: "lower-roman",
			upperRoman: "upper-roman",
			decimalZero: "decimal-leading-zero", // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: "katakana",
			aiueoFullWidth: "katakana",
			chineseCounting: "simp-chinese-informal",
			chineseCountingThousand: "simp-chinese-informal",
			chineseLegalSimplified: "simp-chinese-formal", // 中文大写
			chosung: "hangul-consonant",
			ideographDigital: "cjk-ideographic",
			ideographTraditional: "cjk-heavenly-stem", // 十天干
			ideographLegalTraditional: "trad-chinese-formal",
			ideographZodiac: "cjk-earthly-branch", // 十二地支
			iroha: "katakana-iroha",
			irohaFullWidth: "katakana-iroha",
			japaneseCounting: "japanese-informal",
			japaneseDigitalTenThousand: "cjk-decimal",
			japaneseLegal: "japanese-formal",
			thaiNumbers: "thai",
			koreanCounting: "korean-hangul-formal",
			koreanDigital: "korean-hangul-formal",
			koreanDigital2: "korean-hanja-informal",
			hebrew1: "hebrew",
			hebrew2: "hebrew",
			hindiNumbers: "devanagari",
			ganada: "hangul",
			taiwaneseCounting: "cjk-ideographic",
			taiwaneseCountingThousand: "cjk-ideographic",
			taiwaneseDigital:  "cjk-decimal",
		};

		return mapping[format] ?? format;
	}

	refreshTabStops() {
		if (!this.options.experimental)
			return;

		clearTimeout(this.tabsTimeout);

		this.tabsTimeout = setTimeout(() => {
			const pixelToPoint = computePixelToPoint();

			for (let tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}

	createElement = createElement;
}

type ChildType = Node | string;

function createElement<T extends keyof HTMLElementTagNameMap>(
	tagName: T,
	props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>,
	children?: ChildType[]
): HTMLElementTagNameMap[T] {
	return createElementNS(undefined, tagName, props, children);
}

function createSvgElement<T extends keyof SVGElementTagNameMap>(
	tagName: T,
	props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>,
	children?: ChildType[]
): SVGElementTagNameMap[T] {
	return createElementNS(ns.svg, tagName, props, children);
}

function createElementNS(ns: string, tagName: string, props?: Partial<Record<any, any>>, children?: ChildType[]): any {
	var result = ns ? document.createElementNS(ns, tagName) : document.createElement(tagName);
	Object.assign(result, props);
	children && appendChildren(result, children);
	return result;
}

function removeAllElements(elem: HTMLElement) {
	elem.innerHTML = '';
}

function appendChildren(elem: Element, children: (Node | string)[]) {
	children.forEach(c => elem.appendChild(isString(c) ? document.createTextNode(c) : c));
}

function createStyleElement(cssText: string) {
	return createElement("style", { innerHTML: cssText });
}

function createScriptElement(jsText: string) {
	return createElement("script", { innerHTML: jsText});
}

function appendComment(elem: HTMLElement, comment: string) {
	elem.appendChild(document.createComment(comment));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	var parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}
