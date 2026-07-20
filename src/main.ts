import { Plugin, FileView, WorkspaceLeaf, TFile, Notice } from "obsidian";
import { renderAsync } from "docx-preview";
import * as XLSX from "xlsx";
import { PptxViewer, RECOMMENDED_ZIP_LIMITS } from "@aiden0z/pptx-renderer";
import hljs from "highlight.js/lib/core";
import langBash from "highlight.js/lib/languages/bash";
import langC from "highlight.js/lib/languages/c";
import langCoffeescript from "highlight.js/lib/languages/coffeescript";
import langCpp from "highlight.js/lib/languages/cpp";
import langCsharp from "highlight.js/lib/languages/csharp";
import langCss from "highlight.js/lib/languages/css";
import langDart from "highlight.js/lib/languages/dart";
import langDockerfile from "highlight.js/lib/languages/dockerfile";
import langDos from "highlight.js/lib/languages/dos";
import langElixir from "highlight.js/lib/languages/elixir";
import langErlang from "highlight.js/lib/languages/erlang";
import langGo from "highlight.js/lib/languages/go";
import langGradle from "highlight.js/lib/languages/gradle";
import langGroovy from "highlight.js/lib/languages/groovy";
import langHtml from "highlight.js/lib/languages/xml";
import langIni from "highlight.js/lib/languages/ini";
import langJava from "highlight.js/lib/languages/java";
import langJavascript from "highlight.js/lib/languages/javascript";
import langJson from "highlight.js/lib/languages/json";
import langKotlin from "highlight.js/lib/languages/kotlin";
import langLess from "highlight.js/lib/languages/less";
import langLua from "highlight.js/lib/languages/lua";
import langMakefile from "highlight.js/lib/languages/makefile";
import langObjectivec from "highlight.js/lib/languages/objectivec";
import langPerl from "highlight.js/lib/languages/perl";
import langPhp from "highlight.js/lib/languages/php";
import langPowershell from "highlight.js/lib/languages/powershell";
import langProperties from "highlight.js/lib/languages/properties";
import langPython from "highlight.js/lib/languages/python";
import langR from "highlight.js/lib/languages/r";
import langRuby from "highlight.js/lib/languages/ruby";
import langRust from "highlight.js/lib/languages/rust";
import langScala from "highlight.js/lib/languages/scala";
import langScss from "highlight.js/lib/languages/scss";
import langSql from "highlight.js/lib/languages/sql";
import langStylus from "highlight.js/lib/languages/stylus";
import langSwift from "highlight.js/lib/languages/swift";
import langLatex from "highlight.js/lib/languages/latex";
import langTypescript from "highlight.js/lib/languages/typescript";
import langYaml from "highlight.js/lib/languages/yaml";

hljs.registerLanguage("bash", langBash);
hljs.registerLanguage("c", langC);
hljs.registerLanguage("coffeescript", langCoffeescript);
hljs.registerLanguage("cpp", langCpp);
hljs.registerLanguage("csharp", langCsharp);
hljs.registerLanguage("css", langCss);
hljs.registerLanguage("dart", langDart);
hljs.registerLanguage("dockerfile", langDockerfile);
hljs.registerLanguage("dos", langDos);
hljs.registerLanguage("elixir", langElixir);
hljs.registerLanguage("erlang", langErlang);
hljs.registerLanguage("go", langGo);
hljs.registerLanguage("gradle", langGradle);
hljs.registerLanguage("groovy", langGroovy);
hljs.registerLanguage("html", langHtml);
hljs.registerLanguage("ini", langIni);
hljs.registerLanguage("java", langJava);
hljs.registerLanguage("javascript", langJavascript);
hljs.registerLanguage("json", langJson);
hljs.registerLanguage("kotlin", langKotlin);
hljs.registerLanguage("less", langLess);
hljs.registerLanguage("lua", langLua);
hljs.registerLanguage("makefile", langMakefile);
hljs.registerLanguage("objectivec", langObjectivec);
hljs.registerLanguage("perl", langPerl);
hljs.registerLanguage("php", langPhp);
hljs.registerLanguage("powershell", langPowershell);
hljs.registerLanguage("properties", langProperties);
hljs.registerLanguage("python", langPython);
hljs.registerLanguage("r", langR);
hljs.registerLanguage("ruby", langRuby);
hljs.registerLanguage("rust", langRust);
hljs.registerLanguage("scala", langScala);
hljs.registerLanguage("scss", langScss);
hljs.registerLanguage("sql", langSql);
hljs.registerLanguage("stylus", langStylus);
hljs.registerLanguage("swift", langSwift);
hljs.registerLanguage("tex", langLatex);
hljs.registerLanguage("toml", langIni);
hljs.registerLanguage("typescript", langTypescript);
hljs.registerLanguage("yaml", langYaml);

const VIEW_TYPE = "file-preview-dev";
const CODE_EXTENSIONS = [
	"txt", "sql", "java", "py", "js", "ts", "jsx", "tsx", "json", "xml", "yaml", "yml",
	"properties", "cfg", "ini", "sh", "bat", "cmd", "ps1", "css", "html", "htm",
	"rst", "tex", "php", "rb", "go", "rs", "c", "cpp", "h", "hpp", "cs",
	"swift", "kt", "scala", "groovy", "pl", "pm", "lua", "r", "m", "mm",
	"gradle", "toml", "conf", "env", "makefile", "dockerfile", "gitignore",
	"vue", "sass", "scss", "less", "styl", "coffee", "dart", "erl", "ex", "exs",
];
const BINARY_EXTENSIONS = ["docx", "xlsx", "xls", "pptx"];
const ALL_EXTENSIONS = [...CODE_EXTENSIONS, ...BINARY_EXTENSIONS];

// highlight.js 语言映射表
const HLJS_LANG: Record<string, string | undefined> = {
	sql: "sql", java: "java", py: "python",
	js: "javascript", ts: "typescript", jsx: "jsx", tsx: "tsx",
	json: "json", xml: "xml", yaml: "yaml", yml: "yaml",
	properties: "properties", cfg: "ini", ini: "ini",
	sh: "bash", bat: "dos", cmd: "dos", ps1: "powershell",
	css: "css", html: "html", htm: "html",
	rst: undefined, tex: "tex",
	php: "php", rb: "ruby", go: "go", rs: "rust",
	c: "c", cpp: "cpp", h: undefined, hpp: "cpp", cs: "csharp",
	swift: "swift", kt: "kotlin", scala: "scala", groovy: "groovy",
	pl: "perl", pm: "perl", lua: "lua", r: "r",
	m: "objectivec", mm: "objectivec",
	gradle: "gradle", toml: "toml", conf: "ini",
	env: undefined, makefile: "makefile", dockerfile: "dockerfile",
	gitignore: undefined, vue: "html",
	sass: "scss", scss: "scss", less: "less", styl: "stylus",
	coffee: "coffeescript", dart: "dart",
	erl: "erlang", ex: "elixir", exs: "elixir",
};

const HIGHLIGHT_MAX_SIZE = 100 * 1024; // >100KB 跳过语法高亮
// highlightAuto 只在已映射的语言中检测，避免扫描全部 190+ 种
const HLJS_LANG_VALUES: string[] = Object.values(HLJS_LANG).filter((v): v is string => v !== undefined);

function getHljsLang(ext: string): string | undefined {
	const key = ext.toLowerCase();
	return key in HLJS_LANG ? HLJS_LANG[key] : undefined;
}

/** 将 hljs 高亮后的 HTML 按行拆开，正确处理跨行 span */
function splitHighlightedLines(html: string): string[] {
	const doc = new DOMParser().parseFromString(html, "text/html");
	const container = doc.body;
	const lines: string[] = [];
	let currentLine = "";
	// 栈里存 { open: "<span class=...>", close: "</span>" }
	const tagStack: Array<{ open: string; close: string }> = [];

	function walk(node: Node) {
		if (node.nodeType === Node.TEXT_NODE) {
			const text = node.textContent || "";
			const parts = text.split("\n");
			for (let i = 0; i < parts.length; i++) {
				if (i > 0) {
					// 关闭当前行所有标签
					const closing = tagStack.slice().reverse().map(t => t.close).join("");
					lines.push(currentLine + closing);
					// 新行重开所有标签
					currentLine = tagStack.map(t => t.open).join("");
				}
				// hljs 输出的文本已经是 HTML 安全的，但 textContent 解码了，需要重新转义
				const escaped = parts[i].replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
				currentLine += escaped;
			}
		} else if (node.nodeType === Node.ELEMENT_NODE) {
			const el = node as HTMLElement;
			const tag = el.tagName.toLowerCase();
			let attrs = "";
			for (let i = 0; i < el.attributes.length; i++) {
				const a = el.attributes[i];
				attrs += ` ${a.name}="${a.value.replace(/"/g, "&quot;")}"`;
			}
			const open = `<${tag}${attrs}>`;
			const close = `</${tag}>`;
			tagStack.push({ open, close });
			currentLine += open;
			for (let i = 0; i < el.childNodes.length; i++) {
				walk(el.childNodes[i]);
			}
			currentLine += close;
			tagStack.pop();
		}
	}

	for (let i = 0; i < container.childNodes.length; i++) {
		walk(container.childNodes[i]);
	}
	if (currentLine) lines.push(currentLine);
	// 保留末尾空行
	if (html.endsWith("\n")) lines.push("");
	return lines;
}

export default class FilePreviewPlugin extends Plugin {
	async onload() {
		this.registerView(VIEW_TYPE, (leaf) => new FilePreviewView(leaf));
		// 逐个注册扩展名——某个被占用不影响其他的
		for (const ext of ALL_EXTENSIONS) {
			try {
				this.registerExtensions([ext], VIEW_TYPE);
			} catch (e) {
				// 扩展名已被其他插件注册，跳过即可
			}
		}
		// 添加侧边栏图标
		this.addRibbonIcon("eye", "Vault Lens Dev: 打开当前文件", () => {
			this.openCurrentFileInDevView();
		});
		// 添加命令，手动用开发版打开当前文件
		this.addCommand({
			id: "open-with-vault-lens-dev",
			name: "用 Vault Lens Dev 打开当前文件",
			callback: () => this.openCurrentFileInDevView(),
		});
	}

	private async openCurrentFileInDevView() {
		const file = this.app.workspace.getActiveFile();
		if (!file) {
			new Notice("没有打开的文件");
			return;
		}
		const ext = file.extension.toLowerCase();
		if (!ALL_EXTENSIONS.includes(ext)) {
			new Notice(`Vault Lens Dev 不支持 .${ext} 文件`);
			return;
		}
		// 用 setViewState 强制以开发版的视图类型打开
		const leaf = this.app.workspace.getLeaf();
		await leaf.setViewState({
			type: VIEW_TYPE,
			state: { file: file.path },
		});
	}
}

class FilePreviewView extends FileView {
	private zoomLevel = 1;
	private zoomIndicatorEl!: HTMLElement;
	private contentArea!: HTMLElement;
	private wrapper!: HTMLElement;
	private pptxViewer: any | null = null;
	private editing = false;
	private isCodeFile = false;
	private textareaEl: HTMLTextAreaElement | null = null;

	constructor(leaf: WorkspaceLeaf) { super(leaf); }
	getViewType(): string { return VIEW_TYPE; }
	getDisplayText(): string { return this.file?.basename ?? "Vault Lens"; }

	async onLoadFile(file: TFile) {
		this.isCodeFile = CODE_EXTENSIONS.includes(file.extension.toLowerCase());
		this.editing = false;
		this.contentEl.empty();
		this.contentEl.addClass("file-preview-container");
		const toolbar = this.contentEl.createDiv({ cls: "file-preview-toolbar" });
		this.buildToolbar(toolbar);
		this.wrapper = this.contentEl.createDiv({ cls: "file-preview-zoom-wrapper" });
		this.contentArea = this.wrapper.createDiv({ cls: "file-preview-content" });
		this.zoomIndicatorEl = this.contentEl.createDiv({ cls: "file-preview-zoom-indicator" });
		this.zoomIndicatorEl.setText("100%");
		this.registerZoomEvents();
		if (this.pptxViewer) { this.pptxViewer.destroy(); this.pptxViewer = null; }
		try {
			const ext = file.extension.toLowerCase();
			if (ext === "pptx") {
				await this.renderPptx(file);
			} else if (ext === "docx") {
				await this.renderDocxRender(file);
			} else {
				this.setHtml(this.contentArea, await this.renderFile(file));
				if (["xlsx", "xls"].includes(ext)) {
					this.setupXlsxTabs();
					this.wrapper.addClass("xlsx-preview-active");
				}
			}
			this.setZoom(this.contentArea, 1);
			this.recordContentSize();
		} catch (e) {
			this.setHtml(this.contentArea, `<div class="file-preview-error">${e instanceof Error ? e.message : String(e)}</div>`);
		}
	}

	async onUnloadFile() {
		if (this.pptxViewer) { this.pptxViewer.destroy(); this.pptxViewer = null; }
		this.wrapper.removeClass("docx-preview-active");
		this.wrapper.removeClass("xlsx-preview-active");
		this.contentEl.empty();
	}

	private setHtml(parent: HTMLElement, html: string) {
		parent.empty();
		if (!html) return;
		const parser = new DOMParser();
		const doc = parser.parseFromString(html, "text/html");
		const nodes = doc.body.childNodes;
		for (let i = 0; i < nodes.length; i++) {
			parent.appendChild(nodes[i].cloneNode(true));
		}
	}

	private async renderFile(file: TFile): Promise<string> {
		const ext = file.extension.toLowerCase();
		if (["xlsx", "xls"].includes(ext)) return await this.renderXlsx(file);
		const content = await this.app.vault.read(file);
		return this.renderCodeWithLineNumbers(content, ext);
	}

	private async renderXlsx(file: TFile): Promise<string> {
		const buf = await this.app.vault.readBinary(file);
		const wb = XLSX.read(buf, { type: "array" });
		const sheets: string[] = [];
		const names = wb.SheetNames;
		for (let si = 0; si < names.length; si++) {
			const ws = wb.Sheets[names[si]];
			const ref = ws["!ref"];
			if (!ref) { sheets.push(`<div class="xlsx-empty">Sheet "${names[si]}" 空</div>`); continue; }
			const rg = XLSX.utils.decode_range(ref);
			const merges = ws["!merges"] || [];
			const cols = ws["!cols"] || [];
			// 构建合并单元格映射
			type MergeInfo = { rs: number; cs: number };
			const mergeMap = new Map<string, MergeInfo>();
			for (const m of merges) {
				for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) {
					const k = r + "," + c;
					if (r === m.s.r && c === m.s.c) mergeMap.set(k, { rs: m.e.r - m.s.r + 1, cs: m.e.c - m.s.c + 1 });
					else mergeMap.set(k, { rs: 0, cs: 0 });
				}
			}
			// 列宽适配
			const colWidths: number[] = [];
			for (let c = rg.s.c; c <= rg.e.c; c++) {
				const col = cols[c];
				let w = 80;
				if (col) {
					if (col.width) w = Math.max(40, Math.round(col.width * 8));
					else if (col.wch) w = Math.max(40, Math.round(col.wch * 8));
				}
				colWidths.push(w);
			}
			// 构建表格
			let html = `<div class="xlsx-page"><table><colgroup>${colWidths.map(w => `<col style="min-width:${w}px;">`).join("")}</colgroup>`;
			for (let r = rg.s.r; r <= rg.e.r; r++) {
				const cells: string[] = [];
				for (let c = rg.s.c; c <= rg.e.c; c++) {
					const mk = r + "," + c;
					const mg = mergeMap.get(mk);
					if (mg && mg.rs === 0) continue;
					const addr = XLSX.utils.encode_cell({ r, c });
					const cell = ws[addr];
					const text = cell ? (cell.w !== undefined ? cell.w : String(cell.v ?? "")) : "";
					let st = "";
					if (cell && cell.s) {
						if (cell.s.font) {
							if (cell.s.font.bold) st += "font-weight:bold;";
							if (cell.s.font.sz) st += `font-size:${Math.round(cell.s.font.sz * 0.75)}pt;`;
							if (cell.s.font.color && cell.s.font.color.rgb) st += `color:#${cell.s.font.color.rgb};`;
						}
						if (cell.s.fill && cell.s.fill.fgColor && cell.s.fill.fgColor.rgb) st += `background-color:#${cell.s.fill.fgColor.rgb};`;
						if (cell.s.alignment) {
							if (cell.s.alignment.horizontal) st += `text-align:${cell.s.alignment.horizontal};`;
							if (cell.s.alignment.vertical) st += `vertical-align:${cell.s.alignment.vertical};`;
						}
					}
					let attrs = "";
					if (mg) { if (mg.cs > 1) attrs += ` colspan="${mg.cs}"`; if (mg.rs > 1) attrs += ` rowspan="${mg.rs}"`; }
					if (st) attrs += ` style="${st}"`;
					cells.push(`<td${attrs}>${this.esc(text) || " "}</td>`);
				}
				if (cells.length > 0) html += `<tr>${cells.join("")}</tr>`;
			}
			html += `</table></div>`;
			sheets.push(`<div class="xlsx-sheet${si === 0 ? " active" : ""}" data-sheet="${si}">${html}</div>`);
		}
		const tabs = names.map((n, i) => `<span class="xlsx-tab${i === 0 ? " active" : ""}" data-sheet="${i}">${this.esc(n)}</span>`).join("");
		return `<div class="file-preview-xlsx"><div class="xlsx-toolbar"><div class="xlsx-tabs">${tabs}</div></div>${sheets.join("")}</div>`;
	}

	private async renderPptx(file: TFile) {
		this.setHtml(this.contentArea, '<div class="file-preview-loading">加载 PPT 中...</div>');
		this.pptxViewer = await PptxViewer.open(await this.app.vault.readBinary(file), this.contentArea, {
			zipLimits: RECOMMENDED_ZIP_LIMITS, listOptions: { windowed: false },
		});
	}

	private async renderDocxRender(file: TFile) {
		this.setHtml(this.contentArea, '<div class="file-preview-loading">加载 DOCX 中...</div>');
		this.wrapper.addClass("docx-preview-active");
		const buf = await this.app.vault.readBinary(file);
		this.contentArea.empty();
		await renderAsync(buf, this.contentArea, undefined, {
			breakPages: true,
			ignoreWidth: false,
			ignoreHeight: false,
			ignoreFonts: false,
			renderHeaders: true,
			renderFooters: true,
			renderFootnotes: true,
			renderEndnotes: true,
			inWrapper: true,
			className: "file-preview-docx",
		});
	}

	private renderCodeWithLineNumbers(content: string, ext?: string): string {
		if (ext && content.length <= HIGHLIGHT_MAX_SIZE) {
			try {
				const lang = getHljsLang(ext);
				let highlighted: string;
				if (lang) {
					highlighted = hljs.highlight(content, { language: lang, ignoreIllegals: true }).value;
				} else {
					highlighted = hljs.highlightAuto(content, HLJS_LANG_VALUES).value;
				}
				const hlLines = splitHighlightedLines(highlighted);
				return `<table class="file-preview-txt">${hlLines.map((line, i) =>
					`<tr><td class="line-num">${i + 1}</td><td class="line-content">${line || " "}</td></tr>`
				).join("\n")}</table>`;
			} catch (e) {
				console.warn("Vault Lens: hljs highlighting failed, falling back to plain text", e);
			}
		}
		const lines = content.split("\n");
		return `<table class="file-preview-txt">${lines.map((line, i) =>
			`<tr><td class="line-num">${i + 1}</td><td class="line-content">${this.esc(line) || " "}</td></tr>`).join("\n")}</table>`;
	}

	private buildToolbar(c: HTMLElement) {
		const add = (t: string, cb: () => void) => { const s = c.createSpan({ cls: "file-preview-toolbar-btn" }); s.textContent = t; s.addEventListener("click", cb); };
		add("＋", () => this.adjustZoom(0.25)); add("－", () => this.adjustZoom(-0.25)); add("1:1", () => this.resetZoom());
		if (this.isCodeFile) {
			const eb = c.createSpan({ cls: "file-preview-toolbar-btn", text: "✏️ 编辑" });
			eb.addEventListener("click", async () => {
				if (!this.editing) {
					const lines: string[] = [];
					this.contentArea.querySelectorAll(".line-content").forEach(c => lines.push(c.textContent || ""));
					const fullText = lines.join("\n");
					const table = this.contentArea.querySelector("table.file-preview-txt") as HTMLElement | null;
					if (table) table.addClass("file-preview-hide");
					this.contentArea.addClass("file-preview-hide");
					this.wrapper.addClass("file-preview-editing-wrapper");
					const editorWrap = document.createElement("div");
					editorWrap.className = "file-preview-editor-wrap";
					const lineNumDiv = document.createElement("div");
					lineNumDiv.className = "file-preview-line-nums";
					const textarea = document.createElement("textarea");
					textarea.className = "file-preview-textarea";
					textarea.value = fullText;
					textarea.spellcheck = false;
					this.textareaEl = textarea;
					const lineCount = fullText.split("\n").length;
					for (let i = 0; i < lineCount; i++) {
						lineNumDiv.createSpan({ text: String(i + 1) });
					}
					editorWrap.appendChild(lineNumDiv);
					editorWrap.appendChild(textarea);
					this.wrapper.appendChild(editorWrap);
					textarea.addEventListener("scroll", () => { lineNumDiv.scrollTop = textarea.scrollTop; });
					textarea.addEventListener("input", () => {
						const newCount = textarea.value.split("\n").length;
						const currentCount = lineNumDiv.children.length;
						if (newCount !== currentCount) {
							while (lineNumDiv.firstChild) lineNumDiv.firstChild.remove();
							for (let i = 0; i < newCount; i++) { lineNumDiv.createSpan({ text: String(i + 1) }); }
						}
					});
					textarea.addEventListener("keydown", (e: KeyboardEvent) => {
						if (e.key === "Tab") {
							e.preventDefault();
							const start = textarea.selectionStart;
							const end = textarea.selectionEnd;
							textarea.value = textarea.value.substring(0, start) + "\t" + textarea.value.substring(end);
							textarea.selectionStart = textarea.selectionEnd = start + 1;
							textarea.dispatchEvent(new Event("input"));
						}
					});
					this.editing = true; eb.textContent = "💾 保存"; new Notice("编辑模式 — Ctrl+S 保存");
					textarea.focus();
				} else if (this.file) {
					try {
						const text = this.textareaEl ? this.textareaEl.value : "";
						await this.app.vault.modify(this.file, text);
						new Notice("已保存"); this.editing = false; eb.textContent = "✏️ 编辑";
						const wrap = this.wrapper.querySelector(".file-preview-editor-wrap") as HTMLElement | null;
						if (wrap) { wrap.remove(); }
						this.textareaEl = null;
						this.contentArea.removeClass("file-preview-hide");
						this.wrapper.removeClass("file-preview-editing-wrapper");
						this.setHtml(this.contentArea, this.renderCodeWithLineNumbers(text, this.file!.extension));
					} catch (e) {
						new Notice("保存失败: " + (e instanceof Error ? e.message : String(e)));
					}
				}
			});
			this.contentEl.addEventListener("keydown", (e: KeyboardEvent) => {
				if (this.editing && (e.ctrlKey || e.metaKey) && e.key === "s") { e.preventDefault(); eb.click(); }
			});
		}
	}

	private registerZoomEvents() {
		this.wrapper.addEventListener("wheel", (e: WheelEvent) => { if (e.ctrlKey || e.metaKey) { e.preventDefault(); this.setZoom(this.contentArea, this.zoomLevel + (e.deltaY > 0 ? -0.1 : 0.1)); } });
		let d = false, sx = 0, sy = 0, sl = 0, st = 0;
		this.wrapper.addEventListener("mousedown", (e: MouseEvent) => { if (this.zoomLevel > 1) { d = true; sx = e.pageX - this.wrapper.offsetLeft; sy = e.pageY - this.wrapper.offsetTop; sl = this.wrapper.scrollLeft; st = this.wrapper.scrollTop; this.wrapper.addClass("file-preview-grabbing"); e.preventDefault(); } });
		this.wrapper.addEventListener("mousemove", (e: MouseEvent) => { if (!d) return; e.preventDefault(); this.wrapper.scrollLeft = sl - (e.pageX - this.wrapper.offsetLeft - sx) * 2; this.wrapper.scrollTop = st - (e.pageY - this.wrapper.offsetTop - sy) * 2; });
		const stop = () => { d = false; this.wrapper.removeClass("file-preview-grabbing"); };
		this.wrapper.addEventListener("mouseup", stop); this.wrapper.addEventListener("mouseleave", stop);
	}
	private adjustZoom(d: number) { if (this.contentArea) this.setZoom(this.contentArea, this.zoomLevel + d); }
	private resetZoom() { if (this.contentArea) this.setZoom(this.contentArea, 1); }

	private setZoom(el: HTMLElement, lv: number) {
		this.zoomLevel = Math.max(0.25, Math.min(2, Math.round(lv * 100) / 100));
		// XLSX: 用 CSS zoom 而非 transform scale，zoom 影响实际布局
		if (this.wrapper.hasClass("xlsx-preview-active")) {
			const xlsxEl = this.contentArea.querySelector(".file-preview-xlsx") as HTMLElement | null;
			if (xlsxEl) {
				xlsxEl.style.setProperty("zoom", String(this.zoomLevel));
			}
		} else {
			el.setAttribute("style", `transform:scale(${this.zoomLevel});transform-origin:0 0`);
			if (this.zoomLevel === 1) {
				el.style.removeProperty("width");
				el.style.removeProperty("height");
				el.removeAttribute("data-ow");
				el.removeAttribute("data-oh");
			} else {
				if (!el.getAttribute("data-ow")) el.setAttribute("data-ow", String(el.scrollWidth));
				if (!el.getAttribute("data-oh")) el.setAttribute("data-oh", String(el.scrollHeight));
				const _ow = parseInt(el.getAttribute("data-ow")!);
				const _oh = parseInt(el.getAttribute("data-oh")!);
				el.style.setProperty("width", (_ow * this.zoomLevel) + "px");
				el.style.setProperty("height", (_oh * this.zoomLevel) + "px");
			}
		}
		this.zoomIndicatorEl.setText(`${Math.round(this.zoomLevel * 100)}%`);
	}

	private setupXlsxTabs() {
		const tabs = this.contentArea.querySelectorAll(".xlsx-tab");
		tabs.forEach(t => t.addEventListener("click", () => {
			const idx = t.getAttribute("data-sheet");
			tabs.forEach(x => x.removeClass("active")); t.addClass("active");
			this.contentArea.querySelectorAll(".xlsx-sheet").forEach(s => { s.removeClass("active"); if (s.getAttribute("data-sheet") === idx) s.addClass("active"); });
		}));
	}
	private recordContentSize() {
		const el = this.contentArea;
		window.setTimeout(() => {
			const w = el.scrollWidth;
			const h = el.scrollHeight;
			if (w > 0) el.setAttribute("data-ow", String(w));
			if (h > 0) el.setAttribute("data-oh", String(h));
		}, 0);
	}
	private esc(s: string): string { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
}

