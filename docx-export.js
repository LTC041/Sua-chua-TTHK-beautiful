/* =========================================================
   XUẤT WORD (.docx)
   - Dựng bản chụp hồ sơ trong một iframe ẩn, đo vị trí và định dạng thật của
     từng khối (getBoundingClientRect + getComputedStyle) rồi viết ra Office Open
     XML, nên file Word giữ đúng bố cục trang A4: lề, phông, cỡ chữ, căn lề,
     khoảng cách, bảng, cột chữ ký.
   - Khối dàn ngang (đầu trang hai bên, các cột chữ ký, "Ông (bà) ... Chức vụ ...")
     thành bảng không viền – đúng cách soạn văn bản hành chính trong Word.
   - Mỗi trang A4 là một section, lề trang lấy theo padding của .page.
   - Không dùng thư viện ngoài: tự đóng gói ZIP, nén bằng CompressionStream nếu có.
   ========================================================= */

const DocxExport = (() => {
  "use strict";

  const MIME =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  const PX_TO_TWIP = 15; // 1px CSS = 1/96 inch = 15 twip (1/20 point)
  const PX_PER_MM = 96 / 25.4;
  const A4_TWIP = { width: 11906, height: 16838 };
  const FALLBACK_MARGIN_MM = { Top: 20, Right: 20, Bottom: 20, Left: 30 };
  const DEFAULT_FONT = "Times New Roman";
  const NORMAL_LINE_HEIGHT = 1.149; // line-height: normal của Times New Roman = dòng đơn trong Word
  const SAFE_EDGE_PX = 8 * PX_PER_MM; // cách mép giấy tối thiểu 8mm (vùng máy in không in tới)
  const SLACK_PX = 24; // nới ô chứa chữ để Word không xuống dòng vì đo chữ lệch trình duyệt vài px
  const CENTER_SLACK_PX = 12;
  const MIN_ROOM_PX = 8;
  const EPSILON_PX = 1.5;
  const RENDER_TIMEOUT_MS = 15000;
  const APP_NAME = "Quản lý sửa chữa TTHK";

  const SKIP_TAGS = new Set(["SCRIPT", "STYLE", "TEMPLATE", "NOSCRIPT", "BUTTON", "SELECT"]);
  const BLOCK_DISPLAYS = new Set(["block", "flex", "grid", "table", "list-item", "flow-root"]);
  const NON_TEXT_INPUTS = new Set(["checkbox", "radio", "hidden", "button", "submit", "reset", "file", "image"]);
  const GENERIC_FONTS = {
    serif: DEFAULT_FONT,
    "sans-serif": "Arial",
    monospace: "Courier New",
    "system-ui": "Segoe UI",
  };
  const BORDER_STYLES = {
    solid: "single",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    groove: "threeDEngrave",
    ridge: "threeDEmboss",
    inset: "inset",
    outset: "outset",
  };
  const VERTICAL_ALIGN = { middle: "center", bottom: "bottom" };
  const SIDES = [
    ["top", "Top"],
    ["left", "Left"],
    ["bottom", "Bottom"],
    ["right", "Right"],
  ];

  const COLLAPSIBLE_SPACE = /[ \t\n\r\f]+/g;
  const VISIBLE_TEXT = /[^ \t\n\r\f]/;
  const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

  // Đoạn trống cao 1pt: ngăn cách hai bảng liền nhau, kết thúc ô bảng...
  const TINY_PARAGRAPH =
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>' +
    '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';

  /* ---------- Tiện ích ---------- */

  const twip = (px) => Math.round(px * PX_TO_TWIP);
  const halfPoints = (px) => Math.max(2, Math.round(px * 1.5)); // 1px = 0.75pt
  const num = (value) => parseFloat(value) || 0;
  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
  const sum = (values) => values.reduce((total, value) => total + value, 0);
  const styleOf = (element) => element.ownerDocument.defaultView.getComputedStyle(element);
  const isText = (node) => node.nodeType === 3;
  const isElement = (node) => node.nodeType === 1;

  function escapeXml(value) {
    return String(value)
      .replace(INVALID_XML, "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // left/right: khung chữ đo trên trang; actual: vùng chữ thật của ô trong Word (rộng hơn khi ô được nới)
  function makeFrame(left, right, { cursor = null, safe = null, actual = null, anchor = null } = {}) {
    return {
      left,
      right,
      anchor, // ô dàn ngang bị dồn sát mép: chữ một dòng neo về mép đó
      cursor, // đáy của khối vừa xuất (null: khối đầu tiên trong ô, không tính khoảng cách phía trên)
      safeLeft: safe ? safe.left : -Infinity,
      safeRight: safe ? safe.right : Infinity,
      actualLeft: actual ? actual.left : left,
      actualRight: actual ? actual.right : right,
    };
  }

  /* ---------- Đọc DOM ---------- */

  function isSkipped(element) {
    if (SKIP_TAGS.has(element.tagName) || element.hasAttribute("data-editor-ui")) return true;
    const style = styleOf(element);
    return style.display === "none" || style.position === "fixed";
  }

  function isBlockLevel(element) {
    return BLOCK_DISPLAYS.has(styleOf(element).display);
  }

  function isBlockified(style) {
    return /(flex|grid)$/.test(style.display);
  }

  function isCenteredColumn(style) {
    return (
      /flex$/.test(style.display) &&
      style.flexDirection.startsWith("column") &&
      style.alignItems.includes("center")
    );
  }

  function hasContent(node) {
    if (isText(node)) return VISIBLE_TEXT.test(node.data);
    if (!isElement(node) || isSkipped(node)) return false;
    if (["BR", "INPUT", "TEXTAREA"].includes(node.tagName)) return true;
    return [...node.childNodes].some(hasContent);
  }

  function hasBlockContent(element, style) {
    const blockified = isBlockified(style);
    return [...element.childNodes].some(
      (node) => isElement(node) && !isSkipped(node) && (blockified || isBlockLevel(node)),
    );
  }

  // Bỏ các nút chỉ có khoảng trắng ở hai đầu (không hiển thị trên trang)
  function trimNodes(nodes) {
    let start = 0;
    let end = nodes.length;
    while (start < end && isText(nodes[start]) && !hasContent(nodes[start])) start += 1;
    while (end > start && isText(nodes[end - 1]) && !hasContent(nodes[end - 1])) end -= 1;
    return nodes.slice(start, end);
  }

  function rangeAround(nodes) {
    const range = nodes[0].ownerDocument.createRange();
    range.setStartBefore(nodes[0]);
    range.setEndAfter(nodes[nodes.length - 1]);
    return range;
  }

  function rectOf(item) {
    if (!item.text) return item.node.getBoundingClientRect();
    const range = item.node.ownerDocument.createRange();
    range.selectNodeContents(item.node);
    return range.getBoundingClientRect();
  }

  function contentBox(element, rect = element.getBoundingClientRect(), style = styleOf(element)) {
    return {
      left: rect.left + num(style.borderLeftWidth) + num(style.paddingLeft),
      right: rect.right - num(style.borderRightWidth) - num(style.paddingRight),
      top: rect.top + num(style.borderTopWidth) + num(style.paddingTop),
      bottom: rect.bottom - num(style.borderBottomWidth) - num(style.paddingBottom),
    };
  }

  function paddingOf(style) {
    return {
      top: num(style.paddingTop),
      left: num(style.paddingLeft),
      bottom: num(style.paddingBottom),
      right: num(style.paddingRight),
    };
  }

  function countLines(range, lineHeight) {
    const middles = [...range.getClientRects()]
      .filter((rect) => rect.width > 0.5 && rect.height > 0)
      .map((rect) => (rect.top + rect.bottom) / 2)
      .sort((a, b) => a - b);

    let lines = 0;
    let lineMiddle = -Infinity;
    for (const middle of middles) {
      if (middle - lineMiddle > lineHeight / 2) {
        lines += 1;
        lineMiddle = middle;
      }
    }
    return lines;
  }

  function alignOf(style) {
    switch (style.textAlign) {
      case "center":
      case "-webkit-center":
        return "center";
      case "right":
      case "-webkit-right":
        return "right";
      case "end":
        return style.direction === "rtl" ? "left" : "right";
      case "justify":
        return "both";
      default:
        return "left";
    }
  }

  function lengthPx(value, basis) {
    return String(value).endsWith("%") ? (num(value) * basis) / 100 : num(value);
  }

  // Word "multiple": 240 = dòng đơn
  function lineSpacing(lineHeight, fontPx) {
    return clamp(Math.round((240 * lineHeight) / (fontPx * NORMAL_LINE_HEIGHT)), 120, 960);
  }

  function parseColor(value) {
    const match = /rgba?\(([^)]+)\)/.exec(value || "");
    if (!match) return null;
    const [r, g, b, a = 1] = match[1].split(/[\s,/]+/).filter(Boolean).map(Number);
    return { r, g, b, a };
  }

  function toHex({ r, g, b }) {
    return [r, g, b]
      .map((value) => Math.round(value).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  }

  function borderOf(style, side) {
    const width = num(style[`border${side}Width`]);
    const kind = style[`border${side}Style`];
    const color = parseColor(style[`border${side}Color`]);
    if (width <= 0 || kind === "none" || kind === "hidden" || !color || color.a === 0) return null;
    return {
      val: BORDER_STYLES[kind] || "single",
      sz: clamp(Math.round(width * 6), 2, 96), // 1/8 point
      color: toHex(color),
    };
  }

  function fontFamilyOf(value) {
    const first = value.split(",")[0].trim().replace(/^["']|["']$/g, "");
    return GENERIC_FONTS[first.toLowerCase()] || first || DEFAULT_FONT;
  }

  /* ---------- Chữ trong một đoạn ---------- */

  function decorationOf(style, inherited) {
    const line = style.textDecorationLine || "";
    return {
      underline: inherited.underline || line.includes("underline"),
      strike: inherited.strike || line.includes("line-through"),
    };
  }

  function transformText(text, style) {
    if (style.textTransform === "uppercase") return text.toLocaleUpperCase("vi");
    if (style.textTransform === "lowercase") return text.toLocaleLowerCase("vi");
    return text;
  }

  function runProps(style, decoration) {
    const color = parseColor(style.color);
    const isBlack = color && Math.max(color.r, color.g, color.b) <= 40;
    return {
      font: fontFamilyOf(style.fontFamily),
      size: halfPoints(num(style.fontSize)),
      bold: (parseInt(style.fontWeight, 10) || 400) >= 600,
      italic: style.fontStyle !== "normal",
      underline: decoration.underline,
      strike: decoration.strike,
      // Chữ trong suốt (tên ẩn khi in) -> chữ ẩn của Word: vẫn giữ dòng, không in ra
      hidden: style.visibility !== "visible" || !color || color.a === 0,
      color: color && !isBlack ? toHex(color) : null,
      spacing: style.letterSpacing === "normal" ? 0 : twip(num(style.letterSpacing)),
      vertAlign:
        style.verticalAlign === "super"
          ? "superscript"
          : style.verticalAlign === "sub"
            ? "subscript"
            : null,
    };
  }

  function collectRuns(nodes, rootStyle) {
    const segments = [];

    const visit = (node, decoration) => {
      if (isText(node)) {
        const style = styleOf(node.parentElement);
        segments.push({
          text: transformText(node.data, style),
          whiteSpace: style.whiteSpace,
          props: runProps(style, decoration),
        });
        return;
      }
      if (!isElement(node) || isSkipped(node)) return;
      if (node.tagName === "BR") {
        segments.push({ br: true });
        return;
      }

      const style = styleOf(node);
      const next = decorationOf(style, decoration);
      if (node.tagName === "TEXTAREA" || (node.tagName === "INPUT" && !NON_TEXT_INPUTS.has(node.type))) {
        segments.push({
          text: transformText(node.value, style),
          whiteSpace: "pre-wrap",
          props: runProps(style, next),
        });
        return;
      }
      node.childNodes.forEach((child) => visit(child, next));
    };

    const rootDecoration = decorationOf(rootStyle, { underline: false, strike: false });
    nodes.forEach((node) => visit(node, rootDecoration));
    return normalizeWhitespace(segments);
  }

  // Gộp khoảng trắng giống trình duyệt (white-space: normal / pre / pre-wrap / pre-line)
  function normalizeWhitespace(segments) {
    const runs = [];
    let lineStart = true;
    let spaceBefore = false;

    const trimLineEnd = () => {
      for (let i = runs.length - 1; i >= 0 && !runs[i].br && runs[i].collapsible; i -= 1) {
        runs[i].text = runs[i].text.replace(/ +$/, "");
        if (runs[i].text) break;
        runs.splice(i, 1);
      }
    };
    const pushBreak = () => {
      trimLineEnd();
      runs.push({ br: true });
      lineStart = true;
      spaceBefore = false;
    };

    for (const segment of segments) {
      if (segment.br) {
        pushBreak();
        continue;
      }

      const preserveSpaces = /^(pre|pre-wrap|break-spaces)$/.test(segment.whiteSpace);
      const keepNewlines = preserveSpaces || segment.whiteSpace === "pre-line";
      const parts = keepNewlines ? segment.text.split(/\r\n|\r|\n/) : [segment.text];
      const key = JSON.stringify(segment.props);

      parts.forEach((part, index) => {
        if (index > 0) pushBreak();

        let text;
        if (preserveSpaces) {
          text = part.replace(/\t/g, "    ");
          if (!text) return;
          spaceBefore = false;
        } else {
          text = part.replace(COLLAPSIBLE_SPACE, " ");
          if ((lineStart || spaceBefore) && text.startsWith(" ")) text = text.slice(1);
          if (!text) return;
          spaceBefore = text.endsWith(" ");
        }

        lineStart = false;
        const last = runs[runs.length - 1];
        if (last && !last.br && last.key === key && last.collapsible === !preserveSpaces) {
          last.text += text;
        } else {
          runs.push({ text, props: segment.props, key, collapsible: !preserveSpaces });
        }
      });
    }

    trimLineEnd();
    // <br> cuối khối không tạo thêm dòng trống
    if (runs[runs.length - 1]?.br) runs.pop();
    return runs;
  }

  /* ---------- Đoạn văn ---------- */

  // source: { styleEl, nodes, boxEl?: phần tử tạo đoạn, hbox?: khung ngang của đoạn vô danh }
  // center: khung của flex-column căn giữa chứa đoạn này (nếu có)
  function emitParagraph({ styleEl, nodes, boxEl = null, hbox = null }, frame, out, center = null) {
    const style = styleOf(styleEl);
    const range = nodes.length ? rangeAround(nodes) : null;
    const box = boxEl ? boxEl.getBoundingClientRect() : range?.getBoundingClientRect();
    if (!box) return;
    const runs = collectRuns(nodes, style);
    if (!runs.length && box.height < 1) return; // khối rỗng không chiếm chỗ trên trang

    const fontPx = num(style.fontSize);
    const lineHeight =
      style.lineHeight === "normal" ? fontPx * NORMAL_LINE_HEIGHT : num(style.lineHeight);
    const lines = Math.max(1, range ? countLines(range, lineHeight) : 0);

    // Khoảng cách dọc đo từ vị trí thật trên trang (đã tính margin gộp, position, chiều cao cố định...)
    const topRule = boxEl ? borderOf(style, "Top") : null;
    const bottomRule = boxEl ? borderOf(style, "Bottom") : null;
    const insetTop = boxEl ? num(style.borderTopWidth) + num(style.paddingTop) : 0;
    const gap = frame.cursor == null ? 0 : Math.max(0, box.top - frame.cursor);
    const extra = boxEl
      ? box.height - insetTop - num(style.borderBottomWidth) - lines * lineHeight
      : 0;
    frame.cursor = box.bottom;

    const inner = boxEl ? contentBox(boxEl, box, style) : hbox || box;
    let jc = alignOf(style);
    let indLeft = inner.left - frame.left;
    let indRight = frame.right - inner.right;

    if (center && (lines === 1 || box.width < center.right - center.left - EPSILON_PX)) {
      // Dòng trong cột flex-column căn giữa -> đoạn căn giữa (sửa chữ trong Word vẫn giữ giữa)
      jc = "center";
      indLeft = center.left - frame.left;
      indRight = frame.right - center.right;
    }
    if (lines === 1 && frame.anchor && jc !== "center") jc = frame.anchor;
    if (lines === 1) {
      // Chữ một dòng: bỏ phần thụt lề không làm đổi vị trí chữ,
      // để Word (đo chữ hơi khác trình duyệt) không tự xuống dòng
      if (jc === "center") {
        const shared = Math.min(indLeft, indRight);
        if (shared > 0) {
          indLeft -= shared;
          indRight -= shared;
        }
      } else if (jc === "right") {
        indLeft = Math.min(indLeft, 0);
      } else {
        indRight = Math.min(indRight, 0);
      }
    }

    // Đường kẻ trên/dưới (dưới tên đơn vị, "Độc lập – Tự do – Hạnh phúc"):
    // dài đúng bằng khối và nằm đúng chỗ chữ sẽ hiện trong Word
    let rule = null;
    if (topRule || bottomRule) {
      const regionLeft = frame.actualLeft + indLeft;
      const regionRight = frame.actualRight - indRight;
      let left = regionLeft - (inner.left - box.left);
      if (jc === "center") left = (regionLeft + regionRight - box.width) / 2;
      else if (jc === "right") left = regionRight + (box.right - inner.right) - box.width;
      rule = { indLeft: left - frame.actualLeft, indRight: frame.actualRight - left - box.width };
    }

    if (topRule) out.push(ruleParagraph(topRule, "bottom", rule, gap));
    out.push({
      type: "p",
      jc,
      indLeft,
      indRight,
      firstLine: lengthPx(style.textIndent, inner.right - inner.left),
      before: (topRule ? 0 : gap) + insetTop,
      after: extra > 1 ? extra : 0,
      line: lineSpacing(lineHeight, fontPx),
      lineRule: "auto",
      runs,
      markSize: halfPoints(fontPx),
    });
    if (bottomRule) out.push(ruleParagraph(bottomRule, "top", rule, 0));
  }

  // Đoạn trống cao 1pt chỉ để vẽ đường kẻ
  function ruleParagraph(border, side, { indLeft, indRight }, before) {
    return {
      type: "p",
      jc: "left",
      indLeft,
      indRight,
      firstLine: 0,
      before,
      after: 0,
      line: 20,
      lineRule: "exact",
      runs: [],
      markSize: 2,
      border: { side, ...border },
    };
  }

  function spacerParagraph(heightPx) {
    return {
      type: "p",
      jc: "left",
      indLeft: 0,
      indRight: 0,
      firstLine: 0,
      before: 0,
      after: 0,
      line: Math.max(20, twip(heightPx)),
      lineRule: "exact",
      runs: [],
      markSize: 2,
    };
  }

  // Khoảng trống phía trên một bảng
  function pushGap(out, frame, top) {
    const gap = frame.cursor == null ? 0 : Math.max(0, top - frame.cursor);
    const last = out[out.length - 1];
    if (last?.type === "tbl") {
      // Hai bảng liền nhau sẽ bị Word gộp làm một -> luôn chèn đoạn ngăn cách
      out.push(spacerParagraph(gap));
    } else if (gap >= 1) {
      if (last) last.after += gap;
      else out.push(spacerParagraph(gap));
    }
  }

  /* ---------- Duyệt khối ---------- */

  function emitBlock(element, frame, out, center = null) {
    const style = styleOf(element);
    if (element.tagName === "TABLE" || /table$/.test(style.display)) {
      emitTable(element, frame, out);
      return;
    }
    if (element.tagName === "HR") {
      emitHorizontalRule(element, frame, out);
      return;
    }

    const items = layoutItems(element, style);
    if (items) emitLayout(element, items, frame, out);
    else if (hasBlockContent(element, style)) emitFlow(element, frame, out);
    else emitParagraph({ styleEl: element, boxEl: element, nodes: [...element.childNodes] }, frame, out, center);
  }

  // Nội dung bên trong một phần tử làm khung (trang, ô bảng, ô dàn ngang)
  function emitContents(element, frame, out) {
    const style = styleOf(element);
    const items = layoutItems(element, style);
    if (items) {
      emitLayout(element, items, frame, out);
      return;
    }
    if (hasBlockContent(element, style)) {
      emitFlow(element, frame, out);
      return;
    }
    const nodes = trimNodes([...element.childNodes]);
    if (nodes.some(hasContent)) {
      emitParagraph({ styleEl: element, nodes, hbox: contentBox(element) }, frame, out);
    }
  }

  function emitFlow(container, frame, out) {
    const style = styleOf(container);
    const blockified = isBlockified(style);
    const center = isCenteredColumn(style) ? contentBox(container) : null;
    const hbox = contentBox(container);
    let inline = [];

    const flush = () => {
      const nodes = trimNodes(inline);
      if (nodes.some(hasContent)) emitParagraph({ styleEl: container, nodes, hbox }, frame, out);
      inline = [];
    };

    for (const node of container.childNodes) {
      if (isText(node)) {
        if (!blockified) inline.push(node);
        else if (hasContent(node)) emitParagraph({ styleEl: container, nodes: [node] }, frame, out, center);
      } else if (isElement(node) && !isSkipped(node)) {
        if (!blockified && !isBlockLevel(node)) {
          inline.push(node);
        } else {
          flush();
          emitBlock(node, frame, out, center);
        }
      }
    }
    flush();
  }

  function emitHorizontalRule(element, frame, out) {
    const rect = element.getBoundingClientRect();
    const style = styleOf(element);
    const border = borderOf(style, "Top") ||
      borderOf(style, "Bottom") || { val: "single", sz: 6, color: "000000" };
    const gap = frame.cursor == null ? 0 : Math.max(0, rect.top - frame.cursor);
    const indents = { indLeft: rect.left - frame.actualLeft, indRight: frame.actualRight - rect.right };
    out.push(ruleParagraph(border, "top", indents, gap));
    frame.cursor = rect.bottom;
  }

  /* ---------- Lưới cột ---------- */

  function mergeBounds(values) {
    const bounds = [];
    for (const value of values.filter(Number.isFinite).sort((a, b) => a - b)) {
      if (!bounds.length || value - bounds[bounds.length - 1] > EPSILON_PX) bounds.push(value);
    }
    return bounds;
  }

  function nearestIndex(bounds, x) {
    let best = 0;
    bounds.forEach((bound, index) => {
      if (Math.abs(bound - x) < Math.abs(bounds[best] - x)) best = index;
    });
    return best;
  }

  // Kéo bảng vào vùng in được; phần thừa lấy từ các cột trống trước, sau đó co đều
  function clampBounds(bounds, frame, flexibleColumns) {
    const first = bounds[0];
    const last = bounds[bounds.length - 1];
    const left = Math.max(first, frame.safeLeft);
    const right = Math.min(last, frame.safeRight);
    let excess = last - first - (right - left);
    if (excess <= 0.5) return bounds;

    const widths = bounds.slice(1).map((bound, index) => bound - bounds[index]);
    for (const index of [...flexibleColumns].sort((a, b) => widths[b] - widths[a])) {
      const take = Math.min(Math.max(0, widths[index] - 1), excess);
      widths[index] -= take;
      excess -= take;
      if (excess <= 0) break;
    }
    if (excess > 0) {
      const total = sum(widths);
      widths.forEach((width, index) => {
        widths[index] = (width * (total - excess)) / total;
      });
    }

    const result = [left];
    widths.forEach((width) => result.push(result[result.length - 1] + width));
    return result;
  }

  function toColumns(bounds) {
    const offsets = bounds.map((bound) => twip(bound - bounds[0]));
    return offsets.slice(1).map((offset, index) => Math.max(1, offset - offsets[index]));
  }

  // Thêm ô trống vào chỗ hở để mỗi dòng luôn phủ đủ lưới cột
  function fillGaps(entries, columnCount) {
    const filled = [];
    let column = 0;
    for (const entry of [...entries].sort((a, b) => a.start - b.start)) {
      if (entry.start < column) continue;
      if (entry.start > column) filled.push({ start: column, span: entry.start - column });
      filled.push(entry);
      column = entry.start + entry.span;
    }
    if (column < columnCount) filled.push({ start: column, span: columnCount - column });
    return filled;
  }

  /* ---------- Bảng ---------- */

  function tableXml(columns, indent, rows) {
    const nil = ["top", "left", "bottom", "right", "insideH", "insideV"]
      .map((side) => `<w:${side} w:val="nil"/>`)
      .join("");
    const zeroMargins = ["top", "left", "bottom", "right"]
      .map((side) => `<w:${side} w:w="0" w:type="dxa"/>`)
      .join("");
    return (
      `<w:tbl><w:tblPr><w:tblW w:w="${sum(columns)}" w:type="dxa"/>` +
      `<w:tblInd w:w="${indent}" w:type="dxa"/><w:tblBorders>${nil}</w:tblBorders>` +
      `<w:tblLayout w:type="fixed"/><w:tblCellMar>${zeroMargins}</w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${columns.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>` +
      `${rows.join("")}</w:tbl>`
    );
  }

  function cellXml({ width, span = 1, merge = null, borders = null, margins = null, vAlign = null, content }) {
    let props = `<w:tcW w:w="${width}" w:type="dxa"/>`;
    if (span > 1) props += `<w:gridSpan w:val="${span}"/>`;
    if (merge) props += merge === "restart" ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>";
    if (borders) {
      props += `<w:tcBorders>${borders
        .map(([side, border]) =>
          border
            ? `<w:${side} w:val="${border.val}" w:sz="${border.sz}" w:space="0" w:color="${border.color}"/>`
            : `<w:${side} w:val="nil"/>`,
        )
        .join("")}</w:tcBorders>`;
    }
    if (margins && Object.values(margins).some((value) => value > 0)) {
      props += `<w:tcMar>${SIDES.map(
        ([side]) => `<w:${side} w:w="${twip(margins[side])}" w:type="dxa"/>`,
      ).join("")}</w:tcMar>`;
    }
    if (vAlign) props += `<w:vAlign w:val="${vAlign}"/>`;
    return `<w:tc><w:tcPr>${props}</w:tcPr>${content}</w:tc>`;
  }

  // Ô bảng phải kết thúc bằng một đoạn văn
  function cellBlocksXml(blocks, emptyXml) {
    if (!blocks.length) return emptyXml;
    const xml = blocksXml(blocks);
    return blocks[blocks.length - 1].type === "p" ? xml : xml + TINY_PARAGRAPH;
  }

  function emptyParagraph(style) {
    const fontPx = num(style.fontSize);
    const lineHeight =
      style.lineHeight === "normal" ? fontPx * NORMAL_LINE_HEIGHT : num(style.lineHeight);
    return paragraphXml({
      type: "p",
      jc: alignOf(style),
      indLeft: 0,
      indRight: 0,
      firstLine: 0,
      before: 0,
      after: 0,
      line: lineSpacing(lineHeight, fontPx),
      lineRule: "auto",
      runs: [],
      markSize: halfPoints(fontPx),
    });
  }

  function emitTable(table, frame, out) {
    const rows = [...table.rows].filter((row) => !isSkipped(row));
    const grid = rows.map((row) =>
      [...row.cells]
        .filter((cell) => !isSkipped(cell))
        .map((cell) => ({ cell, rect: cell.getBoundingClientRect() })),
    );
    const cells = grid.flat();
    if (!cells.length) return;

    pushGap(out, frame, Math.min(...cells.map(({ rect }) => rect.top)));

    const bounds = mergeBounds(cells.flatMap(({ rect }) => [rect.left, rect.right]));
    const fitted = clampBounds(bounds, frame, []);
    const columns = toColumns(fitted);
    const columnAt = (x) => nearestIndex(bounds, x);
    const merges = [];

    const rowsXml = grid.map((rowCells, rowIndex) => {
      const entries = rowCells.map(({ cell, rect }) => {
        const start = columnAt(rect.left);
        const span = Math.max(1, columnAt(rect.right) - start);
        if (cell.rowSpan > 1) {
          merges.push({ start, span, cell, from: rowIndex + 1, to: rowIndex + cell.rowSpan - 1 });
        }
        return { start, span, cell, rect, merge: cell.rowSpan > 1 ? "restart" : null };
      });
      merges
        .filter(({ from, to }) => rowIndex >= from && rowIndex <= to)
        .forEach((merge) => entries.push({ ...merge, merge: "continue" }));

      const cellsXml = fillGaps(entries, columns.length).map((entry) => {
        const width = sum(columns.slice(entry.start, entry.start + entry.span));
        if (!entry.cell) return cellXml({ width, span: entry.span, content: TINY_PARAGRAPH });

        const style = styleOf(entry.cell);
        let content = TINY_PARAGRAPH;
        if (entry.merge !== "continue") {
          const inner = contentBox(entry.cell, entry.rect, style);
          const blocks = [];
          emitContents(entry.cell, makeFrame(inner.left, inner.right), blocks);
          content = cellBlocksXml(blocks, emptyParagraph(style));
        }
        return cellXml({
          width,
          span: entry.span,
          merge: entry.merge,
          borders: SIDES.map(([side, name]) => [side, borderOf(style, name)]),
          margins: paddingOf(style),
          vAlign: VERTICAL_ALIGN[style.verticalAlign] || "top",
          content,
        });
      });

      // Word cộng lề trên/dưới của ô vào chiều cao dòng -> trừ ra để dòng cao đúng như trên trang
      const row = rows[rowIndex];
      const inset = Math.max(
        0,
        ...entries
          .filter((entry) => entry.cell)
          .map(({ cell }) => {
            const style = styleOf(cell);
            return (
              num(style.paddingTop) +
              num(style.paddingBottom) +
              (num(style.borderTopWidth) + num(style.borderBottomWidth)) / 2
            );
          }),
      );
      const height = twip(Math.max(1, row.getBoundingClientRect().height - inset));
      const header = row.parentElement?.tagName === "THEAD" ? "<w:tblHeader/>" : "";
      return (
        `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${height}" w:hRule="atLeast"/>${header}</w:trPr>` +
        `${cellsXml.join("")}</w:tr>`
      );
    });

    out.push({ type: "tbl", xml: tableXml(columns, twip(fitted[0] - frame.left), rowsXml) });
    frame.cursor = table.getBoundingClientRect().bottom;
  }

  /* ---------- Khối dàn ngang (flex / grid) -> bảng không viền ---------- */

  function layoutItems(element, style) {
    const isFlexRow = /flex$/.test(style.display) && !style.flexDirection.startsWith("column");
    if (!isFlexRow && !/grid$/.test(style.display)) return null;

    const items = [];
    for (const node of element.childNodes) {
      let item = null;
      if (isText(node)) {
        if (hasContent(node)) item = { node, text: true };
      } else if (isElement(node) && !isSkipped(node) && styleOf(node).position !== "absolute") {
        item = { node, text: false };
      }
      if (!item) continue;
      item.rect = rectOf(item);
      if (item.rect.width >= 1) items.push(item);
    }
    if (items.length < 2) return null;

    // Các ô xếp chồng thẳng hàng (lưới một cột) thì xử lý như khối thường
    const [first] = items;
    const stacked = items.every(
      ({ rect }) =>
        Math.abs(rect.left - first.rect.left) < EPSILON_PX &&
        Math.abs(rect.right - first.rect.right) < EPSILON_PX,
    );
    return stacked ? null : items;
  }

  function itemAlign(item) {
    const element = item.text ? item.node.parentElement : item.node;
    const style = styleOf(element);
    if (!item.text && isCenteredColumn(style)) return "center";
    const jc = alignOf(style);
    return jc === "both" ? "left" : jc;
  }

  function clusterRows(items) {
    const rows = [];
    for (const item of [...items].sort((a, b) => a.rect.top - b.rect.top)) {
      const row = rows[rows.length - 1];
      if (row && item.rect.top < row.bottom - EPSILON_PX) {
        row.items.push(item);
        row.bottom = Math.max(row.bottom, item.rect.bottom);
      } else {
        rows.push({ items: [item], bottom: item.rect.bottom });
      }
    }
    return rows.map((row) => row.items.sort((a, b) => a.rect.left - b.rect.left));
  }

  // Nới ô sang khoảng trống bên cạnh (phía không làm đổi vị trí chữ)
  // để Word không xuống dòng khi chữ đo ra rộng hơn trình duyệt một chút.
  // Khối bị dồn sát mép (vd. "Chức vụ: ..." ở mép phải) thì nới phía còn lại và
  // neo chữ về mép đó (item.anchor) để chữ vẫn nằm đúng chỗ.
  function addSlack(rowItems, outerLeft, outerRight) {
    const count = rowItems.length;
    const gaps = [];
    for (let index = 0; index <= count; index += 1) {
      const left = index > 0 ? rowItems[index - 1].rect.right : outerLeft;
      const right = index < count ? rowItems[index].rect.left : outerRight;
      gaps.push(Math.max(0, right - left));
    }

    rowItems.forEach((item, index) => {
      const roomLeft = gaps[index] >= MIN_ROOM_PX;
      const roomRight = gaps[index + 1] >= MIN_ROOM_PX;
      if (item.align === "center") {
        item.growLeft = item.growRight = true;
      } else if (item.align === "left") {
        item.growLeft = !roomRight && roomLeft;
        item.growRight = !item.growLeft;
        if (item.growLeft) item.anchor = "right";
      } else {
        item.growRight = !roomLeft && roomRight;
        item.growLeft = !item.growRight;
        if (item.growRight) item.anchor = "left";
      }
    });

    const limit = (item) => (item.align === "center" ? CENTER_SLACK_PX : SLACK_PX);
    for (let index = 0; index <= count; index += 1) {
      const previous = rowItems[index - 1];
      const next = rowItems[index];
      const giveLeft = previous?.growRight;
      const giveRight = next?.growLeft;
      const share = giveLeft && giveRight ? gaps[index] / 2 : gaps[index];
      if (giveLeft) previous.cellRight = previous.rect.right + Math.min(share, limit(previous));
      if (giveRight) next.cellLeft = next.rect.left - Math.min(share, limit(next));
    }
  }

  function verticalPosition(rect, top, bottom) {
    const above = rect.top - top;
    const below = bottom - rect.bottom;
    if (Math.abs(above - below) <= 2) return above > EPSILON_PX ? "center" : "top";
    return above > below ? "bottom" : "top";
  }

  function layoutCellContent(item) {
    const blocks = [];
    const growLeft = item.rect.left - item.cellLeft;
    const growRight = item.cellRight - item.rect.right;
    const { anchor = null } = item;
    if (item.text) {
      const actual = { left: item.cellLeft, right: item.cellRight };
      emitParagraph(
        { styleEl: item.node.parentElement, nodes: [item.node], hbox: item.rect },
        makeFrame(item.rect.left, item.rect.right, { actual, anchor }),
        blocks,
      );
    } else {
      const inner = contentBox(item.node);
      const actual = { left: inner.left - growLeft, right: inner.right + growRight };
      emitContents(item.node, makeFrame(inner.left, inner.right, { actual, anchor }), blocks);
    }
    return cellBlocksXml(blocks, TINY_PARAGRAPH);
  }

  function emitLayout(container, items, frame, out) {
    const box = contentBox(container);
    items.forEach((item) => {
      item.align = itemAlign(item);
      item.cellLeft = item.rect.left;
      item.cellRight = item.rect.right;
    });
    const rows = clusterRows(items);
    pushGap(out, frame, Math.min(...items.map((item) => item.rect.top)));

    const outerLeft = Math.min(frame.left, box.left, ...items.map((item) => item.rect.left));
    const outerRight = Math.max(frame.right, box.right, ...items.map((item) => item.rect.right));
    if (rows.length === 1) addSlack(rows[0], outerLeft, outerRight);

    const bounds = mergeBounds([
      outerLeft,
      outerRight,
      ...items.flatMap((item) => [item.cellLeft, item.cellRight]),
    ]);
    const used = new Set();
    for (const item of items) {
      item.start = nearestIndex(bounds, item.cellLeft);
      item.span = Math.max(1, nearestIndex(bounds, item.cellRight) - item.start);
      for (let column = item.start; column < item.start + item.span; column += 1) used.add(column);
    }
    const spacers = bounds.slice(1).map((_, column) => column).filter((column) => !used.has(column));
    const fitted = clampBounds(bounds, frame, spacers);
    const columns = toColumns(fitted);

    const rowsXml = rows.map((rowItems) => {
      const top = Math.min(...rowItems.map((item) => item.rect.top));
      const bottom = Math.max(...rowItems.map((item) => item.rect.bottom));
      const cells = fillGaps(
        rowItems.map((item) => ({ start: item.start, span: item.span, item })),
        columns.length,
      ).map(({ start, span, item }) => {
        const width = sum(columns.slice(start, start + span));
        if (!item) return cellXml({ width, span, content: TINY_PARAGRAPH });
        return cellXml({
          width,
          span,
          margins: item.text ? null : paddingOf(styleOf(item.node)),
          vAlign: verticalPosition(item.rect, top, bottom),
          content: layoutCellContent(item),
        });
      });
      return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells.join("")}</w:tr>`;
    });

    out.push({ type: "tbl", xml: tableXml(columns, twip(fitted[0] - frame.left), rowsXml) });
    frame.cursor = Math.max(box.bottom, ...items.map((item) => item.rect.bottom));
  }

  /* ---------- Ghi XML ---------- */

  function runPropsXml(props) {
    let xml = "";
    if (props.font !== DEFAULT_FONT) {
      const font = escapeXml(props.font);
      xml += `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:eastAsia="${font}"/>`;
    }
    if (props.bold) xml += "<w:b/><w:bCs/>";
    if (props.italic) xml += "<w:i/><w:iCs/>";
    if (props.strike) xml += "<w:strike/>";
    if (props.hidden) xml += "<w:vanish/>";
    if (props.color) xml += `<w:color w:val="${props.color}"/>`;
    if (props.spacing) xml += `<w:spacing w:val="${props.spacing}"/>`;
    xml += `<w:sz w:val="${props.size}"/><w:szCs w:val="${props.size}"/>`;
    if (props.underline) xml += '<w:u w:val="single"/>';
    if (props.vertAlign) xml += `<w:vertAlign w:val="${props.vertAlign}"/>`;
    return `<w:rPr>${xml}</w:rPr>`;
  }

  function runXml(run) {
    if (run.br) return "<w:r><w:br/></w:r>";
    // Word không gộp dấu cách nên dùng dấu cách thường thay cho &nbsp; (chỗ để trống ngày tháng...):
    // độ rộng giữ nguyên, nhưng Word được phép xuống dòng ở đó như trình duyệt
    const text = run.text.replace(/ /g, " ");
    return `<w:r>${runPropsXml(run.props)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  }

  function paragraphXml(paragraph) {
    let props = "";
    const { border } = paragraph;
    if (border) {
      props += `<w:pBdr><w:${border.side} w:val="${border.val}" w:sz="${border.sz}" w:space="0" w:color="${border.color}"/></w:pBdr>`;
    }
    props +=
      `<w:spacing w:before="${Math.max(0, twip(paragraph.before))}" w:after="${Math.max(0, twip(paragraph.after))}"` +
      ` w:line="${paragraph.line}" w:lineRule="${paragraph.lineRule}"/>`;

    const firstLine = twip(paragraph.firstLine);
    const indentFirst =
      firstLine > 0 ? ` w:firstLine="${firstLine}"` : firstLine < 0 ? ` w:hanging="${-firstLine}"` : "";
    props += `<w:ind w:left="${twip(paragraph.indLeft)}" w:right="${twip(paragraph.indRight)}"${indentFirst}/>`;
    props += `<w:jc w:val="${paragraph.jc}"/>`;
    props += `<w:rPr><w:sz w:val="${paragraph.markSize}"/><w:szCs w:val="${paragraph.markSize}"/></w:rPr>`;
    if (paragraph.sectPr) props += paragraph.sectPr;

    return `<w:p><w:pPr>${props}</w:pPr>${paragraph.runs.map(runXml).join("")}</w:p>`;
  }

  function blocksXml(blocks) {
    return blocks.map((block) => (block.type === "tbl" ? block.xml : paragraphXml(block))).join("");
  }

  /* ---------- Trang & tài liệu ---------- */

  function sectionXml(style) {
    const margin = (side) => {
      const px = num(style[`padding${side}`]);
      return twip(px > 0 ? px : FALLBACK_MARGIN_MM[side] * PX_PER_MM);
    };
    return (
      `<w:sectPr><w:pgSz w:w="${A4_TWIP.width}" w:h="${A4_TWIP.height}"/>` +
      `<w:pgMar w:top="${margin("Top")}" w:right="${margin("Right")}" w:bottom="${margin("Bottom")}"` +
      ` w:left="${margin("Left")}" w:header="567" w:footer="567" w:gutter="0"/>` +
      '<w:cols w:space="720"/></w:sectPr>'
    );
  }

  function convertPage(page) {
    const style = styleOf(page);
    const rect = page.getBoundingClientRect();
    const inner = contentBox(page, rect, style);
    const frame = makeFrame(inner.left, inner.right, {
      cursor: inner.top,
      safe: { left: rect.left + SAFE_EDGE_PX, right: rect.right - SAFE_EDGE_PX },
    });
    const blocks = [];
    emitContents(page, frame, blocks);
    return { blocks, sectPr: sectionXml(style) };
  }

  function buildDocumentXml(doc) {
    const pages = [...doc.querySelectorAll(".page")].filter((page) => !isSkipped(page));
    const sections = (pages.length ? pages : [doc.body]).map(convertPage);

    const body = sections
      .map(({ blocks, sectPr }, index) => {
        const isLast = index === sections.length - 1;
        if (blocks[blocks.length - 1]?.type !== "p") blocks.push(spacerParagraph(0));
        // Ngắt section (sang trang mới) gắn vào đoạn cuối của mỗi trang, trang cuối dùng sectPr của body
        if (!isLast) blocks[blocks.length - 1].sectPr = sectPr;
        return blocksXml(blocks) + (isLast ? sectPr : "");
      })
      .join("");

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<w:body>${body}</w:body></w:document>`
    );
  }

  const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  const CONTENT_TYPES_XML =
    XML_HEADER +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    "</Types>";

  const ROOT_RELS_XML =
    XML_HEADER +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    "</Relationships>";

  const DOCUMENT_RELS_XML =
    XML_HEADER +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
    "</Relationships>";

  const STYLES_XML =
    XML_HEADER +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
    `<w:rFonts w:ascii="${DEFAULT_FONT}" w:eastAsia="${DEFAULT_FONT}" w:hAnsi="${DEFAULT_FONT}" w:cs="${DEFAULT_FONT}"/>` +
    // Bật kerning như trình duyệt để chữ dài đúng bằng trên trang A4
    '<w:kern w:val="2"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="vi-VN" w:eastAsia="en-US" w:bidi="ar-SA"/>' +
    "</w:rPr></w:rPrDefault>" +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    "</w:docDefaults>" +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>' +
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    "</w:styles>";

  const SETTINGS_XML =
    XML_HEADER +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>' +
    '<w:characterSpacingControl w:val="doNotCompress"/>' +
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
    '<w:themeFontLang w:val="vi-VN"/>' +
    "</w:settings>";

  function coreXml(title) {
    const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    return (
      XML_HEADER +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
      ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
      ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${escapeXml(title)}</dc:title><dc:creator>${APP_NAME}</dc:creator>` +
      `<cp:lastModifiedBy>${APP_NAME}</cp:lastModifiedBy>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      "</cp:coreProperties>"
    );
  }

  /* ---------- Đóng gói ZIP ---------- */

  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  async function deflateRaw(bytes) {
    if (typeof CompressionStream === "undefined") return null;
    try {
      const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("deflate-raw"));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    } catch {
      return null; // trình duyệt cũ không có deflate-raw -> lưu không nén
    }
  }

  async function createZip(files) {
    const encoder = new TextEncoder();
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

    const parts = [];
    const directory = [];
    let offset = 0;

    for (const { name, content } of files) {
      const nameBytes = encoder.encode(name);
      const raw = encoder.encode(content);
      const deflated = await deflateRaw(raw);
      const compressed = deflated && deflated.length < raw.length;
      const data = compressed ? deflated : raw;
      const method = compressed ? 8 : 0;
      const crc = crc32(raw);

      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // tên file UTF-8
      local.setUint16(8, method, true);
      local.setUint16(10, dosTime, true);
      local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true);
      local.setUint32(22, raw.length, true);
      local.setUint16(26, nameBytes.length, true);
      parts.push(local, nameBytes, data);

      const entry = new DataView(new ArrayBuffer(46));
      entry.setUint32(0, 0x02014b50, true);
      entry.setUint16(4, 20, true);
      entry.setUint16(6, 20, true);
      entry.setUint16(8, 0x0800, true);
      entry.setUint16(10, method, true);
      entry.setUint16(12, dosTime, true);
      entry.setUint16(14, dosDate, true);
      entry.setUint32(16, crc, true);
      entry.setUint32(20, data.length, true);
      entry.setUint32(24, raw.length, true);
      entry.setUint16(28, nameBytes.length, true);
      entry.setUint32(42, offset, true);
      directory.push(entry, nameBytes);

      offset += 30 + nameBytes.length + data.length;
    }

    const directorySize = sum(directory.map((part) => part.byteLength));
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, directorySize, true);
    end.setUint32(16, offset, true);

    return new Blob([...parts, ...directory, end], { type: MIME });
  }

  /* ---------- Dựng hồ sơ trong iframe ẩn ---------- */

  async function renderInFrame(html) {
    const iframe = document.createElement("iframe");
    iframe.setAttribute("aria-hidden", "true");
    iframe.setAttribute("sandbox", "allow-same-origin"); // không chạy script trong bản chụp
    iframe.tabIndex = -1;
    Object.assign(iframe.style, {
      position: "fixed",
      left: "-10000px",
      top: "0",
      width: "1100px",
      height: "800px",
      border: "0",
      visibility: "hidden",
      pointerEvents: "none",
    });

    const loaded = new Promise((resolve, reject) => {
      iframe.addEventListener("load", resolve, { once: true });
      setTimeout(() => reject(new Error("Quá thời gian dựng hồ sơ.")), RENDER_TIMEOUT_MS);
    });
    iframe.srcdoc = html;
    document.body.append(iframe);

    try {
      await loaded;
      const doc = iframe.contentDocument;
      if (!doc?.body) throw new Error("Không đọc được nội dung hồ sơ.");
      await doc.fonts?.ready;
      return iframe;
    } catch (error) {
      iframe.remove();
      throw error;
    }
  }

  /**
   * Chuyển HTML hồ sơ (bản chụp như khi lưu lịch sử) thành file Word.
   * @returns {Promise<Blob>}
   */
  async function fromHtml(html, { title = "Hồ sơ" } = {}) {
    const iframe = await renderInFrame(html);
    try {
      const documentXml = buildDocumentXml(iframe.contentDocument);
      return await createZip([
        { name: "[Content_Types].xml", content: CONTENT_TYPES_XML },
        { name: "_rels/.rels", content: ROOT_RELS_XML },
        { name: "word/document.xml", content: documentXml },
        { name: "word/_rels/document.xml.rels", content: DOCUMENT_RELS_XML },
        { name: "word/styles.xml", content: STYLES_XML },
        { name: "word/settings.xml", content: SETTINGS_XML },
        { name: "docProps/core.xml", content: coreXml(title) },
      ]);
    } finally {
      iframe.remove();
    }
  }

  return { MIME, fromHtml };
})();
