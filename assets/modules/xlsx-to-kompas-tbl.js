import { KOMPAS_COM_ADAPTER } from "../executor-shell.js";

const DEFAULT_LAYOUT = {
  tableWidthMm: 260,
  tableHeightMm: 24,
  cellWidthMm: 30,
  cellHeightMm: 8,
};

const EXPORT_BATCH_SIZE = 200;
const API5_STRUCT_TYPE_PARAGRAPH_PARAM = 27;
const API5_STRUCT_TYPE_TEXT_PARAM = 28;
const API5_STRUCT_TYPE_TEXT_LINE_PARAM = 29;
const API5_STRUCT_TYPE_TEXT_ITEM_FONT = 30;
const API5_STRUCT_TYPE_TEXT_ITEM_PARAM = 31;
const API5_TEXT_LINE_ARRAY_TYPE = 3;
const API5_TEXT_ITEM_ARRAY_TYPE = 4;
const API5_TEXT_ITEM_STRING = 0;
const API5_TEXT_FLAG_ITALIC_ON = 0x40;
const API5_TEXT_FLAG_ITALIC_OFF = 0x80;
const API5_TEXT_FLAG_BOLD_ON = 0x100;
const API5_TEXT_FLAG_BOLD_OFF = 0x200;
const API5_TEXT_FLAG_UNDERLINE_ON = 0x400;
const API5_TEXT_FLAG_UNDERLINE_OFF = 0x800;
const KOMPAS_API7_ATTACHED_ARGUMENTS = {
  attachOnly: true,
  createIfMissing: false,
  visible: false,
  progIds: ["KOMPAS.Application.7"],
};
const KOMPAS_API5_ATTACHED_ARGUMENTS = {
  attachOnly: true,
  createIfMissing: false,
  visible: false,
  progIds: ["KOMPAS.Application.5"],
};
const KOMPAS_API7_OPEN_ARGUMENTS = {
  attachOnly: false,
  createIfMissing: true,
  visible: true,
  progIds: ["KOMPAS.Application.7"],
};

function stripExtension(fileName) {
  return String(fileName || "").replace(/\.[^.]+$/, "");
}

function normalizeWindowsPath(value) {
  return String(value || "")
    .replace(/\//g, "\\")
    .replace(/\\{2,}/g, "\\");
}

function joinWindowsPath(left, right) {
  const lhs = normalizeWindowsPath(left).replace(/[\\]+$/, "");
  const rhs = normalizeWindowsPath(right).replace(/^[\\]+/, "");
  if (!lhs) {
    return rhs;
  }
  if (!rhs) {
    return lhs;
  }
  return `${lhs}\\${rhs}`;
}

function dirname(filePath) {
  const value = normalizeWindowsPath(filePath);
  const index = value.lastIndexOf("\\");
  return index >= 0 ? value.slice(0, index) : "";
}

function sanitizeFileStem(fileName) {
  const stem = stripExtension(fileName || "table")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .trim();
  return stem || "table";
}

function ensureTblExtension(filePath) {
  const value = normalizeWindowsPath(filePath);
  return /\.tbl$/i.test(value) ? value : `${value}.tbl`;
}

function formatSize(rows, cols) {
  return `${rows} x ${cols}`;
}

function formatMm(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return "0";
  }
  return numeric.toFixed(4).replace(/\.?0+$/, "");
}

function roundLayoutValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return 0;
  }
  return Number(numeric.toFixed(4));
}

function coerceNumberLike(value) {
  const direct = Number(value);
  if (Number.isFinite(direct)) {
    return direct;
  }
  if (value && typeof value === "object") {
    for (const key of ["value", "Value", "reference", "Reference", "result", "Result", "rawValue", "RawValue", "currentValue", "CurrentValue"]) {
      const nested = Number(value[key]);
      if (Number.isFinite(nested)) {
        return nested;
      }
    }
  }
  return Number.NaN;
}

function readNumeric(input, fallback = 0) {
  const numeric = Number.parseFloat(input.value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function formatCellValue(cell) {
  if (!cell) {
    return "";
  }
  if (typeof cell.w === "string") {
    return cell.w;
  }
  if (cell.v === null || cell.v === undefined) {
    return "";
  }
  if (typeof cell.v === "boolean") {
    return cell.v ? "True" : "False";
  }
  return String(cell.v);
}

const XML_ATTR_RE = /([^"\s?>\/=]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^'">\s]+))/g;
const POINTS_TO_MM = 25.4 / 72;
const LINE_BREAK_RE = /\r\n|\n|\r/g;
const DEFAULT_FONT_NAME = "Calibri";
const DEFAULT_FONT_SIZE_PT = 11;
const DEFAULT_FONT_RGB = "000000";
const DEFAULT_KOMPAS_ALIGN = 0;
const KOMPAS_ALIGN_BY_HORIZONTAL = {
  center: 1,
  centerContinuous: 1,
  right: 2,
};
const XML_ENTITY_BY_TOKEN = {
  "&amp;": "&",
  "&apos;": "'",
  "&gt;": ">",
  "&lt;": "<",
  "&quot;": "\"",
};

function normalizeZipPath(value) {
  return String(value || "")
    .replace(/\\/g, "/")
    .replace(/^\/+/, "")
    .replace(/\/{2,}/g, "/");
}

function resolveZipTarget(baseDir, targetPath) {
  const target = normalizeZipPath(targetPath);
  if (!target) {
    return "";
  }
  if (String(targetPath || "").startsWith("/")) {
    return target;
  }
  const base = normalizeZipPath(baseDir).replace(/\/+$/, "");
  return base ? normalizeZipPath(`${base}/${target}`) : target;
}

function toUint8Array(value) {
  if (value instanceof Uint8Array) {
    return value;
  }
  if (Array.isArray(value)) {
    return Uint8Array.from(value);
  }
  if (value && typeof value === "object" && value.buffer instanceof ArrayBuffer) {
    return new Uint8Array(value.buffer, value.byteOffset || 0, value.byteLength || value.length || 0);
  }
  return new Uint8Array();
}

function readWorkbookFileText(workbook, filePath) {
  const entry = workbook?.files?.[normalizeZipPath(filePath)];
  if (!entry?.content) {
    return "";
  }
  if (typeof entry.content === "string") {
    return entry.content;
  }
  return new TextDecoder("utf-8").decode(toUint8Array(entry.content));
}

function parseXmlAttributes(source) {
  const attributes = {};
  const text = String(source || "");
  let match;
  XML_ATTR_RE.lastIndex = 0;
  while ((match = XML_ATTR_RE.exec(text))) {
    const key = String(match[1] || "");
    if (!key) {
      continue;
    }
    attributes[key] = match[2] ?? match[3] ?? match[4] ?? "";
  }
  return attributes;
}

function parseXmlTag(token) {
  const match = String(token || "").match(/^<\s*(\/?)\s*(?:[\w-]+:)?([\w-]+)/);
  if (!match) {
    return null;
  }
  return {
    closing: Boolean(match[1]),
    name: match[2],
    attrs: parseXmlAttributes(token),
  };
}

function extractXmlSection(xml, tagName) {
  const match = String(xml || "").match(new RegExp(`<(?:\\w+:)?${tagName}\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?${tagName}>`, "i"));
  return match ? match[1] : "";
}

function extractXmlBlocks(xml, tagName) {
  const matches = [];
  const expression = new RegExp(
    `<(?:\\w+:)?${tagName}\\b[^>]*/>|<(?:\\w+:)?${tagName}\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?${tagName}>`,
    "gi",
  );
  let match;
  while ((match = expression.exec(String(xml || "")))) {
    matches.push({
      full: match[0],
      inner: match[1] || "",
    });
  }
  return matches;
}

function decodeXmlText(value) {
  return String(value || "")
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/g, (token, codePoint) => {
      if (XML_ENTITY_BY_TOKEN[token]) {
        return XML_ENTITY_BY_TOKEN[token];
      }
      const base = token.includes("#x") ? 16 : 10;
      const numeric = Number.parseInt(codePoint, base);
      return Number.isFinite(numeric) ? String.fromCharCode(numeric) : token;
    })
    .replace(/_x([\da-fA-F]{4})_/g, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16)));
}

function normalizeHexColor(value) {
  const hex = String(value || "").replace(/[^0-9a-f]/gi, "").toUpperCase();
  if (hex.length >= 6) {
    return hex.slice(-6);
  }
  return "";
}

function hexColorToRgb(hex) {
  const normalized = normalizeHexColor(hex) || DEFAULT_FONT_RGB;
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16),
  ];
}

function rgbToHex(rgb) {
  return rgb
    .map((value) => {
      const clamped = Math.max(0, Math.min(255, Math.round(Number(value) || 0)));
      return clamped.toString(16).padStart(2, "0");
    })
    .join("")
    .toUpperCase();
}

function rgbToHsl(rgb) {
  const red = rgb[0] / 255;
  const green = rgb[1] / 255;
  const blue = rgb[2] / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const delta = max - min;
  if (delta === 0) {
    return [0, 0, red];
  }
  let hue = 0;
  const lightness = (max + min) / 2;
  const saturation = delta / (lightness > 0.5 ? 2 - max - min : max + min);
  switch (max) {
    case red:
      hue = ((green - blue) / delta + 6) % 6;
      break;
    case green:
      hue = (blue - red) / delta + 2;
      break;
    default:
      hue = (red - green) / delta + 4;
      break;
  }
  return [hue / 6, saturation, lightness];
}

function hslToRgb(hsl) {
  const [hue, saturation, lightness] = hsl;
  const spread = saturation * 2 * (lightness < 0.5 ? lightness : 1 - lightness);
  const base = lightness - spread / 2;
  const channels = [base, base, base];
  const scaledHue = 6 * hue;
  let value;
  if (saturation !== 0) {
    switch (Math.floor(scaledHue)) {
      case 0:
      case 6:
        value = spread * scaledHue;
        channels[0] += spread;
        channels[1] += value;
        break;
      case 1:
        value = spread * (2 - scaledHue);
        channels[0] += value;
        channels[1] += spread;
        break;
      case 2:
        value = spread * (scaledHue - 2);
        channels[1] += spread;
        channels[2] += value;
        break;
      case 3:
        value = spread * (4 - scaledHue);
        channels[1] += value;
        channels[2] += spread;
        break;
      case 4:
        value = spread * (scaledHue - 4);
        channels[2] += spread;
        channels[0] += value;
        break;
      default:
        value = spread * (6 - scaledHue);
        channels[2] += value;
        channels[0] += spread;
        break;
    }
  }
  return channels.map((channel) => Math.round(channel * 255));
}

function applyThemeTint(rgbHex, tint) {
  const numericTint = Number.parseFloat(tint);
  const normalized = normalizeHexColor(rgbHex) || DEFAULT_FONT_RGB;
  if (!Number.isFinite(numericTint) || numericTint === 0) {
    return normalized;
  }
  const hsl = rgbToHsl(hexColorToRgb(normalized));
  if (numericTint < 0) {
    hsl[2] *= 1 + numericTint;
  } else {
    hsl[2] = 1 - (1 - hsl[2]) * (1 - numericTint);
  }
  return rgbToHex(hslToRgb(hsl));
}

function isXmlTruthy(value) {
  return /^(1|true)$/i.test(String(value || ""));
}

function buildThemeColorLookup(themes) {
  return Array.isArray(themes?.themeElements?.clrScheme)
    ? themes.themeElements.clrScheme.map((entry) => normalizeHexColor(entry?.rgb || entry?.lastClr || "") || DEFAULT_FONT_RGB)
    : [];
}

function parseIndexedColors(stylesXml) {
  const indexedSection = extractXmlSection(stylesXml, "indexedColors");
  const colors = [];
  const blocks = extractXmlBlocks(indexedSection, "rgbColor");
  for (const block of blocks) {
    const openTag = block.full.match(/<[^>]+>/)?.[0] || block.full;
    const attrs = parseXmlAttributes(openTag);
    colors.push(normalizeHexColor(attrs.rgb) || "");
  }
  return colors;
}

function resolveColorValue(attrs, styleContext) {
  if (attrs.rgb) {
    return normalizeHexColor(attrs.rgb) || DEFAULT_FONT_RGB;
  }
  if (attrs.indexed !== undefined || attrs.index !== undefined) {
    const colorIndex = Number.parseInt(attrs.indexed ?? attrs.index, 10);
    return styleContext.indexedColors[colorIndex] || DEFAULT_FONT_RGB;
  }
  if (attrs.theme !== undefined) {
    const themeIndex = Number.parseInt(attrs.theme, 10);
    const baseColor = styleContext.themeColors[themeIndex] || DEFAULT_FONT_RGB;
    return applyThemeTint(baseColor, attrs.tint);
  }
  if (attrs.auto !== undefined) {
    return DEFAULT_FONT_RGB;
  }
  return "";
}

function normalizeFontDescriptor(font) {
  const heightPt = Number.parseFloat(font?.heightPt ?? font?.sz ?? DEFAULT_FONT_SIZE_PT);
  return {
    fontName: String(font?.fontName || font?.name || DEFAULT_FONT_NAME),
    heightPt: Number.isFinite(heightPt) && heightPt > 0 ? heightPt : DEFAULT_FONT_SIZE_PT,
    bold: Boolean(font?.bold),
    italic: Boolean(font?.italic),
    underline: Boolean(font?.underline),
    colorHex: normalizeHexColor(font?.colorHex || font?.rgb || "") || DEFAULT_FONT_RGB,
    widthFactor: Number.isFinite(Number(font?.widthFactor)) ? Number(font.widthFactor) : 1,
  };
}

function mergeFontDescriptors(baseFont, overrideFont) {
  return normalizeFontDescriptor({
    ...baseFont,
    ...overrideFont,
  });
}

function parseFontDescriptor(fontXml, styleContext) {
  const font = {};
  for (const token of String(fontXml || "").match(/<[^>]+>/g) || []) {
    const tag = parseXmlTag(token);
    if (!tag || tag.closing) {
      continue;
    }
    switch (tag.name) {
      case "b":
        font.bold = tag.attrs.val === undefined ? true : !/^(0|false)$/i.test(String(tag.attrs.val));
        break;
      case "color": {
        const colorHex = resolveColorValue(tag.attrs, styleContext);
        if (colorHex) {
          font.colorHex = colorHex;
        }
        break;
      }
      case "i":
        font.italic = tag.attrs.val === undefined ? true : !/^(0|false)$/i.test(String(tag.attrs.val));
        break;
      case "name":
      case "rFont":
        if (tag.attrs.val) {
          font.fontName = decodeXmlText(tag.attrs.val);
        }
        break;
      case "sz":
        if (tag.attrs.val !== undefined) {
          font.heightPt = Number.parseFloat(tag.attrs.val);
        }
        break;
      case "u":
        font.underline = tag.attrs.val === undefined ? true : !/^(0|false|none)$/i.test(String(tag.attrs.val));
        break;
      default:
        break;
    }
  }
  return font;
}

function parseFontDescriptors(stylesXml, styleContext) {
  return extractXmlBlocks(extractXmlSection(stylesXml, "fonts"), "font")
    .map((block) => normalizeFontDescriptor(parseFontDescriptor(block.full, styleContext)));
}

function parseCellStyleDescriptors(stylesXml) {
  return extractXmlBlocks(extractXmlSection(stylesXml, "cellXfs"), "xf").map((block) => {
    const openTag = block.full.match(/<[^>]+>/)?.[0] || "";
    const attrs = parseXmlAttributes(openTag);
    const alignmentMatch = block.full.match(/<(?:\w+:)?alignment\b[^>]*\/?>/i);
    const alignmentAttrs = alignmentMatch ? parseXmlAttributes(alignmentMatch[0]) : {};
    const fontId = Number.parseInt(attrs.fontId, 10);
    return {
      fontId: Number.isInteger(fontId) ? fontId : 0,
      alignment: {
        horizontal: String(alignmentAttrs.horizontal || ""),
        wrapText: isXmlTruthy(alignmentAttrs.wrapText),
      },
    };
  });
}

function buildWorkbookStyleContext(workbook) {
  const stylesXml = readWorkbookFileText(workbook, "xl/styles.xml");
  const styleContext = {
    themeColors: buildThemeColorLookup(workbook?.Themes),
    indexedColors: parseIndexedColors(stylesXml),
    fonts: [],
    cellXfs: [],
    defaultFont: normalizeFontDescriptor({}),
  };
  styleContext.fonts = parseFontDescriptors(stylesXml, styleContext);
  styleContext.defaultFont = styleContext.fonts[0] || normalizeFontDescriptor({});
  styleContext.cellXfs = parseCellStyleDescriptors(stylesXml);
  return styleContext;
}

function parseRelationshipTargets(xml, baseDir) {
  const byId = new Map();
  const expression = /<(?:\w+:)?Relationship\b[^>]*\/?>/gi;
  let match;
  while ((match = expression.exec(String(xml || "")))) {
    const attrs = parseXmlAttributes(match[0]);
    if (!attrs.Id || !attrs.Target) {
      continue;
    }
    byId.set(attrs.Id, resolveZipTarget(baseDir, attrs.Target));
  }
  return byId;
}

function resolveWorksheetFilePath(workbook, sheetIndex) {
  const relationshipXml = readWorkbookFileText(workbook, "xl/_rels/workbook.xml.rels");
  const relationships = parseRelationshipTargets(relationshipXml, "xl");
  const sheetId = workbook?.Workbook?.Sheets?.[sheetIndex]?.id;
  return relationships.get(sheetId) || `xl/worksheets/sheet${sheetIndex + 1}.xml`;
}

function parseWorksheetCellMeta(worksheetXml) {
  const byAddress = {};
  const expression = /<(?:\w+:)?c\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>|<(?:\w+:)?c\b([^>]*)\/>/gi;
  let match;
  while ((match = expression.exec(String(worksheetXml || "")))) {
    const attrSource = match[1] ?? match[3] ?? "";
    const innerXml = match[2] || "";
    const attrs = parseXmlAttributes(attrSource);
    const address = String(attrs.r || "");
    if (!address) {
      continue;
    }
    const styleIndex = Number.parseInt(attrs.s, 10);
    const inlineRichMatch = innerXml.match(/<(?:\w+:)?is\b[^>]*>([\s\S]*?)<\/(?:\w+:)?is>/i);
    const rawValueMatch = innerXml.match(/<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/i);
    byAddress[address] = {
      styleIndex: Number.isInteger(styleIndex) ? styleIndex : 0,
      type: String(attrs.t || ""),
      inlineRichXml: inlineRichMatch ? inlineRichMatch[1] : "",
      rawValue: rawValueMatch ? decodeXmlText(rawValueMatch[1]) : "",
    };
  }
  return byAddress;
}

function readRichTextPlainText(richTextXml) {
  const expression = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/gi;
  const parts = [];
  let match;
  while ((match = expression.exec(String(richTextXml || "")))) {
    parts.push(decodeXmlText(match[1]));
  }
  return parts.join("");
}

function parseRichTextRuns(richTextXml, baseFont, styleContext) {
  const source = String(richTextXml || "").replace(/<(?:\w+:)?rPh[\s\S]*?<\/(?:\w+:)?rPh>/gi, "");
  const runBlocks = extractXmlBlocks(source, "r");
  if (!runBlocks.length) {
    const plainText = readRichTextPlainText(source);
    return plainText === "" ? [] : [{
      ...baseFont,
      text: plainText,
    }];
  }

  return runBlocks
    .map((block) => {
      const text = readRichTextPlainText(block.full);
      const richStyleMatch = block.full.match(/<(?:\w+:)?rPr\b[^>]*>([\s\S]*?)<\/(?:\w+:)?rPr>/i);
      const richStyle = richStyleMatch ? parseFontDescriptor(richStyleMatch[1], styleContext) : {};
      return {
        ...mergeFontDescriptors(baseFont, richStyle),
        text,
      };
    })
    .filter((run) => run.text !== "");
}

function pointsToMillimeters(value) {
  const numeric = Number.parseFloat(value);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return roundLayoutValue(DEFAULT_FONT_SIZE_PT * POINTS_TO_MM);
  }
  return roundLayoutValue(numeric * POINTS_TO_MM);
}

function toKompasColor(rgbHex) {
  const normalized = normalizeHexColor(rgbHex) || DEFAULT_FONT_RGB;
  return Number.parseInt(normalized, 16);
}

function toKompasAlign(horizontal) {
  return KOMPAS_ALIGN_BY_HORIZONTAL[String(horizontal || "")] ?? DEFAULT_KOMPAS_ALIGN;
}

function createTransferItem(run) {
  return {
    text: String(run?.text ?? ""),
    fontName: String(run?.fontName || DEFAULT_FONT_NAME),
    heightMm: pointsToMillimeters(run?.heightPt),
    bold: Boolean(run?.bold),
    italic: Boolean(run?.italic),
    underline: Boolean(run?.underline),
    color: toKompasColor(run?.colorHex),
    widthFactor: Number.isFinite(Number(run?.widthFactor)) ? Number(run.widthFactor) : 1,
  };
}

function splitRunsIntoLines(runs, baseFont) {
  const lines = [{ items: [] }];
  for (const run of runs) {
    const chunks = String(run?.text ?? "").split(LINE_BREAK_RE);
    for (let index = 0; index < chunks.length; index += 1) {
      if (index > 0) {
        lines.push({ items: [] });
      }
      const chunk = chunks[index];
      if (chunk !== "" || chunks.length === 1) {
        lines[lines.length - 1].items.push(createTransferItem({
          ...run,
          text: chunk,
        }));
      }
    }
  }
  return lines.map((line) => ({
    items: line.items.length ? line.items : [createTransferItem({ ...baseFont, text: "" })],
  }));
}

function createCellTransferPayload({
  address,
  columnIndex,
  rowIndex,
  text,
  styleIndex,
  richTextXml,
  styleContext,
  sharedStringEntry,
}) {
  const cellStyle = styleContext.cellXfs[styleIndex] || {};
  const baseFont = mergeFontDescriptors(
    styleContext.defaultFont,
    styleContext.fonts[cellStyle.fontId] || {},
  );
  const richSource = richTextXml || sharedStringEntry?.r || "";
  const displayText = String(text ?? sharedStringEntry?.t ?? "");
  const runs = richSource
    ? parseRichTextRuns(richSource, baseFont, styleContext)
    : (displayText === "" ? [] : [{ ...baseFont, text: displayText }]);
  const lines = runs.length ? splitRunsIntoLines(runs, baseFont) : [];
  return {
    address,
    rowIndex,
    columnIndex,
    text: displayText,
    horizontal: String(cellStyle.alignment?.horizontal || ""),
    alignCode: toKompasAlign(cellStyle.alignment?.horizontal),
    wrapText: Boolean(cellStyle.alignment?.wrapText),
    oneLine: !cellStyle.alignment?.wrapText && lines.length <= 1,
    lines,
    hasContent: displayText !== "" || lines.length > 0,
  };
}

function collectFormattedCellWrites(cellMatrix) {
  const writes = [];
  const { rows, cols } = getMatrixDimensions(cellMatrix);
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = Array.isArray(cellMatrix[rowIndex]) ? cellMatrix[rowIndex] : [];
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      const cell = row[columnIndex];
      if (!cell?.hasContent) {
        continue;
      }
      writes.push(cell);
    }
  }
  return writes;
}

function createFormattedCellCommands(cell) {
  const commands = [
    {
      commandId: "xlsx-to-kompas-tbl.table-cell-clear-text",
      arguments: {
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
      },
      timeoutMilliseconds: 30000,
    },
    {
      commandId: "xlsx-to-kompas-tbl.table-cell-set-one-line",
      arguments: {
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        oneLine: Boolean(cell.oneLine),
      },
      timeoutMilliseconds: 30000,
    },
  ];

  cell.lines.forEach((line, lineIndex) => {
    commands.push({
      commandId: "xlsx-to-kompas-tbl.table-cell-add-line",
      arguments: {
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        align: cell.alignCode,
      },
      timeoutMilliseconds: 30000,
    });
    line.items.forEach((item, itemIndex) => {
      commands.push({
        commandId: itemIndex === 0
          ? "xlsx-to-kompas-tbl.table-cell-add-item-before"
          : "xlsx-to-kompas-tbl.table-cell-add-item",
        arguments: {
          rowIndex: cell.rowIndex,
          columnIndex: cell.columnIndex,
          lineIndex,
          itemIndex,
          value: item.text,
          fontName: item.fontName,
          heightMm: item.heightMm,
          bold: item.bold,
          italic: item.italic,
          underline: item.underline,
          color: item.color,
          widthFactor: item.widthFactor,
        },
        timeoutMilliseconds: 30000,
      });
    });
  });

  return commands;
}

function createFormattedCellWriteBatches(cellMatrix, batchSize = EXPORT_BATCH_SIZE) {
  const cells = collectFormattedCellWrites(cellMatrix);
  const batches = [];
  let currentCommands = [];
  let currentCellCount = 0;

  for (const cell of cells) {
    const cellCommands = createFormattedCellCommands(cell);
    if (currentCommands.length && currentCommands.length + cellCommands.length > batchSize) {
      batches.push({
        commands: currentCommands,
        cellCount: currentCellCount,
      });
      currentCommands = [];
      currentCellCount = 0;
    }
    currentCommands.push(...cellCommands);
    currentCellCount += 1;
  }

  if (currentCommands.length) {
    batches.push({
      commands: currentCommands,
      cellCount: currentCellCount,
    });
  }

  return batches;
}

function toApi5TextFlags(item) {
  let flags = 0;
  flags |= item?.italic ? API5_TEXT_FLAG_ITALIC_ON : API5_TEXT_FLAG_ITALIC_OFF;
  flags |= item?.bold ? API5_TEXT_FLAG_BOLD_ON : API5_TEXT_FLAG_BOLD_OFF;
  flags |= item?.underline ? API5_TEXT_FLAG_UNDERLINE_ON : API5_TEXT_FLAG_UNDERLINE_OFF;
  return flags;
}

function toApi5CellNumber(rowIndex, columnIndex, columnCount) {
  const cols = Number(columnCount);
  return (Number(rowIndex) * cols) + Number(columnIndex) + 1;
}

function buildInlineTempOutputPath({ fileName, tempPath }) {
  const directory = joinWindowsPath(tempPath || "", "kompas-pages\\inline");
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "");
  return ensureTblExtension(joinWindowsPath(
    directory,
    `${sanitizeFileStem(fileName || "table")}-inline-${stamp}`,
  ));
}

function getMatrixDimensions(matrix) {
  const rows = Array.isArray(matrix) ? matrix.length : 0;
  let cols = 0;
  for (const row of matrix || []) {
    cols = Math.max(cols, Array.isArray(row) ? row.length : 0);
  }
  return { rows, cols };
}

function reconcileLinkedLayout(layout, rows, cols, source = "table") {
  const safeRows = Math.max(1, Number(rows) || 1);
  const safeCols = Math.max(1, Number(cols) || 1);
  const next = {
    tableWidthMm: roundLayoutValue(layout?.tableWidthMm) || DEFAULT_LAYOUT.tableWidthMm,
    tableHeightMm: roundLayoutValue(layout?.tableHeightMm) || DEFAULT_LAYOUT.tableHeightMm,
    cellWidthMm: roundLayoutValue(layout?.cellWidthMm) || DEFAULT_LAYOUT.cellWidthMm,
    cellHeightMm: roundLayoutValue(layout?.cellHeightMm) || DEFAULT_LAYOUT.cellHeightMm,
  };

  if (source === "cell") {
    next.tableWidthMm = roundLayoutValue(next.cellWidthMm * safeCols);
    next.tableHeightMm = roundLayoutValue(next.cellHeightMm * safeRows);
    return next;
  }

  next.cellWidthMm = roundLayoutValue(next.tableWidthMm / safeCols);
  next.cellHeightMm = roundLayoutValue(next.tableHeightMm / safeRows);
  return next;
}

function buildAutoOutputPath({ documentPath, fileName, tempPath }) {
  const targetDirectory = documentPath
    ? dirname(documentPath)
    : joinWindowsPath(tempPath || "", "kompas-pages");
  if (!targetDirectory) {
    return "";
  }
  return ensureTblExtension(joinWindowsPath(targetDirectory, sanitizeFileStem(fileName)));
}

function collectCellWrites(matrix) {
  const writes = [];
  const { rows, cols } = getMatrixDimensions(matrix);
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      const value = String(row[columnIndex] ?? "");
      if (value === "") {
        continue;
      }
      writes.push({ rowIndex, columnIndex, value });
    }
  }
  return writes;
}

function chunkList(values, size) {
  const chunks = [];
  for (let index = 0; index < values.length; index += size) {
    chunks.push(values.slice(index, index + size));
  }
  return chunks;
}

function delay(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function createCellWriteBatches(matrix, batchSize = EXPORT_BATCH_SIZE) {
  return chunkList(collectCellWrites(matrix), batchSize).map((batch) => batch.map((cell) => ({
    commandId: "xlsx-to-kompas-tbl.table-cell-set-text",
    arguments: {
      rowIndex: cell.rowIndex,
      columnIndex: cell.columnIndex,
      value: cell.value,
    },
    timeoutMilliseconds: 30000,
  })));
}

function readWorkbookMatrix(file, bytes) {
  const workbook = window.XLSX.read(bytes, {
    type: "array",
    cellNF: true,
    cellText: true,
    cellDates: false,
    cellHTML: true,
    cellStyles: true,
    bookFiles: true,
  });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("XLSX workbook does not contain worksheets.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const sheetIndex = Math.max(0, workbook.SheetNames.indexOf(firstSheetName));
  const styleContext = buildWorkbookStyleContext(workbook);
  const worksheetXml = readWorkbookFileText(workbook, resolveWorksheetFilePath(workbook, sheetIndex));
  const worksheetMeta = parseWorksheetCellMeta(worksheetXml);
  const range = window.XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const matrix = [];
  const cellMatrix = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    const cellRow = [];
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = window.XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const sheetCell = sheet[address];
      const cellMeta = worksheetMeta[address] || {};
      const displayText = formatCellValue(sheetCell);
      const sharedStringIndex = Number.parseInt(cellMeta.rawValue, 10);
      const sharedStringEntry = Number.isInteger(sharedStringIndex)
        ? workbook.Strings?.[sharedStringIndex]
        : null;
      row.push(displayText);
      cellRow.push(createCellTransferPayload({
        address,
        rowIndex,
        columnIndex,
        text: displayText,
        styleIndex: cellMeta.styleIndex || 0,
        richTextXml: cellMeta.inlineRichXml,
        styleContext,
        sharedStringEntry,
      }));
    }
    matrix.push(row);
    cellMatrix.push(cellRow);
  }

  return {
    fileName: file.name,
    sheetName: firstSheetName,
    matrix,
    cellMatrix,
  };
}

function createActiveFrameChain(dsl, finalStep) {
  return [
    dsl.step("queryInterface", "IApplication"),
    dsl.step("get", "ActiveDocument"),
    dsl.step("queryInterface", "IKompasDocument2D"),
    dsl.step("get", "DocumentFrames"),
    dsl.step("index", "Item", {
      args: [dsl.literal(0, "int")],
    }),
    dsl.step("queryInterface", "IDocumentFrame"),
    finalStep,
  ];
}

function createZoomScaleCaptureArgs() {
  return [
    { Literal: 0, Converter: "double", ByRef: true, CaptureAs: "x" },
    { Literal: 0, Converter: "double", ByRef: true, CaptureAs: "y" },
    { Literal: 0, Converter: "double", ByRef: true, CaptureAs: "scale" },
  ];
}

function createXlsxToKompasTblModule() {
  return {
    id: "xlsx-to-kompas-tbl",
    title: "XLSX to KOMPAS TBL",
    subtitle: "XLSX парсится в браузере, а export и insert выполняются через runtime overlay WebBridge.Utility.",
    tabLabel: "XLSX/TBL",
    tabDetail: "active",

    getRuntimeContribution(context) {
      return {
        commands: {
          "xlsx-to-kompas-tbl.application.info": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.active-document": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.active-view": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.active-view-x": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
              context.dsl.step("queryInterface", "IView"),
              context.dsl.step("get", "X"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.active-view-y": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
              context.dsl.step("queryInterface", "IView"),
              context.dsl.step("get", "Y"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.active-view-update": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
              context.dsl.step("queryInterface", "IView"),
              context.dsl.step("call", "Update"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.active-frame-center-x": context.dsl.command(
            "kompas",
            "application",
            createActiveFrameChain(context.dsl, context.dsl.step("call", "GetZoomScale", {
              args: createZoomScaleCaptureArgs(),
            })),
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
              returnPath: "stored:x",
            },
          ),
          "xlsx-to-kompas-tbl.active-frame-center-y": context.dsl.command(
            "kompas",
            "application",
            createActiveFrameChain(context.dsl, context.dsl.step("call", "GetZoomScale", {
              args: createZoomScaleCaptureArgs(),
            })),
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
              returnPath: "stored:y",
            },
          ),
          "xlsx-to-kompas-tbl.active-frame-refresh": context.dsl.command(
            "kompas",
            "application",
            createActiveFrameChain(context.dsl, context.dsl.step("call", "RefreshWindow")),
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.view-table-count": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
              context.dsl.step("queryInterface", "ISymbols2DContainer"),
              context.dsl.step("get", "DrawingTables"),
              context.dsl.step("get", "Count"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.create-table": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
              context.dsl.step("queryInterface", "ISymbols2DContainer"),
              context.dsl.step("get", "DrawingTables"),
              context.dsl.step("call", "Add", {
                args: [
                  context.dsl.arg("rows", "int"),
                  context.dsl.arg("cols", "int"),
                  context.dsl.arg("cellHeightMm", "double"),
                  context.dsl.arg("cellWidthMm", "double"),
                  context.dsl.literal(0, "int"),
                ],
              }),
              context.dsl.step("queryInterface", "IDrawingTable"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.table-cell-set-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("set", "Str", {
                valueArgument: "value",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-clear-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("call", "Clear"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-set-one-line": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Range", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableRange"),
              context.dsl.step("get", "CellsFormat"),
              context.dsl.step("queryInterface", "ICellFormat"),
              context.dsl.step("set", "OneLine", {
                valueArgument: "oneLine",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-set-line-align": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("set", "Align", {
                valueArgument: "align",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-add-line": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("call", "Add"),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("set", "Align", {
                valueArgument: "align",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-add-item": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("call", "Add"),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("set", "Str", {
                valueArgument: "value",
              }),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("set", "FontName", {
                valueArgument: "fontName",
              }),
              context.dsl.step("set", "Height", {
                valueArgument: "heightMm",
              }),
              context.dsl.step("set", "Bold", {
                valueArgument: "bold",
              }),
              context.dsl.step("set", "Underline", {
                valueArgument: "underline",
              }),
              context.dsl.step("set", "Color", {
                valueArgument: "color",
              }),
              context.dsl.step("set", "WidthFactor", {
                valueArgument: "widthFactor",
              }),
              context.dsl.step("call", "set_Italic", {
                args: [context.dsl.arg("italic", "bool")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-add-item-before": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("call", "AddBefore", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("set", "Str", {
                valueArgument: "value",
              }),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("set", "FontName", {
                valueArgument: "fontName",
              }),
              context.dsl.step("set", "Height", {
                valueArgument: "heightMm",
              }),
              context.dsl.step("set", "Bold", {
                valueArgument: "bold",
              }),
              context.dsl.step("set", "Underline", {
                valueArgument: "underline",
              }),
              context.dsl.step("set", "Color", {
                valueArgument: "color",
              }),
              context.dsl.step("set", "WidthFactor", {
                valueArgument: "widthFactor",
              }),
              context.dsl.step("call", "set_Italic", {
                args: [context.dsl.arg("italic", "bool")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-set-item": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("set", "Str", {
                valueArgument: "value",
              }),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("set", "FontName", {
                valueArgument: "fontName",
              }),
              context.dsl.step("set", "Height", {
                valueArgument: "heightMm",
              }),
              context.dsl.step("set", "Bold", {
                valueArgument: "bold",
              }),
              context.dsl.step("set", "Underline", {
                valueArgument: "underline",
              }),
              context.dsl.step("set", "Color", {
                valueArgument: "color",
              }),
              context.dsl.step("set", "WidthFactor", {
                valueArgument: "widthFactor",
              }),
              context.dsl.step("call", "set_Italic", {
                args: [context.dsl.arg("italic", "bool")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("get", "Str"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-one-line": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Range", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableRange"),
              context.dsl.step("get", "CellsFormat"),
              context.dsl.step("queryInterface", "ICellFormat"),
              context.dsl.step("get", "OneLine"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-line-count": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("get", "Count"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-line-align": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("get", "Align"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-line-item-count": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("get", "Count"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("get", "Str"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-font-name": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("get", "FontName"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-height": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("get", "Height"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-bold": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("get", "Bold"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-italic": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("call", "get_Italic"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-underline": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("get", "Underline"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-color": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("get", "Color"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-get-item-width-factor": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "ITable"),
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("queryInterface", "ITableCell"),
              context.dsl.step("get", "Text"),
              context.dsl.step("queryInterface", "IText"),
              context.dsl.step("index", "TextLine", {
                args: [context.dsl.arg("lineIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextLine"),
              context.dsl.step("index", "TextItem", {
                args: [context.dsl.arg("itemIndex", "int")],
              }),
              context.dsl.step("queryInterface", "ITextItem"),
              context.dsl.step("queryInterface", "ITextFont"),
              context.dsl.step("get", "WidthFactor"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-save": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("call", "Save", {
                args: [context.dsl.arg("path", "path")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-update": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("call", "Update"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-delete": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("call", "Delete"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-get-temp": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("get", "Temp"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-get-valid": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("get", "Valid"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-get-x": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("get", "X"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-get-y": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("get", "Y"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-get-reference": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("get", "Reference"),
            ],
          ),
          "xlsx-to-kompas-tbl.table-set-position": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("queryInterface", "IDrawingTable"),
              context.dsl.step("set", "X", {
                valueArgument: "x",
              }),
              context.dsl.step("set", "Y", {
                valueArgument: "y",
              }),
              context.dsl.step("call", "Update"),
            ],
          ),
          "xlsx-to-kompas-tbl.insert-table": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "ActiveDocument"),
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
              context.dsl.step("queryInterface", "ISymbols2DContainer"),
              context.dsl.step("get", "DrawingTables"),
              context.dsl.step("call", "Load", {
                args: [context.dsl.arg("path", "path")],
              }),
              context.dsl.step("queryInterface", "IDrawingTable"),
            ],
            {
              defaultArguments: KOMPAS_API7_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.api5-active-document2d": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("call", "ActiveDocument2D"),
            ],
            {
              defaultArguments: KOMPAS_API5_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.api5-create-text-param": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("call", "GetParamStruct", {
                args: [context.dsl.literal(API5_STRUCT_TYPE_TEXT_PARAM, "short")],
              }),
            ],
            {
              defaultArguments: KOMPAS_API5_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.api5-create-text-line-param": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("call", "GetParamStruct", {
                args: [context.dsl.literal(API5_STRUCT_TYPE_TEXT_LINE_PARAM, "short")],
              }),
            ],
            {
              defaultArguments: KOMPAS_API5_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.api5-create-text-item-param": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("call", "GetParamStruct", {
                args: [context.dsl.literal(API5_STRUCT_TYPE_TEXT_ITEM_PARAM, "short")],
              }),
            ],
            {
              defaultArguments: KOMPAS_API5_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.api5-create-dynamic-array": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("call", "GetDynamicArray", {
                args: [context.dsl.arg("arrayType", "int")],
              }),
            ],
            {
              defaultArguments: KOMPAS_API5_ATTACHED_ARGUMENTS,
            },
          ),
          "xlsx-to-kompas-tbl.api5-object-init": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "Init"),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-param-get-line-array": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "GetTextLineArr"),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-param-set-line-array": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "SetTextLineArr", {
                args: [context.dsl.arg("lineArrayHandle", "handle")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-line-param-get-item-array": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "GetTextItemArr"),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-line-param-set-item-array": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "SetTextItemArr", {
                args: [context.dsl.arg("itemArrayHandle", "handle")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-item-param-get-font": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "GetItemFont"),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-item-param-set-basic": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("set", "s", {
                valueArgument: "value",
              }),
              context.dsl.step("set", "type", {
                valueArgument: "itemType",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-item-param-set-font": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "SetItemFont", {
                args: [context.dsl.arg("fontHandle", "handle")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-text-item-font-set": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("set", "fontName", {
                valueArgument: "fontName",
              }),
              context.dsl.step("set", "height", {
                valueArgument: "heightMm",
              }),
              context.dsl.step("set", "color", {
                valueArgument: "color",
              }),
              context.dsl.step("set", "bitVector", {
                valueArgument: "bitVector",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-dynamic-array-count": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "ksGetArrayCount"),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-dynamic-array-add-item": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "ksAddArrayItem", {
                args: [
                  context.dsl.arg("index", "int"),
                  context.dsl.arg("itemHandle", "handle"),
                ],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-document-open-table": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "ksOpenTable", {
                args: [context.dsl.arg("tableReference", "int")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-document-end-obj": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "ksEndObj"),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-document-clear-table-cell-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "ksClearTableColumnText", {
                args: [context.dsl.arg("cellNumber", "int")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.api5-document-set-table-cell-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "ksSetTableColumnText", {
                args: [
                  context.dsl.arg("cellNumber", "int"),
                  context.dsl.arg("textHandle", "handle"),
                ],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.open-document": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("queryInterface", "IApplication"),
              context.dsl.step("get", "Documents"),
              context.dsl.step("call", "Open", {
                args: [
                  context.dsl.arg("path", "path"),
                  context.dsl.literal(false, "bool"),
                  context.dsl.literal(true, "bool"),
                ],
              }),
            ],
            {
              defaultArguments: KOMPAS_API7_OPEN_ARGUMENTS,
            },
          ),
        },
        allowedTypes: [],
        comAdapters: [KOMPAS_COM_ADAPTER],
      };
    },

    mount(container, context) {
      const savedLayout = context.storage.get("layout", DEFAULT_LAYOUT) || DEFAULT_LAYOUT;
      const initialLayoutDriver = savedLayout.layoutDriver === "cell" ? "cell" : "table";
      const state = {
        fileName: "",
        sheetName: "",
        matrix: [],
        cellMatrix: [],
        status: null,
        tempPath: "",
        tempPathLoaded: false,
        lastExport: null,
        lastBytes: null,
        autoFollowOutput: true,
        layoutDriver: initialLayoutDriver,
        layout: reconcileLinkedLayout(savedLayout, 1, 1, initialLayoutDriver),
        pollTimer: null,
        pollBusy: false,
        statusPollingSuspended: false,
        active: false,
        dragDepth: 0,
      };

      container.innerHTML = `
        <div class="module-grid">
          <div class="stack">
            <section class="status-box">
              <div class="status-row"><span>File</span><strong id="xlsx-file-name">not loaded</strong></div>
              <div class="status-row"><span>Sheet</span><strong id="xlsx-sheet-name">-</strong></div>
              <div class="status-row"><span>Matrix</span><strong id="xlsx-matrix-size">0 x 0</strong></div>
              <div class="status-row"><span>Document</span><strong id="xlsx-document-name">no active doc</strong></div>
              <div class="status-row"><span>View</span><strong id="xlsx-view-name">-</strong></div>
            </section>

            <section class="panel panel--inner">
              <div class="panel__head">
                <div>
                  <h2>Source</h2>
                  <p class="panel__subtitle">Автостатус KOMPAS обновляется каждые 2 секунды при активной вкладке.</p>
                </div>
              </div>

              <div class="dropzone" id="xlsx-dropzone">
                <div class="dropzone__copy">
                  <strong>Drop XLSX</strong>
                  <span>Перетащите .xlsx сюда или откройте файл вручную.</span>
                </div>
                <label class="button button--ghost" for="xlsx-file-input">Open XLSX</label>
              </div>

              <input id="xlsx-file-input" type="file" accept=".xlsx,.xlsm,.xlsb,.xls" hidden>

              <div class="field-grid">
                <label class="field">
                  <span>table width, mm</span>
                  <input id="xlsx-table-width" type="number" step="0.1" min="0" value="${state.layout.tableWidthMm}">
                </label>
                <label class="field">
                  <span>table height, mm</span>
                  <input id="xlsx-table-height" type="number" step="0.1" min="0" value="${state.layout.tableHeightMm}">
                </label>
                <label class="field">
                  <span>cell width, mm</span>
                  <input id="xlsx-cell-width" type="number" step="0.1" min="0" value="${state.layout.cellWidthMm}">
                </label>
                <label class="field">
                  <span>cell height, mm</span>
                  <input id="xlsx-cell-height" type="number" step="0.1" min="0" value="${state.layout.cellHeightMm}">
                </label>
              </div>

              <label class="field">
                <span>output path (.tbl)</span>
                <div class="field-inline">
                  <input id="xlsx-output-path" type="text" spellcheck="false" placeholder="%TEMP%\\kompas-pages\\table.tbl">
                  <button class="button button--ghost" type="button" id="xlsx-follow-button">Follow</button>
                </div>
                <small class="field__hint" id="xlsx-output-mode">auto-follow: waiting for XLSX</small>
              </label>

              <div class="summary-box" id="xlsx-layout-summary">Размеры будут рассчитаны после загрузки файла.</div>
              <div class="result-box" id="xlsx-result-box">Ожидание данных.</div>

              <div class="action-row">
                <button class="button" type="button" id="xlsx-export-button">Export</button>
                <button class="button button--ghost" type="button" id="xlsx-inline-button" disabled>Inline</button>
                <button class="button button--ghost" type="button" id="xlsx-insert-button" disabled>Insert</button>
                <button class="button button--ghost" type="button" id="xlsx-download-button" disabled>Download</button>
                <button class="button button--ghost" type="button" id="xlsx-reset-button">Reset</button>
              </div>
            </section>
          </div>

          <section class="panel panel--inner">
            <div class="panel__head">
              <div>
                <h2>Preview</h2>
                <p class="panel__subtitle" id="xlsx-preview-meta">UsedRange первого листа.</p>
              </div>
            </div>
            <div class="preview-wrap">
              <table class="preview-table" id="xlsx-preview-table">
                <tbody>
                  <tr><td class="preview-table__empty">Загрузите XLSX.</td></tr>
                </tbody>
              </table>
            </div>
          </section>
        </div>
      `;

      const refs = {
        dropzone: container.querySelector("#xlsx-dropzone"),
        fileInput: container.querySelector("#xlsx-file-input"),
        fileName: container.querySelector("#xlsx-file-name"),
        sheetName: container.querySelector("#xlsx-sheet-name"),
        matrixSize: container.querySelector("#xlsx-matrix-size"),
        documentName: container.querySelector("#xlsx-document-name"),
        viewName: container.querySelector("#xlsx-view-name"),
        tableWidth: container.querySelector("#xlsx-table-width"),
        tableHeight: container.querySelector("#xlsx-table-height"),
        cellWidth: container.querySelector("#xlsx-cell-width"),
        cellHeight: container.querySelector("#xlsx-cell-height"),
        outputPath: container.querySelector("#xlsx-output-path"),
        outputMode: container.querySelector("#xlsx-output-mode"),
        followButton: container.querySelector("#xlsx-follow-button"),
        layoutSummary: container.querySelector("#xlsx-layout-summary"),
        resultBox: container.querySelector("#xlsx-result-box"),
        exportButton: container.querySelector("#xlsx-export-button"),
        inlineButton: container.querySelector("#xlsx-inline-button"),
        insertButton: container.querySelector("#xlsx-insert-button"),
        downloadButton: container.querySelector("#xlsx-download-button"),
        resetButton: container.querySelector("#xlsx-reset-button"),
        previewMeta: container.querySelector("#xlsx-preview-meta"),
        previewTable: container.querySelector("#xlsx-preview-table"),
      };

      function persistLayout() {
        context.storage.set("layout", {
          ...state.layout,
          layoutDriver: state.layoutDriver,
        });
      }

      function currentDimensions() {
        return getMatrixDimensions(state.matrix);
      }

      function syncLayoutInputs() {
        refs.tableWidth.value = String(state.layout.tableWidthMm || 0);
        refs.tableHeight.value = String(state.layout.tableHeightMm || 0);
        refs.cellWidth.value = String(state.layout.cellWidthMm || 0);
        refs.cellHeight.value = String(state.layout.cellHeightMm || 0);
      }

      function renderPreview() {
        const body = document.createElement("tbody");
        if (!state.matrix.length) {
          const row = document.createElement("tr");
          const cell = document.createElement("td");
          cell.className = "preview-table__empty";
          cell.textContent = "Загрузите XLSX.";
          row.append(cell);
          body.append(row);
          refs.previewTable.replaceChildren(body);
          return;
        }

        for (const rowValues of state.matrix) {
          const row = document.createElement("tr");
          for (const value of rowValues) {
            const cell = document.createElement("td");
            cell.textContent = value;
            row.append(cell);
          }
          body.append(row);
        }
        refs.previewTable.replaceChildren(body);
      }

      function suggestOutputPath() {
        if (!state.autoFollowOutput || !state.fileName) {
          return;
        }
        const suggested = buildAutoOutputPath({
          documentPath: state.status?.documentPath || "",
          fileName: state.fileName,
          tempPath: state.tempPath,
        });
        if (suggested) {
          refs.outputPath.value = suggested;
        }
      }

      function renderSummary() {
        const { rows, cols } = currentDimensions();
        if (!rows || !cols) {
          refs.layoutSummary.textContent = "Размеры будут рассчитаны после загрузки файла.";
          return;
        }
        const writeCount = collectFormattedCellWrites(state.cellMatrix).length;
        refs.layoutSummary.textContent = [
          `cell=${formatMm(state.layout.cellWidthMm)} x ${formatMm(state.layout.cellHeightMm)} mm`,
          `table=${formatMm(state.layout.tableWidthMm)} x ${formatMm(state.layout.tableHeightMm)} mm`,
          `writes=${writeCount}/${rows * cols}`,
        ].join(" | ");
      }

      function renderStatus() {
        refs.fileName.textContent = state.fileName || "not loaded";
        refs.sheetName.textContent = state.sheetName || "-";
        refs.matrixSize.textContent = formatSize(currentDimensions().rows, currentDimensions().cols);
        refs.documentName.textContent = state.status?.documentPath || state.status?.documentName || "no active doc";
        refs.viewName.textContent = state.status?.viewName || "-";
        refs.previewMeta.textContent = state.sheetName
          ? `${state.sheetName} | ${formatSize(currentDimensions().rows, currentDimensions().cols)}`
          : "UsedRange первого листа.";
      }

      function renderOutputMode() {
        refs.followButton.classList.toggle("is-active", state.autoFollowOutput);
        if (state.autoFollowOutput) {
          refs.outputMode.textContent = state.fileName
            ? "auto-follow: path for Export/Insert; Inline materializes the same TBL in %TEMP%"
            : "auto-follow: waiting for XLSX";
          return;
        }
        refs.outputMode.textContent = "manual path: used by Export/Insert until Reset or Follow";
      }

      function updateModuleIndicator() {
        if (!context.getBridgeState().runtimeReady) {
          context.setModuleBadge("doc idle", false);
          context.setModuleMeta("Bridge подключён не полностью или runtime ещё не загружен.");
          return;
        }
        if (state.status?.connected && state.status?.viewHandleId) {
          context.setModuleBadge("doc ok", true);
          context.setModuleMeta(`${state.status.documentPath || state.status.documentName} | view=${state.status.viewName || "-"}`);
          return;
        }
        if (state.status?.connected) {
          context.setModuleBadge("doc none", false);
          context.setModuleMeta("KOMPAS запущен, но активный 2D документ или вид недоступен.");
          return;
        }
        context.setModuleBadge("kompas off", false);
        context.setModuleMeta(state.status?.errorMessage || "KOMPAS не найден.");
      }

      function updateActionState() {
        const bridgeState = context.getBridgeState();
        const hasMatrix = currentDimensions().rows > 0 && currentDimensions().cols > 0;
        const hasExportPath = Boolean((state.lastExport?.outputPath || refs.outputPath.value || "").trim());
        const hasBytes = Boolean(state.lastBytes && state.lastBytes.length > 0);
        const hasActiveView = Boolean(state.status?.viewHandleId);

        refs.exportButton.disabled = !bridgeState.runtimeReady || !hasMatrix || !hasActiveView;
        refs.inlineButton.disabled = !bridgeState.runtimeReady || !hasMatrix || !hasActiveView;
        refs.insertButton.disabled = !bridgeState.runtimeReady || !hasExportPath || !hasActiveView;
        refs.downloadButton.disabled = !hasBytes;
      }

      function renderAll() {
        syncLayoutInputs();
        suggestOutputPath();
        renderSummary();
        renderStatus();
        renderOutputMode();
        renderPreview();
        updateModuleIndicator();
        updateActionState();
      }

      async function ensureTempPathLoaded(force = false) {
        if (!context.getBridgeState().runtimeReady) {
          return "";
        }
        if (state.tempPathLoaded && !force) {
          return state.tempPath;
        }
        try {
          state.tempPath = normalizeWindowsPath(await context.getTempPath());
          state.tempPathLoaded = true;
        } catch {
          state.tempPath = "";
        }
        return state.tempPath;
      }

      async function refreshStatus(options = {}) {
        const quiet = options.quiet === true;
        const commandArguments = options.forceRefreshApplication === true
          ? { refresh: true }
          : {};
        if (!context.getBridgeState().runtimeReady) {
          state.status = null;
          renderAll();
          return state.status;
        }

        await ensureTempPathLoaded();

        try {
          const applicationExecution = await context.executeCommand(
            "xlsx-to-kompas-tbl.application.info",
            commandArguments,
            15000,
          );
          const application = applicationExecution.result || {};
          const documentExecution = await context.executeCommand(
            "xlsx-to-kompas-tbl.active-document",
            commandArguments,
            15000,
          );
          const documentResult = documentExecution.result || null;
          let viewResult = null;
          if (documentResult?.handleId) {
            const viewExecution = await context.executeCommand(
              "xlsx-to-kompas-tbl.active-view",
              commandArguments,
              15000,
            );
            viewResult = viewExecution.result || null;
          }

          state.status = {
            connected: Boolean(application.connected ?? true),
            applicationProgId: String(application.progId || ""),
            documentName: String(documentResult?.name || ""),
            documentPath: String(documentResult?.path || ""),
            documentHandleId: String(documentResult?.handleId || ""),
            viewName: String(viewResult?.name || ""),
            viewHandleId: String(viewResult?.handleId || ""),
            hasActiveDocument: Boolean(documentResult?.handleId),
            errorMessage: "",
          };
        } catch (error) {
          state.status = {
            connected: false,
            applicationProgId: "",
            documentName: "",
            documentPath: "",
            documentHandleId: "",
            viewName: "",
            viewHandleId: "",
            hasActiveDocument: false,
            errorMessage: String(error.message || error),
          };
          if (!quiet) {
            context.logger.error("status", state.status.errorMessage);
          }
        }

        renderAll();
        return state.status;
      }

      function syncStatusPollingSuspended(suspended) {
        state.statusPollingSuspended = Boolean(suspended);
        syncPolling();
      }

      async function runWithStatusPollingSuspended(action) {
        const previous = state.statusPollingSuspended;
        syncStatusPollingSuspended(true);
        try {
          return await action();
        } finally {
          syncStatusPollingSuspended(previous);
        }
      }

      function syncPolling() {
        const shouldPoll = state.active
          && context.getBridgeState().runtimeReady
          && document.visibilityState === "visible"
          && !state.statusPollingSuspended;
        if (!shouldPoll) {
          if (state.pollTimer !== null) {
            window.clearInterval(state.pollTimer);
            state.pollTimer = null;
          }
          return;
        }
        if (state.pollTimer !== null) {
          return;
        }

        state.pollTimer = window.setInterval(() => {
          if (state.pollBusy) {
            return;
          }
          state.pollBusy = true;
          refreshStatus({ quiet: true })
            .catch(() => {})
            .finally(() => {
              state.pollBusy = false;
            });
        }, 2000);
      }

      function applyLayoutChange(driver, patch) {
        const { rows, cols } = currentDimensions();
        state.layoutDriver = driver;
        state.layout = reconcileLinkedLayout(
          { ...state.layout, ...patch },
          rows || 1,
          cols || 1,
          driver,
        );
        persistLayout();
        renderAll();
      }

      async function resolveOutputPath() {
        await ensureTempPathLoaded();
        const manualValue = normalizeWindowsPath(refs.outputPath.value.trim());
        const effective = manualValue
          ? ensureTblExtension(manualValue)
          : buildAutoOutputPath({
            documentPath: state.status?.documentPath || "",
            fileName: state.fileName || "table.xlsx",
            tempPath: state.tempPath,
          });
        if (!effective) {
          throw new Error("Не удалось определить output path.");
        }
        refs.outputPath.value = effective;
        return effective;
      }

      async function ensureActiveViewStatus() {
        let lastStatus = null;
        for (let attempt = 0; attempt < 8; attempt += 1) {
          lastStatus = await refreshStatus({ quiet: true });
          if (lastStatus?.viewHandleId) {
            return lastStatus;
          }
          await delay(500);
        }
        throw new Error(lastStatus?.errorMessage || "Активный 2D документ KOMPAS или его view не найден.");
      }

      async function readViewTableCount() {
        const execution = await context.executeCommand(
          "xlsx-to-kompas-tbl.view-table-count",
          { refresh: true },
          15000,
        );
        return Number(execution.result) || 0;
      }

      async function readActiveViewPoint() {
        const [xExecution, yExecution] = await Promise.all([
          context.executeCommand("xlsx-to-kompas-tbl.active-view-x", {}, 15000),
          context.executeCommand("xlsx-to-kompas-tbl.active-view-y", {}, 15000),
        ]);
        return {
          x: Number(xExecution.result) || 0,
          y: Number(yExecution.result) || 0,
        };
      }

      async function readActiveFrameCenterPoint() {
        const [xExecution, yExecution] = await Promise.all([
          context.executeCommand("xlsx-to-kompas-tbl.active-frame-center-x", {}, 15000),
          context.executeCommand("xlsx-to-kompas-tbl.active-frame-center-y", {}, 15000),
        ]);
        const x = Number(xExecution.result);
        const y = Number(yExecution.result);
        if (!Number.isFinite(x) || !Number.isFinite(y)) {
          throw new Error(`Invalid active frame center: x=${xExecution.result} y=${yExecution.result}`);
        }
        return { x, y };
      }

      async function updateActiveView() {
        const execution = await context.executeCommand(
          "xlsx-to-kompas-tbl.active-view-update",
          { refresh: true },
          30000,
        );
        if (execution.result === false) {
          throw new Error("KOMPAS returned Update=false for active view.");
        }
      }

      async function refreshActiveFrame() {
        await context.executeCommand(
          "xlsx-to-kompas-tbl.active-frame-refresh",
          { refresh: true },
          30000,
        );
      }

      function refreshActiveViewSoon(reason) {
        window.setTimeout(() => {
          updateActiveView().catch((error) => {
            context.logger.error(`view-update:${reason}`, String(error.message || error));
          });
        }, 0);
      }

      function refreshActiveFrameSoon(reason) {
        window.setTimeout(() => {
          refreshActiveFrame().catch((error) => {
            context.logger.error(`frame-refresh:${reason}`, String(error.message || error));
          });
        }, 0);
      }

      function refreshActiveDisplaySoon(reason) {
        refreshActiveFrameSoon(reason);
        refreshActiveViewSoon(reason);
      }

      async function readTableDebug(tableHandleId) {
        const [tempExecution, validExecution, xExecution, yExecution] = await Promise.all([
          context.executeCommand("xlsx-to-kompas-tbl.table-get-temp", { handleId: tableHandleId }, 15000),
          context.executeCommand("xlsx-to-kompas-tbl.table-get-valid", { handleId: tableHandleId }, 15000),
          context.executeCommand("xlsx-to-kompas-tbl.table-get-x", { handleId: tableHandleId }, 15000),
          context.executeCommand("xlsx-to-kompas-tbl.table-get-y", { handleId: tableHandleId }, 15000),
        ]);
        return {
          temp: Boolean(tempExecution.result),
          valid: Boolean(validExecution.result),
          x: Number(xExecution.result) || 0,
          y: Number(yExecution.result) || 0,
        };
      }

      async function createTableHandle(rows, cols) {
        const createExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.create-table",
          {
            rows,
            cols,
            cellWidthMm: state.layout.cellWidthMm,
            cellHeightMm: state.layout.cellHeightMm,
          },
          60000,
        );
        const tableHandleId = String(createExecution.result?.handleId || "");
        if (!tableHandleId) {
          throw new Error("KOMPAS did not return a drawing table handle.");
        }
        return tableHandleId;
      }

      function requireHandleId(execution, label) {
        const handleId = String(execution?.result?.handleId || "");
        if (!handleId) {
          throw new Error(`${label} did not return a handle.`);
        }
        return handleId;
      }

      function expectApi5Success(result, label) {
        if (result !== false && result !== null && result !== undefined) {
          return;
        }
        throw new Error(`${label} failed: ${result}`);
      }

      async function createApi5ParamHandle(commandId, label) {
        const execution = await context.executeCommand(commandId, {}, 15000);
        return requireHandleId(execution, label);
      }

      async function createApi5DynamicArrayHandle(arrayType, label) {
        const execution = await context.executeCommand(
          "xlsx-to-kompas-tbl.api5-create-dynamic-array",
          { arrayType },
          15000,
        );
        return requireHandleId(execution, label);
      }

      async function createApi5TextParamHandle(cell) {
        const textParamHandle = await createApi5ParamHandle(
          "xlsx-to-kompas-tbl.api5-create-text-param",
          "api5-text-param",
        );
        await context.executeCommand("xlsx-to-kompas-tbl.api5-object-init", { handleId: textParamHandle }, 15000);
        const lineArrayHandle = await createApi5DynamicArrayHandle(
          API5_TEXT_LINE_ARRAY_TYPE,
          "api5-text-line-array",
        );

        for (const line of cell.lines) {
          const lineHandle = await createApi5ParamHandle(
            "xlsx-to-kompas-tbl.api5-create-text-line-param",
            "api5-text-line-param",
          );
          await context.executeCommand("xlsx-to-kompas-tbl.api5-object-init", { handleId: lineHandle }, 15000);
          const itemArrayHandle = await createApi5DynamicArrayHandle(
            API5_TEXT_ITEM_ARRAY_TYPE,
            "api5-text-item-array",
          );

          for (const item of line.items) {
            const itemHandle = await createApi5ParamHandle(
              "xlsx-to-kompas-tbl.api5-create-text-item-param",
              "api5-text-item-param",
            );
            await context.executeCommand("xlsx-to-kompas-tbl.api5-object-init", { handleId: itemHandle }, 15000);
            await context.executeCommand(
              "xlsx-to-kompas-tbl.api5-text-item-param-set-basic",
              {
                handleId: itemHandle,
                value: item.text,
                itemType: API5_TEXT_ITEM_STRING,
              },
              15000,
            );

            const fontExecution = await context.executeCommand(
              "xlsx-to-kompas-tbl.api5-text-item-param-get-font",
              { handleId: itemHandle },
              15000,
            );
            const fontHandle = requireHandleId(fontExecution, "api5-text-item-font");
            await context.executeCommand(
              "xlsx-to-kompas-tbl.api5-text-item-font-set",
              {
                handleId: fontHandle,
                fontName: item.fontName,
                heightMm: item.heightMm,
                color: item.color,
                bitVector: toApi5TextFlags(item),
              },
              15000,
            );
            expectApi5Success(
              (await context.executeCommand(
                "xlsx-to-kompas-tbl.api5-text-item-param-set-font",
                {
                  handleId: itemHandle,
                  fontHandle,
                },
                15000,
              )).result,
              "api5-text-item-param-set-font",
            );

            expectApi5Success(
              (await context.executeCommand(
                "xlsx-to-kompas-tbl.api5-dynamic-array-add-item",
                {
                  handleId: itemArrayHandle,
                  index: -1,
                  itemHandle,
                },
                15000,
              )).result,
              "api5-item-array-append",
            );
          }

          expectApi5Success(
            (await context.executeCommand(
              "xlsx-to-kompas-tbl.api5-text-line-param-set-item-array",
              {
                handleId: lineHandle,
                itemArrayHandle,
              },
              15000,
            )).result,
            "api5-text-line-param-set-item-array",
          );

          expectApi5Success(
            (await context.executeCommand(
              "xlsx-to-kompas-tbl.api5-dynamic-array-add-item",
              {
                handleId: lineArrayHandle,
                index: -1,
                itemHandle: lineHandle,
              },
              15000,
            )).result,
            "api5-line-array-append",
          );
        }

        expectApi5Success(
          (await context.executeCommand(
            "xlsx-to-kompas-tbl.api5-text-param-set-line-array",
            {
              handleId: textParamHandle,
              lineArrayHandle,
            },
            15000,
          )).result,
          "api5-text-param-set-line-array",
        );
        return textParamHandle;
      }

      async function populateTableHandle(tableHandleId, progressLabel) {
        const writes = collectFormattedCellWrites(state.cellMatrix);
        if (writes.length === 0) {
          refs.resultBox.textContent = `${progressLabel}: 0/0 cells`;
          return;
        }

        const { cols } = getMatrixDimensions(state.cellMatrix);
        const api5DocumentHandle = requireHandleId(
          await context.executeCommand("xlsx-to-kompas-tbl.api5-active-document2d", {}, 15000),
          "api5-active-document2d",
        );
        const prepareReferenceExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-update",
          { handleId: tableHandleId },
          30000,
        );
        if (prepareReferenceExecution.result === false) {
          throw new Error("KOMPAS returned Update=false before reading table reference.");
        }
        const tableReferenceExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-get-reference",
          { handleId: tableHandleId },
          15000,
        );
        const tableReference = coerceNumberLike(tableReferenceExecution.result);
        if (!Number.isFinite(tableReference) || tableReference <= 0) {
          context.logger.info("api5-table-reference", JSON.stringify(tableReferenceExecution.result));
          throw new Error(`Invalid table reference: ${tableReferenceExecution.result}`);
        }

        expectApi5Success(
          (await context.executeCommand(
            "xlsx-to-kompas-tbl.api5-document-open-table",
            {
              handleId: api5DocumentHandle,
              tableReference,
            },
            30000,
          )).result,
          "api5-document-open-table",
        );

        let completedWrites = 0;
        try {
          for (const cell of writes) {
            const textParamHandle = await createApi5TextParamHandle(cell);
            const cellNumber = toApi5CellNumber(cell.rowIndex, cell.columnIndex, cols);
            expectApi5Success(
              (await context.executeCommand(
                "xlsx-to-kompas-tbl.api5-document-set-table-cell-text",
                {
                  handleId: api5DocumentHandle,
                  cellNumber,
                  textHandle: textParamHandle,
                },
                30000,
              )).result,
              "api5-document-set-table-cell-text",
            );
            completedWrites += 1;
            refs.resultBox.textContent = `${progressLabel}: ${completedWrites}/${writes.length} cells`;
          }
        } finally {
          expectApi5Success(
            (await context.executeCommand(
              "xlsx-to-kompas-tbl.api5-document-end-obj",
              {
                handleId: api5DocumentHandle,
              },
              30000,
            )).result,
            "api5-document-end-obj",
          );
        }

        for (const cell of writes) {
          await context.executeCommand(
            "xlsx-to-kompas-tbl.table-cell-set-one-line",
            {
              handleId: tableHandleId,
              rowIndex: cell.rowIndex,
              columnIndex: cell.columnIndex,
              oneLine: Boolean(cell.oneLine),
            },
            15000,
          );
          for (let lineIndex = 0; lineIndex < cell.lines.length; lineIndex += 1) {
            await context.executeCommand(
              "xlsx-to-kompas-tbl.table-cell-set-line-align",
              {
                handleId: tableHandleId,
                rowIndex: cell.rowIndex,
                columnIndex: cell.columnIndex,
                lineIndex,
                align: cell.alignCode,
              },
              15000,
            );
          }
        }
      }

      async function materializeTableFileViaRuntime(outputPath, rows, cols, progressLabel) {
        const tableHandleId = await createTableHandle(rows, cols);
        await populateTableHandle(tableHandleId, progressLabel);
        const updateExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-update",
          {
            handleId: tableHandleId,
          },
          30000,
        );
        if (updateExecution.result === false) {
          throw new Error("KOMPAS returned Update=false before Save.");
        }

        const saveExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-save",
          {
            handleId: tableHandleId,
            path: outputPath,
          },
          60000,
        );
        if (saveExecution.result === false) {
          throw new Error("KOMPAS returned Save=false.");
        }
        const deleteExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-delete",
          {
            handleId: tableHandleId,
          },
          30000,
        );
        if (deleteExecution.result === false) {
          throw new Error("KOMPAS returned Delete=false after Save.");
        }

        return {
          tableHandleId,
          outputPath,
          rows,
          cols,
        };
      }

      async function exportTableViaRuntime(outputPath, rows, cols) {
        return materializeTableFileViaRuntime(outputPath, rows, cols, "Export");
      }

      async function inlineTableViaRuntime(rows, cols) {
        let anchorPoint;
        let anchorSource = "frame-center";
        let anchorFallbackReason = "";
        try {
          anchorPoint = await readActiveFrameCenterPoint();
        } catch (error) {
          anchorPoint = await readActiveViewPoint();
          anchorSource = "view-fallback";
          anchorFallbackReason = String(error.message || error);
          context.logger.info("inline-anchor", anchorFallbackReason);
        }
        const targetX = anchorSource === "frame-center" ? anchorPoint.x : anchorPoint.x + 20;
        const targetY = anchorSource === "frame-center" ? anchorPoint.y : anchorPoint.y + 20;
        const inlinePath = buildInlineTempOutputPath({
          fileName: state.fileName || "table.xlsx",
          tempPath: state.tempPath,
        });
        const inlineDirectory = dirname(inlinePath);
        if (!inlineDirectory) {
          throw new Error("Inline temp directory is empty.");
        }
        await context.ensureDirectory(inlineDirectory);
        if (await context.fileExists(inlinePath)) {
          await context.deleteFile(inlinePath);
        }
        await materializeTableFileViaRuntime(inlinePath, rows, cols, "Inline");
        const insertResult = await insertTableViaRuntime(inlinePath);
        const tableHandleId = insertResult.tableHandleId;
        if (!tableHandleId) {
          throw new Error("KOMPAS did not return a drawing table handle after inline insert.");
        }
        const positionExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-set-position",
          {
            handleId: tableHandleId,
            x: targetX,
            y: targetY,
          },
          30000,
        );
        if (positionExecution.result === false) {
          throw new Error("KOMPAS returned Update=false while positioning Inline table.");
        }
        const debugAfterUpdate = await readTableDebug(tableHandleId);

        return {
          outputPath: inlinePath,
          tableHandleId,
          tableCountBefore: insertResult.tableCountBefore,
          tableCountAfter: insertResult.tableCountAfter,
          anchorPoint,
          anchorSource,
          anchorFallbackReason,
          targetPoint: { x: targetX, y: targetY },
          debugAfterUpdate,
          rows,
          cols,
        };
      }

      async function exportTable() {
        const { rows, cols } = currentDimensions();
        if (!rows || !cols) {
          throw new Error("XLSX matrix is empty.");
        }

        await ensureActiveViewStatus();
        const outputPath = await resolveOutputPath();
        const outputDirectory = dirname(outputPath);
        if (!outputDirectory) {
          throw new Error("Output directory is empty.");
        }

        refs.resultBox.textContent = "Экспорт выполняется...";
        refs.exportButton.disabled = true;

        try {
          await runWithStatusPollingSuspended(async () => {
            await context.ensureDirectory(outputDirectory);
            if (await context.fileExists(outputPath)) {
              await context.deleteFile(outputPath);
            }
            const exportResult = await exportTableViaRuntime(outputPath, rows, cols);

            state.lastBytes = await context.readFileBytes(exportResult.outputPath);
            state.lastExport = {
              outputPath: exportResult.outputPath,
              fileSize: state.lastBytes.length,
              rows: exportResult.rows,
              cols: exportResult.cols,
            };
            refs.resultBox.textContent = `OK | ${exportResult.outputPath} | ${state.lastBytes.length} bytes`;
            context.logger.info("export", exportResult.outputPath);
            await refreshStatus({ quiet: true, forceRefreshApplication: true });
          });
        } finally {
          refs.exportButton.disabled = false;
          updateActionState();
        }
      }

      async function inlineTable() {
        const { rows, cols } = currentDimensions();
        if (!rows || !cols) {
          throw new Error("XLSX matrix is empty.");
        }

        await ensureActiveViewStatus();
        refs.resultBox.textContent = "Inline is running...";
        refs.inlineButton.disabled = true;

        try {
          await runWithStatusPollingSuspended(async () => {
            const result = await inlineTableViaRuntime(rows, cols);
            refs.resultBox.textContent = `Inline | ${result.outputPath} | anchor=${result.anchorSource} (${result.anchorPoint.x},${result.anchorPoint.y}) target=(${result.targetPoint.x},${result.targetPoint.y}) | temp=${result.debugAfterUpdate.temp} valid=${result.debugAfterUpdate.valid} x=${result.debugAfterUpdate.x} y=${result.debugAfterUpdate.y} | ${result.tableCountBefore} -> ${result.tableCountAfter}`;
            context.logger.info("inline", JSON.stringify({
              outputPath: result.outputPath,
              before: result.tableCountBefore,
              after: result.tableCountAfter,
              temp: result.debugAfterUpdate.temp,
              valid: result.debugAfterUpdate.valid,
              x: result.debugAfterUpdate.x,
              y: result.debugAfterUpdate.y,
              anchorX: result.anchorPoint.x,
              anchorY: result.anchorPoint.y,
              anchorSource: result.anchorSource,
              anchorFallbackReason: result.anchorFallbackReason,
              targetX: result.targetPoint.x,
              targetY: result.targetPoint.y,
            }));
            refreshActiveDisplaySoon("inline");
            await refreshStatus({ quiet: true, forceRefreshApplication: true });
          });
        } finally {
          refs.inlineButton.disabled = false;
          updateActionState();
        }
      }

      async function insertTableViaRuntime(tblPath) {
        const tableCountBefore = await readViewTableCount();
        const insertExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.insert-table",
          {
            path: tblPath,
          },
          60000,
        );
        const tableHandleId = String(insertExecution.result?.handleId || "");
        const debugAfterInsert = tableHandleId
          ? await readTableDebug(tableHandleId)
          : null;
        const tableCountAfter = await readViewTableCount();
        return {
          tableHandleId,
          tableCountBefore,
          tableCountAfter,
          debugAfterInsert,
        };
      }

      async function insertTable() {
        await ensureActiveViewStatus();
        const tblPath = await resolveOutputPath();
        if (!await context.fileExists(tblPath)) {
          throw new Error(`Table file was not found: ${tblPath}`);
        }

        refs.resultBox.textContent = "Вставка выполняется...";
        await runWithStatusPollingSuspended(async () => {
          const result = await insertTableViaRuntime(tblPath);

          refs.resultBox.textContent = result.debugAfterInsert
            ? `Inserted | ${result.tableCountBefore} -> ${result.tableCountAfter} | temp=${result.debugAfterInsert.temp} valid=${result.debugAfterInsert.valid} x=${result.debugAfterInsert.x} y=${result.debugAfterInsert.y}`
            : `Inserted | ${result.tableCountBefore} -> ${result.tableCountAfter}`;
          context.logger.info("insert", tblPath);
          refreshActiveDisplaySoon("insert");
          await refreshStatus({ quiet: true, forceRefreshApplication: true });
        });
      }

      async function downloadTable() {
        if ((!state.lastBytes || state.lastBytes.length === 0) && state.lastExport?.outputPath) {
          state.lastBytes = await context.readFileBytes(state.lastExport.outputPath);
        }
        if (!state.lastBytes || state.lastBytes.length === 0) {
          throw new Error("Нет экспортированного файла для загрузки.");
        }
        const fileName = state.lastExport?.outputPath
          ? state.lastExport.outputPath.split(/[\\/]/).pop()
          : `${sanitizeFileStem(state.fileName || "table")}.tbl`;
        context.downloadBytes(state.lastBytes, fileName || "table.tbl");
        context.logger.info("download", fileName || "table.tbl");
      }

      function resetState() {
        state.fileName = "";
        state.sheetName = "";
        state.matrix = [];
        state.cellMatrix = [];
        state.lastExport = null;
        state.lastBytes = null;
        state.autoFollowOutput = true;
        refs.fileInput.value = "";
        refs.outputPath.value = "";
        refs.resultBox.textContent = "Ожидание данных.";
        renderAll();
      }

      async function handleWorkbookFile(file) {
        if (!file) {
          return;
        }
        const parsed = readWorkbookMatrix(file, await file.arrayBuffer());
        state.fileName = parsed.fileName;
        state.sheetName = parsed.sheetName;
        state.matrix = parsed.matrix;
        state.cellMatrix = parsed.cellMatrix;
        state.lastExport = null;
        state.lastBytes = null;
        state.autoFollowOutput = true;
        const { rows, cols } = currentDimensions();
        state.layout = reconcileLinkedLayout(state.layout, rows || 1, cols || 1, state.layoutDriver);
        persistLayout();
        refs.resultBox.textContent = "Matrix loaded. Export or Inline is ready.";
        context.logger.info("xlsx-parsed", `${state.fileName} ${formatSize(rows, cols)}`);
        await ensureTempPathLoaded();
        renderAll();
      }

      function bindLayoutInput(input, property, driver) {
        input.addEventListener("input", () => {
          applyLayoutChange(driver, {
            [property]: readNumeric(input, DEFAULT_LAYOUT[property]),
          });
        });
      }

      refs.fileInput.addEventListener("change", (event) => {
        handleWorkbookFile(event.target.files?.[0]).catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("xlsx-parse", refs.resultBox.textContent);
        });
      });

      refs.dropzone.addEventListener("dragenter", (event) => {
        event.preventDefault();
        state.dragDepth += 1;
        refs.dropzone.classList.add("is-dragging");
      });
      refs.dropzone.addEventListener("dragover", (event) => {
        event.preventDefault();
      });
      refs.dropzone.addEventListener("dragleave", (event) => {
        event.preventDefault();
        state.dragDepth = Math.max(0, state.dragDepth - 1);
        if (state.dragDepth === 0) {
          refs.dropzone.classList.remove("is-dragging");
        }
      });
      refs.dropzone.addEventListener("drop", (event) => {
        event.preventDefault();
        state.dragDepth = 0;
        refs.dropzone.classList.remove("is-dragging");
        const file = event.dataTransfer?.files?.[0];
        handleWorkbookFile(file).catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("drop", refs.resultBox.textContent);
        });
      });

      refs.outputPath.addEventListener("input", () => {
        state.autoFollowOutput = false;
        renderOutputMode();
        updateActionState();
      });

      refs.followButton.addEventListener("click", async () => {
        state.autoFollowOutput = true;
        await ensureTempPathLoaded();
        renderAll();
      });

      refs.exportButton.addEventListener("click", () => {
        exportTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("export", refs.resultBox.textContent);
          updateActionState();
        });
      });

      refs.inlineButton.addEventListener("click", () => {
        inlineTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("inline", refs.resultBox.textContent);
          updateActionState();
        });
      });

      refs.insertButton.addEventListener("click", () => {
        insertTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("insert", refs.resultBox.textContent);
        });
      });

      refs.downloadButton.addEventListener("click", () => {
        downloadTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("download", refs.resultBox.textContent);
        });
      });

      refs.resetButton.addEventListener("click", resetState);

      bindLayoutInput(refs.tableWidth, "tableWidthMm", "table");
      bindLayoutInput(refs.tableHeight, "tableHeightMm", "table");
      bindLayoutInput(refs.cellWidth, "cellWidthMm", "cell");
      bindLayoutInput(refs.cellHeight, "cellHeightMm", "cell");

      const onRuntimeLoaded = () => {
        ensureTempPathLoaded(true)
          .then(() => refreshStatus({ quiet: true }))
          .catch(() => {})
          .finally(() => {
            syncPolling();
          });
      };
      const onBridgeDisconnected = () => {
        state.status = null;
        syncPolling();
        renderAll();
      };
      const onVisibilityChange = () => {
        syncPolling();
      };

      context.events.addEventListener("runtime-loaded", onRuntimeLoaded);
      context.events.addEventListener("bridge-disconnected", onBridgeDisconnected);
      document.addEventListener("visibilitychange", onVisibilityChange);

      renderAll();

      return {
        activate() {
          state.active = true;
          syncPolling();
          if (context.getBridgeState().runtimeReady) {
            return refreshStatus({ quiet: true });
          }
          return Promise.resolve();
        },
        deactivate() {
          state.active = false;
          syncPolling();
          return Promise.resolve();
        },
        dispose() {
          if (state.pollTimer !== null) {
            window.clearInterval(state.pollTimer);
            state.pollTimer = null;
          }
          context.events.removeEventListener("runtime-loaded", onRuntimeLoaded);
          context.events.removeEventListener("bridge-disconnected", onBridgeDisconnected);
          document.removeEventListener("visibilitychange", onVisibilityChange);
        },
      };
    },
  };
}

export {
  DEFAULT_LAYOUT,
  EXPORT_BATCH_SIZE,
  buildWorkbookStyleContext,
  buildAutoOutputPath,
  buildInlineTempOutputPath,
  collectFormattedCellWrites,
  createCellTransferPayload,
  createCellWriteBatches,
  createFormattedCellWriteBatches,
  createXlsxToKompasTblModule,
  parseWorksheetCellMeta,
  readWorkbookMatrix,
  reconcileLinkedLayout,
};
