/* global DOMParser, Element, getComputedStyle, document, XMLDocument, XMLSerializer */

export interface PreparedSvg {
  markup: string;
  name: string;
  aspectRatio: number;
  colors: SvgColor[];
}

export interface SvgColor {
  source: string;
  value: string;
}

const paintAttributes = [
  "fill",
  "stroke",
  "color",
  "stop-color",
  "flood-color",
  "lighting-color",
] as const;
const forbiddenElements = new Set([
  "script",
  "foreignobject",
  "iframe",
  "object",
  "embed",
  "audio",
  "video",
  "canvas",
]);

export function prepareSvg(markup: string, name = "Untitled.svg"): PreparedSvg {
  if (markup.length > 2 * 1024 * 1024) {
    throw new Error("The SVG is larger than the 2 MB limit.");
  }

  const document = parseSvg(markup);
  sanitizeSvg(document);
  const root = document.documentElement;
  root.setAttribute("xmlns", "http://www.w3.org/2000/svg");

  const { width, height } = getSvgDimensions(root);
  const colors = collectColors(root);

  return {
    markup: new XMLSerializer().serializeToString(root),
    name: normalizeFileName(name),
    aspectRatio: width / height,
    colors,
  };
}

export function recolorSvg(
  markup: string,
  replacements: Readonly<Record<string, string>>,
  defaultFill?: string
): string {
  const document = parseSvg(markup);
  const root = document.documentElement;

  if (defaultFill) {
    root.setAttribute("fill", defaultFill);
  }

  getElements(root).forEach((element) => {
    paintAttributes.forEach((attribute) => {
      const value = element.getAttribute(attribute);
      const replacement = value ? replacements[value] : undefined;
      if (replacement) {
        element.setAttribute(attribute, replacement);
      }
    });

    const style = element.getAttribute("style");
    if (style) {
      element.setAttribute("style", replaceStylePaints(style, replacements));
    }
  });
  root.querySelectorAll("style").forEach((styleElement) => {
    styleElement.textContent = replaceCssPaints(styleElement.textContent ?? "", replacements);
  });

  return new XMLSerializer().serializeToString(root);
}

export function recolorAllSvg(markup: string, color: string): string {
  const document = parseSvg(markup);
  const root = document.documentElement;
  root.setAttribute("fill", color);
  root.setAttribute("color", color);

  getElements(root).forEach((element) => {
    paintAttributes.forEach((attribute) => {
      const value = element.getAttribute(attribute);
      if (value && value.trim().toLowerCase() !== "none" && !value.trim().startsWith("url(")) {
        element.setAttribute(attribute, color);
      }
    });

    const style = element.getAttribute("style");
    if (style) {
      const replacements = Object.fromEntries(
        collectStylePaints(style).map((paint) => [paint, color])
      );
      element.setAttribute("style", replaceStylePaints(style, replacements));
    }
  });
  root.querySelectorAll("style").forEach((styleElement) => {
    styleElement.textContent = replaceAllCssPaints(styleElement.textContent ?? "", color);
  });

  return new XMLSerializer().serializeToString(root);
}

function parseSvg(markup: string): XMLDocument {
  const document = new DOMParser().parseFromString(markup.trim(), "image/svg+xml");
  if (
    document.querySelector("parsererror") ||
    document.documentElement.localName.toLowerCase() !== "svg"
  ) {
    throw new Error("This file does not contain valid SVG markup.");
  }
  return document;
}

function sanitizeSvg(document: XMLDocument): void {
  document.querySelectorAll("*").forEach((element) => {
    if (forbiddenElements.has(element.localName.toLowerCase())) {
      element.remove();
      return;
    }

    if (element.localName.toLowerCase() === "style" && isUnsafeCss(element.textContent ?? "")) {
      element.remove();
      return;
    }

    Array.from(element.attributes).forEach((attribute) => {
      const name = attribute.name.toLowerCase();
      const value = attribute.value.trim();
      if (
        name.startsWith("on") ||
        (name === "style" && isUnsafeCss(value)) ||
        ((name === "href" || name.endsWith(":href")) && !isSafeReference(value))
      ) {
        element.removeAttribute(attribute.name);
      }
    });
  });
}

function isSafeReference(value: string): boolean {
  return value === "" || value.startsWith("#") || value.startsWith("data:image/");
}

function isUnsafeCss(value: string): boolean {
  const normalized = value.toLowerCase().replace(/\s/g, "");
  return (
    normalized.includes("javascript:") ||
    normalized.includes("expression(") ||
    /url\((?!['"]?#)/.test(normalized)
  );
}

function getSvgDimensions(root: Element): { width: number; height: number } {
  const viewBox = root
    .getAttribute("viewBox")
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);
  if (viewBox?.length === 4 && viewBox.every(Number.isFinite) && viewBox[2] > 0 && viewBox[3] > 0) {
    return { width: viewBox[2], height: viewBox[3] };
  }

  const width = parseLength(root.getAttribute("width"));
  const height = parseLength(root.getAttribute("height"));
  return {
    width: width > 0 ? width : 100,
    height: height > 0 ? height : 100,
  };
}

function parseLength(value: string | null): number {
  const parsed = Number.parseFloat(value ?? "");
  return Number.isFinite(parsed) ? parsed : 0;
}

function collectColors(root: Element): SvgColor[] {
  const colors = new Map<string, SvgColor>();

  getElements(root).forEach((element) => {
    paintAttributes.forEach((attribute) => {
      const value = element.getAttribute(attribute);
      addColor(colors, value);
    });
    collectStylePaints(element.getAttribute("style") ?? "").forEach((value) =>
      addColor(colors, value)
    );
  });
  root.querySelectorAll("style").forEach((styleElement) => {
    collectCssPaints(styleElement.textContent ?? "").forEach((value) => addColor(colors, value));
  });

  return Array.from(colors.values());
}

function collectStylePaints(style: string): string[] {
  return style
    .split(";")
    .map((declaration) => declaration.split(":"))
    .filter(
      ([property, value]) =>
        paintAttributes.includes(
          property?.trim().toLowerCase() as (typeof paintAttributes)[number]
        ) && Boolean(value)
    )
    .map(([, value]) => value.trim());
}

function collectCssPaints(css: string): string[] {
  const values: string[] = [];
  const pattern = new RegExp(`(?:${paintAttributes.join("|")})\\s*:\\s*([^;}]+)`, "gi");
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(css)) !== null) {
    values.push(match[1].trim());
  }
  return values;
}

function addColor(colors: Map<string, SvgColor>, value: string | null): void {
  if (!value || value.trim().toLowerCase() === "none" || value.trim().startsWith("url(")) {
    return;
  }
  const normalized = toHexColor(value);
  if (normalized && !colors.has(value)) {
    colors.set(value, { source: value, value: normalized });
  }
}

function toHexColor(value: string): string | undefined {
  const probe = document.createElement("span");
  probe.style.color = "";
  probe.style.color = value;
  if (!probe.style.color) {
    return undefined;
  }

  document.body.appendChild(probe);
  const normalized = getComputedStyle(probe).color;
  probe.remove();
  const match = normalized.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!match) {
    return undefined;
  }
  return `#${[match[1], match[2], match[3]]
    .map((component) => Number(component).toString(16).padStart(2, "0"))
    .join("")}`;
}

function replaceStylePaints(style: string, replacements: Readonly<Record<string, string>>): string {
  return style
    .split(";")
    .map((declaration) => {
      const separator = declaration.indexOf(":");
      if (separator === -1) {
        return declaration;
      }
      const property = declaration.slice(0, separator).trim();
      const value = declaration.slice(separator + 1).trim();
      if (
        paintAttributes.includes(property.toLowerCase() as (typeof paintAttributes)[number]) &&
        replacements[value]
      ) {
        return `${property}: ${replacements[value]}`;
      }
      return declaration;
    })
    .join("; ");
}

function replaceCssPaints(css: string, replacements: Readonly<Record<string, string>>): string {
  const pattern = new RegExp(`(${paintAttributes.join("|")})(\\s*:\\s*)([^;}]+)`, "gi");
  return css.replace(pattern, (declaration, property: string, separator: string, value: string) => {
    const trimmed = value.trim();
    return replacements[trimmed] ? `${property}${separator}${replacements[trimmed]}` : declaration;
  });
}

function replaceAllCssPaints(css: string, color: string): string {
  const pattern = new RegExp(`(${paintAttributes.join("|")})(\\s*:\\s*)([^;}]+)`, "gi");
  return css.replace(pattern, (declaration, property: string, separator: string, value: string) => {
    const trimmed = value.trim().toLowerCase();
    return trimmed === "none" || trimmed.startsWith("url(")
      ? declaration
      : `${property}${separator}${color}`;
  });
}

function normalizeFileName(name: string): string {
  const trimmed = name.trim() || "Untitled.svg";
  return trimmed.toLowerCase().endsWith(".svg") ? trimmed : `${trimmed}.svg`;
}

function getElements(root: Element): Element[] {
  return [root, ...Array.from(root.querySelectorAll("*"))];
}
