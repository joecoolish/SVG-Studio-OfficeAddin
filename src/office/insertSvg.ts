/* global Office */

/**
 * Inserts SVG markup into the active PowerPoint slide.
 *
 * PowerPoint supports inserting SVG as a native, scalable graphic via
 * `Office.context.document.setSelectedDataAsync` with
 * `Office.CoercionType.XmlSvg`. The SVG is placed at the current selection on the
 * active slide.
 */
export function insertSvgIntoSlide(
  svgMarkup: string,
  widthInches: number,
  heightInches: number
): Promise<void> {
  return new Promise((resolve, reject) => {
    if (typeof Office === "undefined" || !Office.context?.document) {
      reject(new Error("Office context is not available."));
      return;
    }

    const cleaned = ensureSvgRoot(svgMarkup);

    Office.context.document.setSelectedDataAsync(
      cleaned,
      {
        coercionType: Office.CoercionType.XmlSvg,
        imageWidth: widthInches * 72,
        imageHeight: heightInches * 72,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "Failed to insert SVG into the slide."));
        }
      }
    );
  });
}

/**
 * Ensures the markup is a standalone <svg> element. Strips XML prolog/doctype
 * and leading whitespace that some sources include, which PowerPoint rejects.
 */
export function ensureSvgRoot(markup: string): string {
  let s = markup.trim();
  const start = s.indexOf("<svg");
  if (start > 0) {
    s = s.slice(start);
  }
  const end = s.lastIndexOf("</svg>");
  if (end !== -1) {
    s = s.slice(0, end + "</svg>".length);
  }
  return s;
}
