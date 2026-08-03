/* global Office, PowerPoint */

export async function getSelectedShapeColors(): Promise<string[]> {
  if (
    typeof Office === "undefined" ||
    !Office.context.requirements.isSetSupported("PowerPointApi", "1.5")
  ) {
    throw new Error("Reading colors from selected shapes requires PowerPointApi 1.5.");
  }

  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();

    shapes.items.forEach((shape) => {
      shape.fill.load("foregroundColor,type");
      shape.lineFormat.load("color,visible");
    });
    await context.sync();

    const colors: string[] = [];
    shapes.items.forEach((shape) => {
      if (shape.fill.type === PowerPoint.ShapeFillType.solid && shape.fill.foregroundColor) {
        colors.push(shape.fill.foregroundColor);
      }
      if (shape.lineFormat.visible && shape.lineFormat.color) {
        colors.push(shape.lineFormat.color);
      }
    });
    return colors;
  });
}
