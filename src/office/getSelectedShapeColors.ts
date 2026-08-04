/* global Office, PowerPoint */

export interface SelectedShapeProperties {
  widthInches: number;
  heightInches: number;
  colors: string[];
}

export async function getSelectedShapeColors(): Promise<string[]> {
  return (await getSelectedShapeProperties(false)).colors;
}

export async function getSelectedShapeProperties(
  requireSingleShape = true
): Promise<SelectedShapeProperties> {
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

    if (shapes.items.length === 0) {
      throw new Error("Select an image or shape on the slide, then try again.");
    }
    if (requireSingleShape && shapes.items.length !== 1) {
      throw new Error("Select exactly one image or shape, then try again.");
    }

    shapes.items.forEach((shape) => {
      shape.load("width,height");
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
    const selectedShape = shapes.items[0];
    return {
      widthInches: selectedShape.width / 72,
      heightInches: selectedShape.height / 72,
      colors,
    };
  });
}
