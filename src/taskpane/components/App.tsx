import * as React from "react";
import {
  Button,
  Checkbox,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  SplitButton,
  Spinner,
  Text,
  makeStyles,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import Header from "./Header";
import { getSelectedShapeColors } from "../../office/getSelectedShapeColors";
import { insertSvgIntoSlide } from "../../office/insertSvg";
import { PreparedSvg, prepareSvg, recolorAllSvg, recolorSvg } from "../../svg/svgTools";

type SizeUnit = "in" | "px";
type ReplacementMode = "preserve" | "reset";

const pixelsPerInch = 96;
const recentColorsKey = "svg-studio-recent-colors";
const sizeUnitKey = "svg-studio-size-unit";
const defaultPalette = [
  "#000000",
  "#ffffff",
  "#7f7f7f",
  "#4472c4",
  "#5b9bd5",
  "#00b0f0",
  "#70ad47",
  "#ffc000",
  "#ed7d31",
  "#c00000",
  "#a5a5a5",
  "#7030a0",
];

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    minHeight: "100vh",
    boxSizing: "border-box",
    padding: "12px",
    gap: "12px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  dropZone: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    minHeight: "150px",
    padding: "20px 12px",
    textAlign: "center",
    borderRadius: tokens.borderRadiusLarge,
    ...shorthands.border("2px", "dashed", tokens.colorNeutralStroke1),
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: "pointer",
    ":hover": {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  dropZoneActive: {
    ...shorthands.borderColor(tokens.colorBrandStroke1),
    backgroundColor: tokens.colorBrandBackground2,
  },
  hiddenInput: {
    display: "none",
  },
  workspace: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  fileRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  fileActions: {
    display: "flex",
    gap: "6px",
    flexShrink: 0,
  },
  splitButton: {
    minWidth: "88px",
  },
  dropActions: {
    display: "flex",
    flexWrap: "wrap",
    justifyContent: "center",
    gap: "8px",
  },
  fileName: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  preview: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    minHeight: "210px",
    maxHeight: "330px",
    padding: "20px",
    overflow: "hidden",
    borderRadius: tokens.borderRadiusLarge,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: "#ffffff",
    backgroundImage:
      "linear-gradient(45deg, #f3f3f3 25%, transparent 25%), linear-gradient(-45deg, #f3f3f3 25%, transparent 25%), linear-gradient(45deg, transparent 75%, #f3f3f3 75%), linear-gradient(-45deg, transparent 75%, #f3f3f3 75%)",
    backgroundSize: "20px 20px",
    backgroundPosition: "0 0, 0 10px, 10px -10px, -10px 0px",
    "& svg": {
      width: "100%",
      height: "100%",
      maxWidth: "270px",
      maxHeight: "270px",
    },
  },
  panel: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
  },
  panelTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  panelHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  unitSelect: {
    minHeight: "30px",
    padding: "3px 8px",
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke1),
    color: tokens.colorNeutralForeground1,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  palette: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(120px, 1fr))",
    gap: "8px",
  },
  colorControl: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    minWidth: 0,
  },
  selectedColorControl: {
    borderRadius: tokens.borderRadiusMedium,
    outline: `2px solid ${tokens.colorBrandStroke1}`,
    outlineOffset: "2px",
  },
  quickPalette: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
  },
  swatch: {
    width: "28px",
    height: "28px",
    padding: 0,
    borderRadius: tokens.borderRadiusSmall,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke1),
    cursor: "pointer",
  },
  paletteActions: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    flexWrap: "wrap",
    gap: "8px",
  },
  colorInput: {
    width: "38px",
    height: "32px",
    padding: "1px",
    flexShrink: 0,
  },
  colorLabel: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
  },
  sizeRow: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "10px",
  },
  field: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  numberInput: {
    width: "100%",
    boxSizing: "border-box",
    minHeight: "32px",
    padding: "5px 8px",
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke1),
    color: tokens.colorNeutralForeground1,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  actions: {
    display: "flex",
    gap: "8px",
  },
});

interface AppProps {
  darkMode: boolean;
  onToggleTheme: () => void;
}

const App: React.FC<AppProps> = ({ darkMode, onToggleTheme }) => {
  const styles = useStyles();
  const inputRef = React.useRef<HTMLInputElement>(null);
  const fileReplacementMode = React.useRef<ReplacementMode>("reset");
  const [svg, setSvg] = React.useState<PreparedSvg>();
  const [replacements, setReplacements] = React.useState<Record<string, string>>({});
  const [singleColor, setSingleColor] = React.useState(false);
  const [allColor, setAllColor] = React.useState("#0078d4");
  const [activeColorSource, setActiveColorSource] = React.useState<string>();
  const [recentColors, setRecentColors] = React.useState<string[]>(loadRecentColors);
  const [width, setWidth] = React.useState(2);
  const [height, setHeight] = React.useState(2);
  const [sizeUnit, setSizeUnit] = React.useState<SizeUnit>(loadSizeUnit);
  const [widthInput, setWidthInput] = React.useState("2");
  const [heightInput, setHeightInput] = React.useState("2");
  const [lockAspect, setLockAspect] = React.useState(true);
  const [dragging, setDragging] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [readingClipboard, setReadingClipboard] = React.useState(false);
  const [readingPowerPointColors, setReadingPowerPointColors] = React.useState(false);
  const [message, setMessage] = React.useState<{ intent: "error" | "success"; text: string }>();

  const quickPalette = React.useMemo(
    () => Array.from(new Set([...recentColors, ...defaultPalette])).slice(0, 24),
    [recentColors]
  );

  const renderedMarkup = React.useMemo(() => {
    if (!svg) {
      return "";
    }
    return singleColor ? recolorAllSvg(svg.markup, allColor) : recolorSvg(svg.markup, replacements);
  }, [allColor, replacements, singleColor, svg]);

  const loadMarkup = React.useCallback(
    (markup: string, name?: string, mode: ReplacementMode = "reset") => {
      try {
        const prepared = prepareSvg(extractSvgMarkup(markup), name);
        const initialWidth =
          prepared.aspectRatio >= 1 ? 2 : Math.max(0.25, 2 * prepared.aspectRatio);
        const initialHeight =
          prepared.aspectRatio >= 1 ? Math.max(0.25, 2 / prepared.aspectRatio) : 2;
        const preserveSettings = mode === "preserve" && svg;
        const currentColors = svg?.colors.map(
          (color) => replacements[color.source] ?? color.value
        );
        setSvg(prepared);
        setReplacements(
          Object.fromEntries(
            prepared.colors.map((color, index) => [
              color.source,
              preserveSettings ? (currentColors?.[index] ?? color.value) : color.value,
            ])
          )
        );
        setActiveColorSource(prepared.colors[0]?.source);
        if (!preserveSettings) {
          setWidth(roundSize(initialWidth));
          setHeight(roundSize(initialHeight));
          setWidthInput(String(fromInches(initialWidth, sizeUnit)));
          setHeightInput(String(fromInches(initialHeight, sizeUnit)));
          setSingleColor(false);
        }
        setMessage(undefined);
      } catch (error) {
        setMessage({ intent: "error", text: (error as Error).message });
      }
    },
    [replacements, sizeUnit, svg]
  );

  const loadFile = React.useCallback(
    async (file: File, mode: ReplacementMode = "reset") => {
      if (!file.name.toLowerCase().endsWith(".svg") && file.type !== "image/svg+xml") {
        setMessage({ intent: "error", text: "Choose an .svg file." });
        return;
      }
      try {
        loadMarkup(await file.text(), file.name, mode);
      } catch (error) {
        setMessage({
          intent: "error",
          text: `Could not read the SVG: ${(error as Error).message}`,
        });
      }
    },
    [loadMarkup]
  );

  React.useEffect(() => {
    const handlePaste = (event: ClipboardEvent) => {
      const svgFile = Array.from(event.clipboardData?.files ?? []).find(
        (file) => file.type === "image/svg+xml" || file.name.toLowerCase().endsWith(".svg")
      );
      if (svgFile) {
        event.preventDefault();
        void loadFile(svgFile);
        return;
      }

      const text = event.clipboardData?.getData("text/plain") ?? "";
      if (text.includes("<svg")) {
        event.preventDefault();
        loadMarkup(text, "Pasted SVG.svg");
      }
    };
    window.addEventListener("paste", handlePaste);
    return () => window.removeEventListener("paste", handlePaste);
  }, [loadFile, loadMarkup]);

  const readClipboard = async (mode: ReplacementMode = "reset") => {
    if (!navigator.clipboard) {
      setMessage({
        intent: "error",
        text: "Clipboard access is not available in this PowerPoint version.",
      });
      return;
    }

    setReadingClipboard(true);
    setMessage(undefined);
    try {
      if (navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          for (const type of ["image/svg+xml", "text/plain", "text/html"]) {
            if (!item.types.includes(type)) {
              continue;
            }
            const markup = await (await item.getType(type)).text();
            if (markup.includes("<svg")) {
              loadMarkup(markup, "Clipboard SVG.svg", mode);
              return;
            }
          }
        }
      } else {
        const markup = await navigator.clipboard.readText();
        if (markup.includes("<svg")) {
          loadMarkup(markup, "Clipboard SVG.svg", mode);
          return;
        }
      }

      setMessage({
        intent: "error",
        text: "The clipboard does not contain SVG markup.",
      });
    } catch (error) {
      setMessage({
        intent: "error",
        text: `PowerPoint could not read the clipboard: ${(error as Error).message}`,
      });
    } finally {
      setReadingClipboard(false);
    }
  };

  const commitWidth = () => {
    const parsed = Number(widthInput);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      setWidthInput(String(fromInches(width, sizeUnit)));
      return;
    }
    const value = clampSize(toInches(parsed, sizeUnit), sizeUnit);
    setWidth(value);
    if (lockAspect && svg) {
      const nextHeight = value / svg.aspectRatio;
      setHeight(nextHeight);
      setHeightInput(String(fromInches(nextHeight, sizeUnit)));
    }
    setWidthInput(String(fromInches(value, sizeUnit)));
  };

  const changeWidthInput = (raw: string) => {
    setWidthInput(raw);
    const parsed = Number(raw);
    if (!lockAspect || !svg || !Number.isFinite(parsed) || parsed <= 0) {
      return;
    }

    const nextWidth = toInches(parsed, sizeUnit);
    const nextHeight = nextWidth / svg.aspectRatio;
    setWidth(nextWidth);
    setHeight(nextHeight);
    setHeightInput(String(fromInches(nextHeight, sizeUnit)));
  };

  const commitHeight = () => {
    const parsed = Number(heightInput);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      setHeightInput(String(fromInches(height, sizeUnit)));
      return;
    }
    const value = clampSize(toInches(parsed, sizeUnit), sizeUnit);
    setHeight(value);
    if (lockAspect && svg) {
      const nextWidth = value * svg.aspectRatio;
      setWidth(nextWidth);
      setWidthInput(String(fromInches(nextWidth, sizeUnit)));
    }
    setHeightInput(String(fromInches(value, sizeUnit)));
  };

  const changeHeightInput = (raw: string) => {
    setHeightInput(raw);
    const parsed = Number(raw);
    if (!lockAspect || !svg || !Number.isFinite(parsed) || parsed <= 0) {
      return;
    }

    const nextHeight = toInches(parsed, sizeUnit);
    const nextWidth = nextHeight * svg.aspectRatio;
    setHeight(nextHeight);
    setWidth(nextWidth);
    setWidthInput(String(fromInches(nextWidth, sizeUnit)));
  };

  const changeSizeUnit = (nextUnit: SizeUnit) => {
    setSizeUnit(nextUnit);
    saveSizeUnit(nextUnit);
    setWidthInput(String(fromInches(width, nextUnit)));
    setHeightInput(String(fromInches(height, nextUnit)));
  };

  const rememberColors = React.useCallback((colors: string[]) => {
    setRecentColors((current) => {
      const next = Array.from(
        new Set([...colors.map(normalizePaletteColor).filter(isColor), ...current])
      ).slice(0, 12);
      saveRecentColors(next);
      return next;
    });
  }, []);

  const applyQuickColor = (color: string) => {
    if (singleColor || !activeColorSource) {
      setAllColor(color);
      setSingleColor(true);
    } else {
      setReplacements((current) => ({ ...current, [activeColorSource]: color }));
    }
    rememberColors([color]);
  };

  const readPowerPointColors = async () => {
    setReadingPowerPointColors(true);
    setMessage(undefined);
    try {
      const colors = (await getSelectedShapeColors()).map(normalizePaletteColor).filter(isColor);
      if (colors.length === 0) {
        setMessage({
          intent: "error",
          text: "Select one or more solid-color PowerPoint shapes, then try again.",
        });
        return;
      }
      rememberColors(colors);
      setMessage({
        intent: "success",
        text: `Added ${new Set(colors).size} color(s) from the selected PowerPoint shapes.`,
      });
    } catch (error) {
      setMessage({ intent: "error", text: (error as Error).message });
    } finally {
      setReadingPowerPointColors(false);
    }
  };

  const handleInsert = async () => {
    if (!svg) {
      return;
    }
    setBusy(true);
    setMessage(undefined);
    try {
      await insertSvgIntoSlide(renderedMarkup, width, height);
      setMessage({ intent: "success", text: "SVG inserted on the current slide." });
    } catch (error) {
      setMessage({ intent: "error", text: (error as Error).message });
    } finally {
      setBusy(false);
    }
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    const file = Array.from(event.dataTransfer.files).find(
      (candidate) =>
        candidate.type === "image/svg+xml" || candidate.name.toLowerCase().endsWith(".svg")
    );
    if (file) {
      void loadFile(file);
      return;
    }
    const text = event.dataTransfer.getData("text/plain");
    if (text.includes("<svg")) {
      loadMarkup(text, "Dropped SVG.svg");
    } else {
      setMessage({ intent: "error", text: "Drop an SVG file or SVG markup." });
    }
  };

  const chooseReplacementFile = (mode: ReplacementMode) => {
    fileReplacementMode.current = mode;
    inputRef.current?.click();
  };

  return (
    <main className={styles.root}>
      <Header darkMode={darkMode} onToggleTheme={onToggleTheme} />

      {message && (
        <MessageBar intent={message.intent}>
          <MessageBarBody>
            <MessageBarTitle>{message.intent === "success" ? "Done" : "SVG error"}</MessageBarTitle>
            {message.text}
          </MessageBarBody>
        </MessageBar>
      )}

      {!svg ? (
        <div
          className={`${styles.dropZone} ${dragging ? styles.dropZoneActive : ""}`}
          role="button"
          tabIndex={0}
          onClick={() => inputRef.current?.click()}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") {
              inputRef.current?.click();
            }
          }}
          onDragEnter={(event) => {
            event.preventDefault();
            setDragging(true);
          }}
          onDragOver={(event) => event.preventDefault()}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
        >
          <Text size={500} weight="semibold">
            Add an SVG
          </Text>
          <Text>Drop it here, read the clipboard, or choose a file.</Text>
          <div className={styles.dropActions}>
            <Button appearance="primary">Choose SVG</Button>
            <Button
              disabled={readingClipboard}
              onClick={(event) => {
                event.stopPropagation();
                void readClipboard();
              }}
            >
              {readingClipboard ? <Spinner size="tiny" /> : "Read clipboard"}
            </Button>
          </div>
        </div>
      ) : (
        <div className={styles.workspace}>
          <div className={styles.fileRow}>
            <Text className={styles.fileName} weight="semibold">
              {svg.name}
            </Text>
            <div className={styles.fileActions}>
              <Menu positioning="below-end">
                <MenuTrigger disableButtonEnhancement>
                  {(triggerProps) => (
                    <SplitButton
                      {...triggerProps}
                      className={styles.splitButton}
                      size="small"
                      disabled={readingClipboard}
                      primaryActionButton={{
                        onClick: () => void readClipboard("preserve"),
                      }}
                      menuButton={{ "aria-label": "More clipboard replacement options" }}
                    >
                      {readingClipboard ? <Spinner size="tiny" /> : "Clipboard"}
                    </SplitButton>
                  )}
                </MenuTrigger>
                <MenuPopover>
                  <MenuList>
                    <MenuItem onClick={() => void readClipboard("reset")}>
                      Replace and reset to clipboard defaults
                    </MenuItem>
                  </MenuList>
                </MenuPopover>
              </Menu>
              <Menu positioning="below-end">
                <MenuTrigger disableButtonEnhancement>
                  {(triggerProps) => (
                    <SplitButton
                      {...triggerProps}
                      className={styles.splitButton}
                      size="small"
                      primaryActionButton={{ onClick: () => chooseReplacementFile("preserve") }}
                      menuButton={{ "aria-label": "More file replacement options" }}
                    >
                      File
                    </SplitButton>
                  )}
                </MenuTrigger>
                <MenuPopover>
                  <MenuList>
                    <MenuItem onClick={() => chooseReplacementFile("reset")}>
                      Replace and reset to file defaults
                    </MenuItem>
                  </MenuList>
                </MenuPopover>
              </Menu>
            </div>
          </div>

          <div
            className={styles.preview}
            aria-label="SVG preview"
            dangerouslySetInnerHTML={{ __html: renderedMarkup }}
          />

          <section className={styles.panel}>
            <Text className={styles.panelTitle}>Colors</Text>
            <Checkbox
              checked={singleColor}
              label="Use one color for the entire SVG"
              onChange={(_, data) => setSingleColor(Boolean(data.checked))}
            />
            {singleColor ? (
              <div className={styles.colorControl}>
                <input
                  className={styles.colorInput}
                  type="color"
                  value={allColor}
                  aria-label="SVG color"
                  onChange={(event) => {
                    setAllColor(event.target.value);
                    rememberColors([event.target.value]);
                  }}
                />
                <Text className={styles.colorLabel}>{allColor}</Text>
              </div>
            ) : svg.colors.length > 0 ? (
              <div className={styles.palette}>
                {svg.colors.map((color) => (
                  <label
                    className={`${styles.colorControl} ${
                      activeColorSource === color.source ? styles.selectedColorControl : ""
                    }`}
                    key={color.source}
                    onClick={() => setActiveColorSource(color.source)}
                  >
                    <input
                      className={styles.colorInput}
                      type="color"
                      value={replacements[color.source] ?? color.value}
                      onChange={(event) => {
                        setReplacements((current) => ({
                          ...current,
                          [color.source]: event.target.value,
                        }));
                        rememberColors([event.target.value]);
                      }}
                    />
                    <span className={styles.colorLabel} title={color.source}>
                      {color.source}
                    </span>
                  </label>
                ))}
              </div>
            ) : (
              <Text size={200}>
                This SVG uses the default black fill. Enable one-color mode to change it.
              </Text>
            )}
            <div className={styles.paletteActions}>
              <Text size={200}>
                {singleColor ? "Quick colors" : "Quick colors for the selected SVG color"}
              </Text>
              <Button
                size="small"
                disabled={readingPowerPointColors}
                onClick={() => void readPowerPointColors()}
              >
                {readingPowerPointColors ? <Spinner size="tiny" /> : "Colors from selection"}
              </Button>
            </div>
            <div className={styles.quickPalette}>
              {quickPalette.map((color) => (
                <button
                  className={styles.swatch}
                  key={color}
                  type="button"
                  title={color}
                  aria-label={`Use color ${color}`}
                  style={{ backgroundColor: color }}
                  onClick={() => applyQuickColor(color)}
                />
              ))}
            </div>
          </section>

          <section className={styles.panel}>
            <div className={styles.panelHeader}>
              <Text className={styles.panelTitle}>Size on slide</Text>
              <select
                className={styles.unitSelect}
                value={sizeUnit}
                aria-label="Size unit"
                onChange={(event) => changeSizeUnit(event.target.value as SizeUnit)}
              >
                <option value="in">Inches</option>
                <option value="px">Pixels</option>
              </select>
            </div>
            <div className={styles.sizeRow}>
              <label className={styles.field}>
                <Text size={200}>Width ({sizeUnit === "in" ? "inches" : "pixels"})</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={sizeUnit === "in" ? "0.1" : "1"}
                  max={sizeUnit === "in" ? "20" : "1920"}
                  step={sizeUnit === "in" ? "0.1" : "1"}
                  value={widthInput}
                  onChange={(event) => changeWidthInput(event.target.value)}
                  onBlur={commitWidth}
                />
              </label>
              <label className={styles.field}>
                <Text size={200}>Height ({sizeUnit === "in" ? "inches" : "pixels"})</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={sizeUnit === "in" ? "0.1" : "1"}
                  max={sizeUnit === "in" ? "20" : "1920"}
                  step={sizeUnit === "in" ? "0.1" : "1"}
                  value={heightInput}
                  onChange={(event) => changeHeightInput(event.target.value)}
                  onBlur={commitHeight}
                />
              </label>
            </div>
            <Checkbox
              checked={lockAspect}
              label="Lock aspect ratio"
              onChange={(_, data) => setLockAspect(Boolean(data.checked))}
            />
          </section>

          <div className={styles.actions}>
            <Button appearance="primary" disabled={busy} onClick={handleInsert}>
              {busy ? <Spinner size="tiny" /> : "Insert into PowerPoint"}
            </Button>
            <Button
              disabled={busy}
              onClick={() => {
                setReplacements(
                  Object.fromEntries(svg.colors.map((color) => [color.source, color.value]))
                );
                setSingleColor(false);
              }}
            >
              Reset colors
            </Button>
          </div>
        </div>
      )}

      <input
        ref={inputRef}
        className={styles.hiddenInput}
        type="file"
        accept=".svg,image/svg+xml"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            void loadFile(file, svg ? fileReplacementMode.current : "reset");
          }
          fileReplacementMode.current = "reset";
          event.target.value = "";
        }}
      />
    </main>
  );
};

function clampSize(value: number, unit: SizeUnit): number {
  const minimum = unit === "px" ? 1 / pixelsPerInch : 0.1;
  return Math.min(20, Math.max(minimum, Number.isFinite(value) ? value : minimum));
}

function roundSize(value: number): number {
  return Math.round(value * 100) / 100;
}

function toInches(value: number, unit: SizeUnit): number {
  return unit === "px" ? value / pixelsPerInch : value;
}

function fromInches(value: number, unit: SizeUnit): number {
  return unit === "px" ? Math.round(value * pixelsPerInch) : roundSize(value);
}

function loadRecentColors(): string[] {
  try {
    const stored = JSON.parse(localStorage.getItem(recentColorsKey) ?? "[]") as unknown;
    return Array.isArray(stored)
      ? stored
          .map((color) => normalizePaletteColor(String(color)))
          .filter(isColor)
          .slice(0, 12)
      : [];
  } catch {
    return [];
  }
}

function saveRecentColors(colors: string[]): void {
  try {
    localStorage.setItem(recentColorsKey, JSON.stringify(colors));
  } catch {
    // The palette still works for this session when persistent storage is unavailable.
  }
}

function loadSizeUnit(): SizeUnit {
  try {
    return localStorage.getItem(sizeUnitKey) === "in" ? "in" : "px";
  } catch {
    return "px";
  }
}

function saveSizeUnit(unit: SizeUnit): void {
  try {
    localStorage.setItem(sizeUnitKey, unit);
  } catch {
    // The selected unit still works for this session when persistent storage is unavailable.
  }
}

function normalizePaletteColor(color: string): string | undefined {
  const trimmed = color.trim();
  const hex = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  if (/^#[0-9a-f]{6}$/i.test(hex)) {
    return hex.toLowerCase();
  }

  const probe = document.createElement("span");
  probe.style.color = "";
  probe.style.color = trimmed;
  if (!probe.style.color) {
    return undefined;
  }
  document.body.appendChild(probe);
  const normalized = getComputedStyle(probe).color;
  probe.remove();
  const match = normalized.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  return match
    ? `#${[match[1], match[2], match[3]]
        .map((component) => Number(component).toString(16).padStart(2, "0"))
        .join("")}`
    : undefined;
}

function isColor(color: string | undefined): color is string {
  return Boolean(color);
}

function extractSvgMarkup(markup: string): string {
  const normalized = markup.toLowerCase();
  const start = normalized.indexOf("<svg");
  const end = normalized.lastIndexOf("</svg>");
  return start >= 0 && end > start ? markup.slice(start, end + "</svg>".length) : markup;
}

export default App;
