#!/usr/bin/env node

// src/cli.ts
import { existsSync, readFileSync as readFileSync2 } from "node:fs";
import { resolve as resolve2 } from "node:path";

// src/core/units.ts
var DPI = 96;
var PX_TO_PT = 0.75;
var pxToInches = (px) => px / DPI;
var pxToPoints = (px) => px * PX_TO_PT;
function clamp(value, lo, hi) {
  return Number.isFinite(value) ? Math.max(lo, Math.min(hi, value)) : lo;
}

// src/core/sanitize.ts
function sanitizeStrings(value) {
  return Array.isArray(value) ? value.filter((v) => typeof v === "string") : [];
}
function sanitizeFontSwaps(value) {
  return Array.isArray(value) ? value.filter(
    (v) => !!v && typeof v.from === "string" && typeof v.to === "string"
  ) : [];
}

// vendor/pptxgenjs/src/pptxgen.ts
import JSZip2 from "jszip";

// vendor/pptxgenjs/src/core-enums.ts
var EMU = 914400;
var ONEPT = 12700;
var CRLF = "\r\n";
var LAYOUT_IDX_SERIES_BASE = 2147483649;
var REGEX_HEX_COLOR = /^[0-9a-fA-F]{6}$/;
var LINEH_MODIFIER = 1.67;
var DEF_BULLET_MARGIN = 27;
var DEF_CELL_BORDER = { type: "solid", color: "666666", pt: 1 };
var DEF_CELL_MARGIN_IN = [0.05, 0.1, 0.05, 0.1];
var DEF_CHART_BORDER = { type: "solid", color: "363636", pt: 1 };
var DEF_CHART_GRIDLINE = { color: "888888", style: "solid", size: 1, cap: "flat" };
var DEF_FONT_COLOR = "000000";
var DEF_FONT_SIZE = 12;
var DEF_FONT_TITLE_SIZE = 18;
var DEF_PRES_LAYOUT = "LAYOUT_16x9";
var DEF_PRES_LAYOUT_NAME = "DEFAULT";
var DEF_SHAPE_LINE_COLOR = "333333";
var DEF_SHAPE_SHADOW = { type: "outer", blur: 3, offset: 23e3 / 12700, angle: 90, color: "000000", opacity: 0.35, rotateWithShape: true };
var DEF_SLIDE_MARGIN_IN = [0.5, 0.5, 0.5, 0.5];
var DEF_TEXT_SHADOW = { type: "outer", blur: 8, offset: 4, angle: 270, color: "000000", opacity: 0.75 };
var DEF_TEXT_GLOW = { size: 8, color: "FFFFFF", opacity: 0.75 };
var AXIS_ID_VALUE_PRIMARY = "2094734552";
var AXIS_ID_VALUE_SECONDARY = "2094734553";
var AXIS_ID_CATEGORY_PRIMARY = "2094734554";
var AXIS_ID_CATEGORY_SECONDARY = "2094734555";
var AXIS_ID_SERIES_PRIMARY = "2094734556";
var LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
var BARCHART_COLORS = [
  "C0504D",
  "4F81BD",
  "9BBB59",
  "8064A2",
  "4BACC6",
  "F79646",
  "628FC6",
  "C86360",
  "C0504D",
  "4F81BD",
  "9BBB59",
  "8064A2",
  "4BACC6",
  "F79646",
  "628FC6",
  "C86360"
];
var PIECHART_COLORS = [
  "5DA5DA",
  "FAA43A",
  "60BD68",
  "F17CB0",
  "B2912F",
  "B276B2",
  "DECF3F",
  "F15854",
  "A7A7A7",
  "5DA5DA",
  "FAA43A",
  "60BD68",
  "F17CB0",
  "B2912F",
  "B276B2",
  "DECF3F",
  "F15854",
  "A7A7A7"
];
var SLDNUMFLDID = "{F7021451-1387-4CA6-816F-3879F97B5CBC}";
var OutputType = /* @__PURE__ */ ((OutputType2) => {
  OutputType2["arraybuffer"] = "arraybuffer";
  OutputType2["base64"] = "base64";
  OutputType2["binarystring"] = "binarystring";
  OutputType2["blob"] = "blob";
  OutputType2["nodebuffer"] = "nodebuffer";
  OutputType2["uint8array"] = "uint8array";
  return OutputType2;
})(OutputType || {});
var ChartType = /* @__PURE__ */ ((ChartType2) => {
  ChartType2["area"] = "area";
  ChartType2["bar"] = "bar";
  ChartType2["bar3d"] = "bar3D";
  ChartType2["bubble"] = "bubble";
  ChartType2["bubble3d"] = "bubble3D";
  ChartType2["doughnut"] = "doughnut";
  ChartType2["line"] = "line";
  ChartType2["pie"] = "pie";
  ChartType2["radar"] = "radar";
  ChartType2["scatter"] = "scatter";
  return ChartType2;
})(ChartType || {});
var ShapeType = /* @__PURE__ */ ((ShapeType2) => {
  ShapeType2["accentBorderCallout1"] = "accentBorderCallout1";
  ShapeType2["accentBorderCallout2"] = "accentBorderCallout2";
  ShapeType2["accentBorderCallout3"] = "accentBorderCallout3";
  ShapeType2["accentCallout1"] = "accentCallout1";
  ShapeType2["accentCallout2"] = "accentCallout2";
  ShapeType2["accentCallout3"] = "accentCallout3";
  ShapeType2["actionButtonBackPrevious"] = "actionButtonBackPrevious";
  ShapeType2["actionButtonBeginning"] = "actionButtonBeginning";
  ShapeType2["actionButtonBlank"] = "actionButtonBlank";
  ShapeType2["actionButtonDocument"] = "actionButtonDocument";
  ShapeType2["actionButtonEnd"] = "actionButtonEnd";
  ShapeType2["actionButtonForwardNext"] = "actionButtonForwardNext";
  ShapeType2["actionButtonHelp"] = "actionButtonHelp";
  ShapeType2["actionButtonHome"] = "actionButtonHome";
  ShapeType2["actionButtonInformation"] = "actionButtonInformation";
  ShapeType2["actionButtonMovie"] = "actionButtonMovie";
  ShapeType2["actionButtonReturn"] = "actionButtonReturn";
  ShapeType2["actionButtonSound"] = "actionButtonSound";
  ShapeType2["arc"] = "arc";
  ShapeType2["bentArrow"] = "bentArrow";
  ShapeType2["bentUpArrow"] = "bentUpArrow";
  ShapeType2["bevel"] = "bevel";
  ShapeType2["blockArc"] = "blockArc";
  ShapeType2["borderCallout1"] = "borderCallout1";
  ShapeType2["borderCallout2"] = "borderCallout2";
  ShapeType2["borderCallout3"] = "borderCallout3";
  ShapeType2["bracePair"] = "bracePair";
  ShapeType2["bracketPair"] = "bracketPair";
  ShapeType2["callout1"] = "callout1";
  ShapeType2["callout2"] = "callout2";
  ShapeType2["callout3"] = "callout3";
  ShapeType2["can"] = "can";
  ShapeType2["chartPlus"] = "chartPlus";
  ShapeType2["chartStar"] = "chartStar";
  ShapeType2["chartX"] = "chartX";
  ShapeType2["chevron"] = "chevron";
  ShapeType2["chord"] = "chord";
  ShapeType2["circularArrow"] = "circularArrow";
  ShapeType2["cloud"] = "cloud";
  ShapeType2["cloudCallout"] = "cloudCallout";
  ShapeType2["corner"] = "corner";
  ShapeType2["cornerTabs"] = "cornerTabs";
  ShapeType2["cube"] = "cube";
  ShapeType2["curvedDownArrow"] = "curvedDownArrow";
  ShapeType2["curvedLeftArrow"] = "curvedLeftArrow";
  ShapeType2["curvedRightArrow"] = "curvedRightArrow";
  ShapeType2["curvedUpArrow"] = "curvedUpArrow";
  ShapeType2["custGeom"] = "custGeom";
  ShapeType2["decagon"] = "decagon";
  ShapeType2["diagStripe"] = "diagStripe";
  ShapeType2["diamond"] = "diamond";
  ShapeType2["dodecagon"] = "dodecagon";
  ShapeType2["donut"] = "donut";
  ShapeType2["doubleWave"] = "doubleWave";
  ShapeType2["downArrow"] = "downArrow";
  ShapeType2["downArrowCallout"] = "downArrowCallout";
  ShapeType2["ellipse"] = "ellipse";
  ShapeType2["ellipseRibbon"] = "ellipseRibbon";
  ShapeType2["ellipseRibbon2"] = "ellipseRibbon2";
  ShapeType2["flowChartAlternateProcess"] = "flowChartAlternateProcess";
  ShapeType2["flowChartCollate"] = "flowChartCollate";
  ShapeType2["flowChartConnector"] = "flowChartConnector";
  ShapeType2["flowChartDecision"] = "flowChartDecision";
  ShapeType2["flowChartDelay"] = "flowChartDelay";
  ShapeType2["flowChartDisplay"] = "flowChartDisplay";
  ShapeType2["flowChartDocument"] = "flowChartDocument";
  ShapeType2["flowChartExtract"] = "flowChartExtract";
  ShapeType2["flowChartInputOutput"] = "flowChartInputOutput";
  ShapeType2["flowChartInternalStorage"] = "flowChartInternalStorage";
  ShapeType2["flowChartMagneticDisk"] = "flowChartMagneticDisk";
  ShapeType2["flowChartMagneticDrum"] = "flowChartMagneticDrum";
  ShapeType2["flowChartMagneticTape"] = "flowChartMagneticTape";
  ShapeType2["flowChartManualInput"] = "flowChartManualInput";
  ShapeType2["flowChartManualOperation"] = "flowChartManualOperation";
  ShapeType2["flowChartMerge"] = "flowChartMerge";
  ShapeType2["flowChartMultidocument"] = "flowChartMultidocument";
  ShapeType2["flowChartOfflineStorage"] = "flowChartOfflineStorage";
  ShapeType2["flowChartOffpageConnector"] = "flowChartOffpageConnector";
  ShapeType2["flowChartOnlineStorage"] = "flowChartOnlineStorage";
  ShapeType2["flowChartOr"] = "flowChartOr";
  ShapeType2["flowChartPredefinedProcess"] = "flowChartPredefinedProcess";
  ShapeType2["flowChartPreparation"] = "flowChartPreparation";
  ShapeType2["flowChartProcess"] = "flowChartProcess";
  ShapeType2["flowChartPunchedCard"] = "flowChartPunchedCard";
  ShapeType2["flowChartPunchedTape"] = "flowChartPunchedTape";
  ShapeType2["flowChartSort"] = "flowChartSort";
  ShapeType2["flowChartSummingJunction"] = "flowChartSummingJunction";
  ShapeType2["flowChartTerminator"] = "flowChartTerminator";
  ShapeType2["folderCorner"] = "folderCorner";
  ShapeType2["frame"] = "frame";
  ShapeType2["funnel"] = "funnel";
  ShapeType2["gear6"] = "gear6";
  ShapeType2["gear9"] = "gear9";
  ShapeType2["halfFrame"] = "halfFrame";
  ShapeType2["heart"] = "heart";
  ShapeType2["heptagon"] = "heptagon";
  ShapeType2["hexagon"] = "hexagon";
  ShapeType2["homePlate"] = "homePlate";
  ShapeType2["horizontalScroll"] = "horizontalScroll";
  ShapeType2["irregularSeal1"] = "irregularSeal1";
  ShapeType2["irregularSeal2"] = "irregularSeal2";
  ShapeType2["leftArrow"] = "leftArrow";
  ShapeType2["leftArrowCallout"] = "leftArrowCallout";
  ShapeType2["leftBrace"] = "leftBrace";
  ShapeType2["leftBracket"] = "leftBracket";
  ShapeType2["leftCircularArrow"] = "leftCircularArrow";
  ShapeType2["leftRightArrow"] = "leftRightArrow";
  ShapeType2["leftRightArrowCallout"] = "leftRightArrowCallout";
  ShapeType2["leftRightCircularArrow"] = "leftRightCircularArrow";
  ShapeType2["leftRightRibbon"] = "leftRightRibbon";
  ShapeType2["leftRightUpArrow"] = "leftRightUpArrow";
  ShapeType2["leftUpArrow"] = "leftUpArrow";
  ShapeType2["lightningBolt"] = "lightningBolt";
  ShapeType2["line"] = "line";
  ShapeType2["lineInv"] = "lineInv";
  ShapeType2["mathDivide"] = "mathDivide";
  ShapeType2["mathEqual"] = "mathEqual";
  ShapeType2["mathMinus"] = "mathMinus";
  ShapeType2["mathMultiply"] = "mathMultiply";
  ShapeType2["mathNotEqual"] = "mathNotEqual";
  ShapeType2["mathPlus"] = "mathPlus";
  ShapeType2["moon"] = "moon";
  ShapeType2["noSmoking"] = "noSmoking";
  ShapeType2["nonIsoscelesTrapezoid"] = "nonIsoscelesTrapezoid";
  ShapeType2["notchedRightArrow"] = "notchedRightArrow";
  ShapeType2["octagon"] = "octagon";
  ShapeType2["parallelogram"] = "parallelogram";
  ShapeType2["pentagon"] = "pentagon";
  ShapeType2["pie"] = "pie";
  ShapeType2["pieWedge"] = "pieWedge";
  ShapeType2["plaque"] = "plaque";
  ShapeType2["plaqueTabs"] = "plaqueTabs";
  ShapeType2["plus"] = "plus";
  ShapeType2["quadArrow"] = "quadArrow";
  ShapeType2["quadArrowCallout"] = "quadArrowCallout";
  ShapeType2["rect"] = "rect";
  ShapeType2["ribbon"] = "ribbon";
  ShapeType2["ribbon2"] = "ribbon2";
  ShapeType2["rightArrow"] = "rightArrow";
  ShapeType2["rightArrowCallout"] = "rightArrowCallout";
  ShapeType2["rightBrace"] = "rightBrace";
  ShapeType2["rightBracket"] = "rightBracket";
  ShapeType2["round1Rect"] = "round1Rect";
  ShapeType2["round2DiagRect"] = "round2DiagRect";
  ShapeType2["round2SameRect"] = "round2SameRect";
  ShapeType2["roundRect"] = "roundRect";
  ShapeType2["rtTriangle"] = "rtTriangle";
  ShapeType2["smileyFace"] = "smileyFace";
  ShapeType2["snip1Rect"] = "snip1Rect";
  ShapeType2["snip2DiagRect"] = "snip2DiagRect";
  ShapeType2["snip2SameRect"] = "snip2SameRect";
  ShapeType2["snipRoundRect"] = "snipRoundRect";
  ShapeType2["squareTabs"] = "squareTabs";
  ShapeType2["star10"] = "star10";
  ShapeType2["star12"] = "star12";
  ShapeType2["star16"] = "star16";
  ShapeType2["star24"] = "star24";
  ShapeType2["star32"] = "star32";
  ShapeType2["star4"] = "star4";
  ShapeType2["star5"] = "star5";
  ShapeType2["star6"] = "star6";
  ShapeType2["star7"] = "star7";
  ShapeType2["star8"] = "star8";
  ShapeType2["stripedRightArrow"] = "stripedRightArrow";
  ShapeType2["sun"] = "sun";
  ShapeType2["swooshArrow"] = "swooshArrow";
  ShapeType2["teardrop"] = "teardrop";
  ShapeType2["trapezoid"] = "trapezoid";
  ShapeType2["triangle"] = "triangle";
  ShapeType2["upArrow"] = "upArrow";
  ShapeType2["upArrowCallout"] = "upArrowCallout";
  ShapeType2["upDownArrow"] = "upDownArrow";
  ShapeType2["upDownArrowCallout"] = "upDownArrowCallout";
  ShapeType2["uturnArrow"] = "uturnArrow";
  ShapeType2["verticalScroll"] = "verticalScroll";
  ShapeType2["wave"] = "wave";
  ShapeType2["wedgeEllipseCallout"] = "wedgeEllipseCallout";
  ShapeType2["wedgeRectCallout"] = "wedgeRectCallout";
  ShapeType2["wedgeRoundRectCallout"] = "wedgeRoundRectCallout";
  return ShapeType2;
})(ShapeType || {});
var SchemeColor = /* @__PURE__ */ ((SchemeColor2) => {
  SchemeColor2["text1"] = "tx1";
  SchemeColor2["text2"] = "tx2";
  SchemeColor2["background1"] = "bg1";
  SchemeColor2["background2"] = "bg2";
  SchemeColor2["accent1"] = "accent1";
  SchemeColor2["accent2"] = "accent2";
  SchemeColor2["accent3"] = "accent3";
  SchemeColor2["accent4"] = "accent4";
  SchemeColor2["accent5"] = "accent5";
  SchemeColor2["accent6"] = "accent6";
  return SchemeColor2;
})(SchemeColor || {});
var AlignH = /* @__PURE__ */ ((AlignH2) => {
  AlignH2["left"] = "left";
  AlignH2["center"] = "center";
  AlignH2["right"] = "right";
  AlignH2["justify"] = "justify";
  return AlignH2;
})(AlignH || {});
var AlignV = /* @__PURE__ */ ((AlignV2) => {
  AlignV2["top"] = "top";
  AlignV2["middle"] = "middle";
  AlignV2["bottom"] = "bottom";
  return AlignV2;
})(AlignV || {});
var SHAPE_TYPE = /* @__PURE__ */ ((SHAPE_TYPE2) => {
  SHAPE_TYPE2["ACTION_BUTTON_BACK_OR_PREVIOUS"] = "actionButtonBackPrevious";
  SHAPE_TYPE2["ACTION_BUTTON_BEGINNING"] = "actionButtonBeginning";
  SHAPE_TYPE2["ACTION_BUTTON_CUSTOM"] = "actionButtonBlank";
  SHAPE_TYPE2["ACTION_BUTTON_DOCUMENT"] = "actionButtonDocument";
  SHAPE_TYPE2["ACTION_BUTTON_END"] = "actionButtonEnd";
  SHAPE_TYPE2["ACTION_BUTTON_FORWARD_OR_NEXT"] = "actionButtonForwardNext";
  SHAPE_TYPE2["ACTION_BUTTON_HELP"] = "actionButtonHelp";
  SHAPE_TYPE2["ACTION_BUTTON_HOME"] = "actionButtonHome";
  SHAPE_TYPE2["ACTION_BUTTON_INFORMATION"] = "actionButtonInformation";
  SHAPE_TYPE2["ACTION_BUTTON_MOVIE"] = "actionButtonMovie";
  SHAPE_TYPE2["ACTION_BUTTON_RETURN"] = "actionButtonReturn";
  SHAPE_TYPE2["ACTION_BUTTON_SOUND"] = "actionButtonSound";
  SHAPE_TYPE2["ARC"] = "arc";
  SHAPE_TYPE2["BALLOON"] = "wedgeRoundRectCallout";
  SHAPE_TYPE2["BENT_ARROW"] = "bentArrow";
  SHAPE_TYPE2["BENT_UP_ARROW"] = "bentUpArrow";
  SHAPE_TYPE2["BEVEL"] = "bevel";
  SHAPE_TYPE2["BLOCK_ARC"] = "blockArc";
  SHAPE_TYPE2["CAN"] = "can";
  SHAPE_TYPE2["CHART_PLUS"] = "chartPlus";
  SHAPE_TYPE2["CHART_STAR"] = "chartStar";
  SHAPE_TYPE2["CHART_X"] = "chartX";
  SHAPE_TYPE2["CHEVRON"] = "chevron";
  SHAPE_TYPE2["CHORD"] = "chord";
  SHAPE_TYPE2["CIRCULAR_ARROW"] = "circularArrow";
  SHAPE_TYPE2["CLOUD"] = "cloud";
  SHAPE_TYPE2["CLOUD_CALLOUT"] = "cloudCallout";
  SHAPE_TYPE2["CORNER"] = "corner";
  SHAPE_TYPE2["CORNER_TABS"] = "cornerTabs";
  SHAPE_TYPE2["CROSS"] = "plus";
  SHAPE_TYPE2["CUBE"] = "cube";
  SHAPE_TYPE2["CURVED_DOWN_ARROW"] = "curvedDownArrow";
  SHAPE_TYPE2["CURVED_DOWN_RIBBON"] = "ellipseRibbon";
  SHAPE_TYPE2["CURVED_LEFT_ARROW"] = "curvedLeftArrow";
  SHAPE_TYPE2["CURVED_RIGHT_ARROW"] = "curvedRightArrow";
  SHAPE_TYPE2["CURVED_UP_ARROW"] = "curvedUpArrow";
  SHAPE_TYPE2["CURVED_UP_RIBBON"] = "ellipseRibbon2";
  SHAPE_TYPE2["CUSTOM_GEOMETRY"] = "custGeom";
  SHAPE_TYPE2["DECAGON"] = "decagon";
  SHAPE_TYPE2["DIAGONAL_STRIPE"] = "diagStripe";
  SHAPE_TYPE2["DIAMOND"] = "diamond";
  SHAPE_TYPE2["DODECAGON"] = "dodecagon";
  SHAPE_TYPE2["DONUT"] = "donut";
  SHAPE_TYPE2["DOUBLE_BRACE"] = "bracePair";
  SHAPE_TYPE2["DOUBLE_BRACKET"] = "bracketPair";
  SHAPE_TYPE2["DOUBLE_WAVE"] = "doubleWave";
  SHAPE_TYPE2["DOWN_ARROW"] = "downArrow";
  SHAPE_TYPE2["DOWN_ARROW_CALLOUT"] = "downArrowCallout";
  SHAPE_TYPE2["DOWN_RIBBON"] = "ribbon";
  SHAPE_TYPE2["EXPLOSION1"] = "irregularSeal1";
  SHAPE_TYPE2["EXPLOSION2"] = "irregularSeal2";
  SHAPE_TYPE2["FLOWCHART_ALTERNATE_PROCESS"] = "flowChartAlternateProcess";
  SHAPE_TYPE2["FLOWCHART_CARD"] = "flowChartPunchedCard";
  SHAPE_TYPE2["FLOWCHART_COLLATE"] = "flowChartCollate";
  SHAPE_TYPE2["FLOWCHART_CONNECTOR"] = "flowChartConnector";
  SHAPE_TYPE2["FLOWCHART_DATA"] = "flowChartInputOutput";
  SHAPE_TYPE2["FLOWCHART_DECISION"] = "flowChartDecision";
  SHAPE_TYPE2["FLOWCHART_DELAY"] = "flowChartDelay";
  SHAPE_TYPE2["FLOWCHART_DIRECT_ACCESS_STORAGE"] = "flowChartMagneticDrum";
  SHAPE_TYPE2["FLOWCHART_DISPLAY"] = "flowChartDisplay";
  SHAPE_TYPE2["FLOWCHART_DOCUMENT"] = "flowChartDocument";
  SHAPE_TYPE2["FLOWCHART_EXTRACT"] = "flowChartExtract";
  SHAPE_TYPE2["FLOWCHART_INTERNAL_STORAGE"] = "flowChartInternalStorage";
  SHAPE_TYPE2["FLOWCHART_MAGNETIC_DISK"] = "flowChartMagneticDisk";
  SHAPE_TYPE2["FLOWCHART_MANUAL_INPUT"] = "flowChartManualInput";
  SHAPE_TYPE2["FLOWCHART_MANUAL_OPERATION"] = "flowChartManualOperation";
  SHAPE_TYPE2["FLOWCHART_MERGE"] = "flowChartMerge";
  SHAPE_TYPE2["FLOWCHART_MULTIDOCUMENT"] = "flowChartMultidocument";
  SHAPE_TYPE2["FLOWCHART_OFFLINE_STORAGE"] = "flowChartOfflineStorage";
  SHAPE_TYPE2["FLOWCHART_OFFPAGE_CONNECTOR"] = "flowChartOffpageConnector";
  SHAPE_TYPE2["FLOWCHART_OR"] = "flowChartOr";
  SHAPE_TYPE2["FLOWCHART_PREDEFINED_PROCESS"] = "flowChartPredefinedProcess";
  SHAPE_TYPE2["FLOWCHART_PREPARATION"] = "flowChartPreparation";
  SHAPE_TYPE2["FLOWCHART_PROCESS"] = "flowChartProcess";
  SHAPE_TYPE2["FLOWCHART_PUNCHED_TAPE"] = "flowChartPunchedTape";
  SHAPE_TYPE2["FLOWCHART_SEQUENTIAL_ACCESS_STORAGE"] = "flowChartMagneticTape";
  SHAPE_TYPE2["FLOWCHART_SORT"] = "flowChartSort";
  SHAPE_TYPE2["FLOWCHART_STORED_DATA"] = "flowChartOnlineStorage";
  SHAPE_TYPE2["FLOWCHART_SUMMING_JUNCTION"] = "flowChartSummingJunction";
  SHAPE_TYPE2["FLOWCHART_TERMINATOR"] = "flowChartTerminator";
  SHAPE_TYPE2["FOLDED_CORNER"] = "folderCorner";
  SHAPE_TYPE2["FRAME"] = "frame";
  SHAPE_TYPE2["FUNNEL"] = "funnel";
  SHAPE_TYPE2["GEAR_6"] = "gear6";
  SHAPE_TYPE2["GEAR_9"] = "gear9";
  SHAPE_TYPE2["HALF_FRAME"] = "halfFrame";
  SHAPE_TYPE2["HEART"] = "heart";
  SHAPE_TYPE2["HEPTAGON"] = "heptagon";
  SHAPE_TYPE2["HEXAGON"] = "hexagon";
  SHAPE_TYPE2["HORIZONTAL_SCROLL"] = "horizontalScroll";
  SHAPE_TYPE2["ISOSCELES_TRIANGLE"] = "triangle";
  SHAPE_TYPE2["LEFT_ARROW"] = "leftArrow";
  SHAPE_TYPE2["LEFT_ARROW_CALLOUT"] = "leftArrowCallout";
  SHAPE_TYPE2["LEFT_BRACE"] = "leftBrace";
  SHAPE_TYPE2["LEFT_BRACKET"] = "leftBracket";
  SHAPE_TYPE2["LEFT_CIRCULAR_ARROW"] = "leftCircularArrow";
  SHAPE_TYPE2["LEFT_RIGHT_ARROW"] = "leftRightArrow";
  SHAPE_TYPE2["LEFT_RIGHT_ARROW_CALLOUT"] = "leftRightArrowCallout";
  SHAPE_TYPE2["LEFT_RIGHT_CIRCULAR_ARROW"] = "leftRightCircularArrow";
  SHAPE_TYPE2["LEFT_RIGHT_RIBBON"] = "leftRightRibbon";
  SHAPE_TYPE2["LEFT_RIGHT_UP_ARROW"] = "leftRightUpArrow";
  SHAPE_TYPE2["LEFT_UP_ARROW"] = "leftUpArrow";
  SHAPE_TYPE2["LIGHTNING_BOLT"] = "lightningBolt";
  SHAPE_TYPE2["LINE_CALLOUT_1"] = "borderCallout1";
  SHAPE_TYPE2["LINE_CALLOUT_1_ACCENT_BAR"] = "accentCallout1";
  SHAPE_TYPE2["LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout1";
  SHAPE_TYPE2["LINE_CALLOUT_1_NO_BORDER"] = "callout1";
  SHAPE_TYPE2["LINE_CALLOUT_2"] = "borderCallout2";
  SHAPE_TYPE2["LINE_CALLOUT_2_ACCENT_BAR"] = "accentCallout2";
  SHAPE_TYPE2["LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout2";
  SHAPE_TYPE2["LINE_CALLOUT_2_NO_BORDER"] = "callout2";
  SHAPE_TYPE2["LINE_CALLOUT_3"] = "borderCallout3";
  SHAPE_TYPE2["LINE_CALLOUT_3_ACCENT_BAR"] = "accentCallout3";
  SHAPE_TYPE2["LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout3";
  SHAPE_TYPE2["LINE_CALLOUT_3_NO_BORDER"] = "callout3";
  SHAPE_TYPE2["LINE_CALLOUT_4"] = "borderCallout4";
  SHAPE_TYPE2["LINE_CALLOUT_4_ACCENT_BAR"] = "accentCallout3=4";
  SHAPE_TYPE2["LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR"] = "accentBorderCallout4";
  SHAPE_TYPE2["LINE_CALLOUT_4_NO_BORDER"] = "callout4";
  SHAPE_TYPE2["LINE"] = "line";
  SHAPE_TYPE2["LINE_INVERSE"] = "lineInv";
  SHAPE_TYPE2["MATH_DIVIDE"] = "mathDivide";
  SHAPE_TYPE2["MATH_EQUAL"] = "mathEqual";
  SHAPE_TYPE2["MATH_MINUS"] = "mathMinus";
  SHAPE_TYPE2["MATH_MULTIPLY"] = "mathMultiply";
  SHAPE_TYPE2["MATH_NOT_EQUAL"] = "mathNotEqual";
  SHAPE_TYPE2["MATH_PLUS"] = "mathPlus";
  SHAPE_TYPE2["MOON"] = "moon";
  SHAPE_TYPE2["NON_ISOSCELES_TRAPEZOID"] = "nonIsoscelesTrapezoid";
  SHAPE_TYPE2["NOTCHED_RIGHT_ARROW"] = "notchedRightArrow";
  SHAPE_TYPE2["NO_SYMBOL"] = "noSmoking";
  SHAPE_TYPE2["OCTAGON"] = "octagon";
  SHAPE_TYPE2["OVAL"] = "ellipse";
  SHAPE_TYPE2["OVAL_CALLOUT"] = "wedgeEllipseCallout";
  SHAPE_TYPE2["PARALLELOGRAM"] = "parallelogram";
  SHAPE_TYPE2["PENTAGON"] = "homePlate";
  SHAPE_TYPE2["PIE"] = "pie";
  SHAPE_TYPE2["PIE_WEDGE"] = "pieWedge";
  SHAPE_TYPE2["PLAQUE"] = "plaque";
  SHAPE_TYPE2["PLAQUE_TABS"] = "plaqueTabs";
  SHAPE_TYPE2["QUAD_ARROW"] = "quadArrow";
  SHAPE_TYPE2["QUAD_ARROW_CALLOUT"] = "quadArrowCallout";
  SHAPE_TYPE2["RECTANGLE"] = "rect";
  SHAPE_TYPE2["RECTANGULAR_CALLOUT"] = "wedgeRectCallout";
  SHAPE_TYPE2["REGULAR_PENTAGON"] = "pentagon";
  SHAPE_TYPE2["RIGHT_ARROW"] = "rightArrow";
  SHAPE_TYPE2["RIGHT_ARROW_CALLOUT"] = "rightArrowCallout";
  SHAPE_TYPE2["RIGHT_BRACE"] = "rightBrace";
  SHAPE_TYPE2["RIGHT_BRACKET"] = "rightBracket";
  SHAPE_TYPE2["RIGHT_TRIANGLE"] = "rtTriangle";
  SHAPE_TYPE2["ROUNDED_RECTANGLE"] = "roundRect";
  SHAPE_TYPE2["ROUNDED_RECTANGULAR_CALLOUT"] = "wedgeRoundRectCallout";
  SHAPE_TYPE2["ROUND_1_RECTANGLE"] = "round1Rect";
  SHAPE_TYPE2["ROUND_2_DIAG_RECTANGLE"] = "round2DiagRect";
  SHAPE_TYPE2["ROUND_2_SAME_RECTANGLE"] = "round2SameRect";
  SHAPE_TYPE2["SMILEY_FACE"] = "smileyFace";
  SHAPE_TYPE2["SNIP_1_RECTANGLE"] = "snip1Rect";
  SHAPE_TYPE2["SNIP_2_DIAG_RECTANGLE"] = "snip2DiagRect";
  SHAPE_TYPE2["SNIP_2_SAME_RECTANGLE"] = "snip2SameRect";
  SHAPE_TYPE2["SNIP_ROUND_RECTANGLE"] = "snipRoundRect";
  SHAPE_TYPE2["SQUARE_TABS"] = "squareTabs";
  SHAPE_TYPE2["STAR_10_POINT"] = "star10";
  SHAPE_TYPE2["STAR_12_POINT"] = "star12";
  SHAPE_TYPE2["STAR_16_POINT"] = "star16";
  SHAPE_TYPE2["STAR_24_POINT"] = "star24";
  SHAPE_TYPE2["STAR_32_POINT"] = "star32";
  SHAPE_TYPE2["STAR_4_POINT"] = "star4";
  SHAPE_TYPE2["STAR_5_POINT"] = "star5";
  SHAPE_TYPE2["STAR_6_POINT"] = "star6";
  SHAPE_TYPE2["STAR_7_POINT"] = "star7";
  SHAPE_TYPE2["STAR_8_POINT"] = "star8";
  SHAPE_TYPE2["STRIPED_RIGHT_ARROW"] = "stripedRightArrow";
  SHAPE_TYPE2["SUN"] = "sun";
  SHAPE_TYPE2["SWOOSH_ARROW"] = "swooshArrow";
  SHAPE_TYPE2["TEAR"] = "teardrop";
  SHAPE_TYPE2["TRAPEZOID"] = "trapezoid";
  SHAPE_TYPE2["UP_ARROW"] = "upArrow";
  SHAPE_TYPE2["UP_ARROW_CALLOUT"] = "upArrowCallout";
  SHAPE_TYPE2["UP_DOWN_ARROW"] = "upDownArrow";
  SHAPE_TYPE2["UP_DOWN_ARROW_CALLOUT"] = "upDownArrowCallout";
  SHAPE_TYPE2["UP_RIBBON"] = "ribbon2";
  SHAPE_TYPE2["U_TURN_ARROW"] = "uturnArrow";
  SHAPE_TYPE2["VERTICAL_SCROLL"] = "verticalScroll";
  SHAPE_TYPE2["WAVE"] = "wave";
  return SHAPE_TYPE2;
})(SHAPE_TYPE || {});
var CHART_TYPE = /* @__PURE__ */ ((CHART_TYPE2) => {
  CHART_TYPE2["AREA"] = "area";
  CHART_TYPE2["BAR"] = "bar";
  CHART_TYPE2["BAR3D"] = "bar3D";
  CHART_TYPE2["BUBBLE"] = "bubble";
  CHART_TYPE2["BUBBLE3D"] = "bubble3D";
  CHART_TYPE2["DOUGHNUT"] = "doughnut";
  CHART_TYPE2["LINE"] = "line";
  CHART_TYPE2["PIE"] = "pie";
  CHART_TYPE2["RADAR"] = "radar";
  CHART_TYPE2["SCATTER"] = "scatter";
  return CHART_TYPE2;
})(CHART_TYPE || {});
var SCHEME_COLOR_NAMES = /* @__PURE__ */ ((SCHEME_COLOR_NAMES2) => {
  SCHEME_COLOR_NAMES2["TEXT1"] = "tx1";
  SCHEME_COLOR_NAMES2["TEXT2"] = "tx2";
  SCHEME_COLOR_NAMES2["BACKGROUND1"] = "bg1";
  SCHEME_COLOR_NAMES2["BACKGROUND2"] = "bg2";
  SCHEME_COLOR_NAMES2["ACCENT1"] = "accent1";
  SCHEME_COLOR_NAMES2["ACCENT2"] = "accent2";
  SCHEME_COLOR_NAMES2["ACCENT3"] = "accent3";
  SCHEME_COLOR_NAMES2["ACCENT4"] = "accent4";
  SCHEME_COLOR_NAMES2["ACCENT5"] = "accent5";
  SCHEME_COLOR_NAMES2["ACCENT6"] = "accent6";
  return SCHEME_COLOR_NAMES2;
})(SCHEME_COLOR_NAMES || {});
var MASTER_OBJECTS = /* @__PURE__ */ ((MASTER_OBJECTS2) => {
  MASTER_OBJECTS2["chart"] = "chart";
  MASTER_OBJECTS2["image"] = "image";
  MASTER_OBJECTS2["line"] = "line";
  MASTER_OBJECTS2["rect"] = "rect";
  MASTER_OBJECTS2["text"] = "text";
  MASTER_OBJECTS2["placeholder"] = "placeholder";
  return MASTER_OBJECTS2;
})(MASTER_OBJECTS || {});
var PLACEHOLDER_TYPES = /* @__PURE__ */ ((PLACEHOLDER_TYPES2) => {
  PLACEHOLDER_TYPES2["title"] = "title";
  PLACEHOLDER_TYPES2["body"] = "body";
  PLACEHOLDER_TYPES2["image"] = "pic";
  PLACEHOLDER_TYPES2["chart"] = "chart";
  PLACEHOLDER_TYPES2["table"] = "tbl";
  PLACEHOLDER_TYPES2["media"] = "media";
  return PLACEHOLDER_TYPES2;
})(PLACEHOLDER_TYPES || {});
var IMG_BROKEN = "data:image/png;base64,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";
var IMG_PLAYBTN = "data:image/png;base64,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AAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKSEAhgAAAAAAAAgJRTAAAAAAAAAACmhAAYAAAAAAABICQUwAAAAAAAAQEoogAEAAAAAAABSQgEMAAAAAAAAkBIKYAAAAAAAAICUUAADAAAAAAAApIQCGAAAAAAAACAlFMAAAAAAAAAAKaEABgAAAAAAAEgJBTAAAAAAAABASiiAAQAAAAAAAFJCAQwAAAAAAACQEgpgAAAAAAAAgJRQAAMAAAAAAACkhAIYAAAAAAAAICUUwAAAAAAAAAApoQAGAAAAAAAASAkFMAAAAAAAAEBKKIABAAAAAAAAUkIBDAAAAAAAAJASCmAAAAAAAACAlFAAAwAAAAAAAKREkRUAAACwrUpLSwuGDRvWfMCAAS26du3avKysrLiioqKkZcuWzZs1a1bcvHnz0tLS0rJsNtusuLi4ebNmzUoLCgo+8/eijY2N9Zs3b66pra2tqqur21xTU1NdVVVVs2nTptqNGzdWbdiwoeYvf/nL5hUrVlQtWLBgw6xZs6pqamoaJQYAAEDaKYABAACIiIghQ4aUHnTQQW379u3bql27dq3at2/fpkWLFq2bN29eWVpa2qpZs2bNCwsLm2ez2fLCwsLyoqKi8sLCwtKknK+hoaG6vr6+qqGh4S91dXV/aWhoqNq8eXNVTU3NuqqqqvUbNmxYu2rVqjWrV69e99Zbb6177rnnPpgzZ06NTwYAAABJogAGAADIA8OGDWt+xBFHdBwwYECnLl26dGjdunXHFi1adCgtLe1YUlLSvlmzZq0KCgqK07yDwsLCssLCwrKIaPdp/zuNjY21mzdvXrdp06ZVNTU172/YsGHl2rVr31+2bNnKBQsWrHjyySffnzVrVpVPGAAAAE1Fpuexsd9HfaF+ZcSal0ptCAAAIAE6deqUPf744zvtueeeXbp3796lbdu2XSorKzuXlpZ2KS0t7VBYWFhhSztGQ0PDxpqampU1NTXL169fv+yDDz5Y9s477yybPXv2sj/96U8rVqxYUWdLAAAAbE9t9q6Jog4f/TUFMAAAQEJks9nMt7/97Y4jRozo1bdv397t2rXrXl5e3rWsrKxzcXFx+4gosKUmp7G2tnZVTU3Nso0bNy5btWrV0tdff/2tJ598cvG999672noAAADYFgpgAACAhPne977X6a9Fb/v27Xu1bNmyV1lZWa8kvXOXLauvr9/wl7/8ZdG6desWL1u2bNHChQsX/fGPf1w8derUjbYDAADAliiAAQAAmqhsNps59dRTuxx66KH9+/Tp87n27dv3Ly8v719UVOSRzXlq06ZNKzZu3Pj6+++//8abb775xqOPPvrG3XffvcpmAAAA+CsFMAAAQBNx6qmndvniF784qHfv3v3btWv3uYqKis8VFhaW2wxbUl9fv37Dhg1vfPDBB68vXrz4jccee2z+jTfeuNxmAAAA8pMCGAAAYBc45phjWn/rW9/aq3///kPatGnTv6Kiop9HOLO9NDQ0VG/cuPGtNWvWLFy4cOGcO+6445WHHnporc0AAACknwIYAABgJzjjjDO6f+lLX9qrV69eg1u3bj2orKysR0RkbIadJFddXb103bp18xcvXjz30UcffeXqq69+x1oAAADSRwEMAACwnZWWlhb86le/2u3QQw8d1r17931btmw5qLCwsMxmaEoaGhqqP/zww/nvvPPOzGeeeWbW2LFj36ipqWm0GQAAgGRTAAMAAGwHP/7xj7t+9atf3bdXr15D27Ztu1c2m21jKyRJXV3dmg8++OCVRYsWvfznP/95xh/+8IdltgIAAJA8CmAAAIBtcOKJJ7Y75ZRTDujXr9+w1q1bD81ms61shTSpq6tbt3bt2pfffPPNWbfccsvUe++9d7WtAAAANH0KYAAAgE+hoqKi4IILLhg0YsSI/bp27bpfy5YtB2YymUKbIR/kcrmGDz/8cP6777474/nnn59x4YUXvrZx40aPiwYAAGiCFMAAAAAf4/jjj2/7/e9//8D+/fsf2Lp1630KCgpKbAUiGhsbN61fv37eW2+9NeWGG2545u67715lKwAAAE2DAhgAAOB/ZLPZzAUXXPC5I4888sDu3bsfWFFRsVtEFNgMbFl1dfWSd999d8qsWbNmnnvuuS+vW7euwVYAAAB2DQUwAACQ10pLSwsuvfTSQYcccsjBXbt2HVFWVtbDVmDb1dbWrnr//fdfmDp16uRf/vKXL65evbreVgAAAHYeBTAAAJB3Bg0aVHrBBRd8fs899zywQ4cOBxQVFbWwFdj+Ghsba9euXTtrzpw5T59//vmTX3755WpbAQAA2LEUwAAAQF4YNmxY8/POO+/gIUOGHOZ9vrDz/W0ZfNFFFz07a9asKlsBAADY/hTAAABAarVq1arwyiuv3HfEiBEjO3TocFBhYWGZrcCu19DQUP3+++8/O2XKlIk/+clPZm7cuLHRVgAAALYPBTAAAJAqrVq1Kvztb3+7/3777Xd4x44dRxQWFpbbCjRdDQ0NG99///0pM2bMeOqHP/zhC8pgAACAz0YBDAAApMJZZ53V45vf/OaRvXr1GllaWtrVRiB5ampq3l28ePHEO++8c9LVV1/9jo0AAABsPQUwAACQWMOHDy+/6KKLvjB48OCjW7RoMdBGID0+/PDDV+fNmzfhvPPOe3L69Ol/sREAAIBPRwEMAAAkSqtWrQpvuOGGQ/bbb79/atOmzX6ZTCZrK5BeuVyubs2aNTNmzJjx2JgxYyavW7euwVYAAAA+ngIYAABIhB//+Mddv/e9732lZ8+e/1RcXNzWRiD/1NbWfvD2228/dssttzz029/+9l0bAQAA+L8UwAAAQJNVUVFRcO21137+4IMPPrZ169b7ZTKZAlsBIqJxzZo1M59//vnxp5122hR3BQMAAPx/CmAAAKDJOeWUUzqefvrpx/bu3ftL2Wy2jY0AH6e+vn7j0qVLH/vd7373x+uvv36ZjQAAAPlOAQwAADQJ2Ww2c+uttx5wyCGHnNC6deu9I8LdvsDWaFy7du1L06ZN+/OPfvSjZ1evXl1vJQAAQD5SAAMAALtU//79S6655pp/2nPPPY8tLy/vayPAZ1VTU7NswYIF488999wHp06dutFGAACAfKIABgAAdomf//znPU855ZQTu3btemRhYWGZjQDbW2NjY92KFSuevOWWW+689NJLF9kIAACQDxTAAADATuMxz8Cusn79+rlPP/30f5188slT6+rqcjYCAACklQIYAADY4fr27Vv8hz/84a+Pee5nI8CuUlNT8+68efPu/8EPfvDgwoULN9kIAACQNgpgAABghxkyZEjpNddc89XBgwefWFxc3MFGgKaitrZ21dy5c+/5yU9+8uc5c+bU2AgAAJAWWyqAPYoNAADYJqNHj+4wb968n06ZMuXRYcOGnaH8BZqa4uLi9sOGDTtjypQpj86bN++nJ510UntbAQAA0s4dwAAAwFY599xze33/+9//dufOnY/IZDJZGwGSIpfL1S1fvvzJG2644fbLLrvsbRsBAACSyiOgAQCAz+y8887r+53vfOfbHTt2PDyTyRTaCJBUuVyuYcWKFU/cdNNN//XrX/96sY0AAABJowAGAAC22WWXXTboG9/4xg9at249zDaAtFm7du2su++++9pzzjnnNdsAAACSQgEMAABsNcUvkE8UwQAAQJIogAEAgE9N8Qvks7Vr18665557rvv5z38+3zYAAICmaksFcGHlwOj6UV9orIqoWZG1PQAAyBO/+MUvet9xxx3nHHrooT8pLS3tYiNAPiotLe2y7777HvP973+/X1lZ2ZIpU6assxUAAKCpKetcHwXlH/01BTAAAOS5M844o/u99957zpe//OWflZeX94qIjK0AeS5TXl7e8+CDDx71/e9/v3dEvDVjxowPrQUAAGgqFMAAAMD/ceKJJ7a77777fjJq1Kh/KS8v7xOKX4B/lCkvL+99+OGHj/rWt77VfvXq1Qvnz59fbS0AAMCutqUC2DuAAQAgzwwdOrTs+uuvP6l///4nFRYWltkI20NjY2Ns2rQpqquro6amJurr62PTpk2xefPmqK+vj+rq6qivr4/NmzfHpk2boqGhYZv/fxUWFkZJSUk0a9YsioqKoqysLIqKiqJZs2ZRUlISRUVFUVpa+r9/FRQUCIjtoqGhoeq11167a8yYMffMmTOnxkYAAIBdZUvvAFYAAwBAnujUqVP2nnvuGbXXXnudnM1mK22Ej9PQ0BAbN26MDRs2/J+/Nm7cGBs3boyamprYtGlTbNq0KWpqaqK2trbJnqe4uDhKSkqitLT0f/9eUVERFRUV0aJFi//zV0VFRRQWFvog8LHq6urWvvjii7eceOKJf169enW9jQAAADubAhgAAPLcXXfdddAXv/jF00tLS7vZRn7L5XKxYcOGWLt2baxbty7Wrl37d3+tW7cuNmzYkPd7atGiRbRu3TpatWoVrVu3jjZt2vzvf27dunW0aNHCh4morq5e+sgjj1zzne98Z6ptAAAAO5MCGAAA8tTVV189+MQTTzyzoqJioG3kj8bGxli5cmUsX748Pvjgg1i9evX//n3t2rXR2NhoSZ9RYWFhtGrVKtq1axdt27b937937tw5OnTo4LHTeWbDhg3z77333qvOPPPMebYBAADsDApgAADIM1/72tfaXHrppad27979qIjQRKVUQ0NDrFq1KlasWBHvv//+//595cqVTfqRzGlXXFwcHTp0iI4dO0bnzp2jY8eO0alTp2jXrp1HS6dYLpdrfOeddx76+c9/fv2ECRPW2QgAALAjKYABACBP9OrVq9ldd931jT322OM7hYWFZTaSHh9++GG88847sXTp0njvvfdixYoVsXr16mhoaLCchCgsLIz27dtHp06dolu3btG9e/fo3r27x0mnTENDQ9W8efNu++Y3v/nHJUuWbLYRAABgR1AAAwBAHrjrrrtG/NM//dOZJSUlXWwj2davXx9Lly6Nd955539L3w8//NBiUqqysvJ/y+C//tWqVSuLSbiamppljz322G9Gjx49xTYAAIDtTQEMAAAp9qtf/arPD3/4w5+1atVqL9tIno0bN8aSJUvirbfeikWLFsV7770XmzZtspg8V1JSEl27do0+ffpE3759o3fv3lFeXm4xCbRu3bqXr7322ivGjh27yDYAAIDtRQEMAAApNGjQoNI77rjju7vttttJBQUFWRtJhtWrV8ebb74ZixcvjiVLlsTy5cujsbHRYtiigoKC6Ny5c/Tu3Tt69+4d/fr1i7Zt21pMQjQ2Nta98cYbd33rW9+6ff78+TU2AgAAfFYKYAAASJHS0tKCBx988Jj99tvvn7PZbBsbaboaGhri7bffjrfeeisWLFgQS5YscXcv201FRUX06tUr+vbtG3379o2ePXtGYWGhxTRhdXV1a2bMmHHjV77ylYdqamr85gcAALDNFMAAAJASp59+erdf/vKX51ZWVu5jG03T6tWr47XXXouFCxfGm2++GRs3brQUdooWLVpE3759Y8CAATFw4EB3CDdh69evf/E//uM//vPqq69+xzYAAIBtoQAGAICEGzRoUOm99977w969ex+byWTc4teErF+/PubNmxcLFiyIN954Q+FLk9GiRYvo169fDBgwIPbYY4+orKy0lCYkl8s1LF68eNyJJ554rcdCAwAAW0sBDAAACXbNNdcMOemkk35RVlbWyzZ2vVwuF++++27MnTs3XnvttViyZIl3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B7lcrm7y5Mm//vKXv/yIbQAAwK63pXcAK4ABAMg7paWlBTNnzjyzT58+x9vG39uwYUOMGzcuZsyY4VHPQF7KZDKx3377xde//vWoqKiwkH+waNGiP+27775X1dTUeCwEAADsQgpgAAD4H926dctOnjz5V506dRppG/9fLpeLqVOnxp///OfYuHGjhQB5r6KiIkaNGhX777+/x0L/g+XLlz9+6KGHXvLuu+/W2QYAAOwaWyqAvQMYAIC8MXz48PInnnjiynbt2o2wjf/vnXfeiWuvvTaee+652Lx5s4UARMTmzZtjzpw58dprr0XPnj2jRYsWlvI/Kioq+n7rW98aMnXq1Ofee+89f3AAAMAusKV3ACuAAQDIC9/+9rc73n777X9o0aLFANv4b1VVVXHXXXfFvffeG+vXr7cQgI+wbt26eP7552P9+vWx2267RVFRkaVERElJSefjjjvuoA8++GDKK6+88hcbAQCAnUsBDABAXjv//PP7XXzxxX8oKSnpbBv/bfr06XHttdfGokWLLAPgU3jnnXdi2rRp0bp16+jc2R8nERHZbLbyC1/4whElJSUvTp48eY2NAADAzqMABgAgb/3ud7/b60c/+tFVRUVFrWwjYs2aNXHzzTfHpEmTora21kIAtkJtbW289NJL8c4770Tfvn2jtLQ073dSWFhYNnz48C/26dNn4UMPPbTMpwQAAHYOBTAAAHnp1ltv3f+b3/zmfxYWFjbP913kcrl4/vnn4/rrr4/ly5f7cAB8BitXroxp06ZFRUVFdOvWLTKZTF7vo6CgIDto0KBDBw0atOiBBx54xycEAAB2vC0VwJmex8Z+H/WF+pURa17ym6wAACTTww8//KXDDjvsXzKZTN6/rPGDDz6I22+/Pd544w0fDIDtbMCAAfGtb30r2rRpk/e7yOVyjVOmTPn1yJEjH/LJAACAHavN3jVR1OGjv6YABgAgdV555ZXTPve5z30r3/fQ0NAQjz32WDz++ONRV1fngwGwg2Sz2Tj66KPjC1/4QhQUFOT9Pl5//fU79tprr9/7ZAAAwI6jAAYAIC9ks9nMyy+/fFafPn2Oz/ddvPvuu3HbbbfFe++954MBsJN069YtvvOd70S3bt3yfhdLliy5f5999rmypqam0ScDAAC2PwUwAACpV1paWjBr1qyzevfufVw+7yGXy8WTTz4ZDz74oLt+AXaBbDYbxxxzTBxxxBF5fzfw0qVLHxg6dOjlSmAAANj+FMAAAKRar169mk2ePHlsu3btDsrnPaxcuTJuueWWePvtt30oAHaxnj17ximnnBIdOnTI6z2sXr16yiGHHPIvS5Ys2exTAQAA28+WCuDCyoHR9aO+0FgVUbMia3sAADRpQ4cOLXvqqacub9Omzf75uoNcLhfPPPNMXH/99bF27VofCoAmYP369TFlypQoKSmJnj17RiaTycs9NG/evPtJJ500ZPLkyc+sWLHCoykAAGA7KetcHwXlH/01BTAAAIk1ZMiQ0kceeeSKVq1a7Z2vO6iuro7bb789nnjiiWhs9IRNgKaksbEx5s+fH++//34MGDAgstn8/DlLaWlpp6997WuDn3rqqadXrlxZ75MBAACfnQIYAIDUOfTQQ1s8+OCDv2/ZsuUe+bqDOXPmxNVXX+2RzwBN3PLly+OFF16Ijh075u0joUtLSzudcMIJ+7/00ktPv/3227U+FQAA8NkogAEASJVhw4Y1v++++37TsmXLQfl4/vr6+hg/fnz88Y9/jNpaP0MHSILNmzfHiy++GJs3b47ddtstCgoK8m4HxcXFbY866qg9n3vuuaeXL1/ucdAAAPAZKIABAEiNI488snLcuHG/b9GixcB8PP97770XV111VcyZM8eHASCBFi1aFC+//HL069cvWrRokXfnLykp6XDcccftP2fOnGcWLVq0yScCAAC2jQIYAIBUOPLIIyvvvPPO35aXl++Wj+d/+umn48Ybb4wPP/zQhwEgwf7yl7/ECy+8ECUlJdGrV6+8O3+zZs3aHHXUUfspgQEAYNspgAEASLxjjz227W233faH5s2b98m3s1dVVcXNN98cTz31VDQ2NvowAKRAY2NjzJ8/P5YtWxYDBgyIZs2a5dX5mzVr1uaYY4458M0333xm4cKFNT4RAACwdRTAAAAk2qGHHtritttuuzofy9+33347rrnmmli8eLEPAkAKvf/++/HKK69Enz59orKyMq/Ons1mK4888sh9Zs6c+dTSpUs3+zQAAMCnpwAGACCxjjjiiJb33nvvteXl5f3y6dy5XC4mTZoUN998c1RVVfkgAKRYVVVVTJ06NbLZbPTp0ycymUzenL24uLjtV7/61c+/8sorTy1evLjWpwEAAD4dBTAAAIl06KGHtrj33nt/l2/lb3V1ddx0000xefLkyOVyPggAeSCXy8WCBQvi3Xffjd133z2y2fz5mUyzZs1aH3300fvNmDHjSXcCAwDAp6MABgAgcYYOHVo2fvz4qysqKgbk07mXLVsWV111lUc+A+SplStXxiuvvBKf+9znoqKiIm/O3axZszZHH3300GeeeebJFStW1PkkAADAlimAAQBIlCFDhpQ++uij17Rs2XL3fDr31KlT49prr42NGzf6EADksaqqqpg+fXq0bds2unTpkjfnLikpaT9q1KihTz755JMrV66s90kAAICPt6UCuMB6AABoSjp16pSdMGHCv1dWVu6RL2dubGyMcePGxR133BF1dW56AiCitrY2br755hg/fnw0NjbmzbkrKyv3mDBhwr9369bNXQkAALCNFMAAADQZrVq1Kpw+ffolbdq02T9fzlxdXR2/+93vYtKkSd73C8DfyeVy8fjjj8fvf//7qK6uzptzt2nTZv8pU6Zc0qpVq0KfAgAA2HoKYAAAmoSKioqC2bNnX9KuXbuD8uXMS5cujYsuuijmz5/vAwDAx3r11VfjoosuiqVLl+bNmdu1a3fQ7Nmz/72iosLPrgAAYCu5iAYAoEmYOXPmz9q1a3dIvpz35ZdfjiuuuCLWrVsnfAA+0bp16+KKK66Il19+OW/O3K5du4Nnzpz5M+kDAMDWUQADALDLvfjii2N69OgxKh/Omsvl4oEHHogbbrghamtrhQ/Ap1ZbWxs33HBDPPDAA3nz2oAePXqMevHFF8dIHwAAPj0FMAAAu9SkSZO+NnDgwFPy4ax1dXVx8803x8SJE73vF4BtksvlYuLEiXHLLbdEXV1dXpx54MCBJ0+aNOlr0gcAgE9HAQwAwC7z6KOPHnXggQeekw9nXbduXfz617+OWbNmCR6Az2zmzJnx61//Ol9eJZA58MADz3n00UePkjwAAHyywsqB0fWjvtBYFVGzImtDAADsEDfeeOO+Rx999EWZTKYw7Wddvnx5XHXVVbFy5UrBA7DdbNiwIWbPnh0DBw6MioqKtB8307179/179uz56sMPP7xc+gAA5LuyzvVRUP7RX1MAAwCw011xxRV7fPe7372qoKCgWdrPOmfOnPjtb38bGzduFDwA2111dXVMmzYtOnfuHB07dkz1WTOZTOHuu+9+eJs2bV6aNGnSKukDAJDPFMAAADQZZ5xxRvef/exnvy0sLCxP+1knTJgQd999d9TX1wsegB2moaEhXnrppchms9G3b99UnzWTyRTttddeB/3lL395dubMmRukDwBAvlIAAwDQJBx00EEVf/jDH64pLi7ulOZz5nK5eOCBB+Kxxx4TOgA77c+eBQsWRF1dXfTv3z8ymUxqz1pQUFBywAEHDJs+ffqkpUuXbpY+AAD5aEsFcIH1AACwMwwaNKj0vvvuu7qsrKxXms9ZV1cX1113XUyaNEnoAOx0EydOjOuvvz7q6upSfc6ysrJef/rTn67u379/idQBAODvKYABANjhKioqCh577LGLKyoqBqb5nNXV1XHNNdfE7NmzhQ7ALvPKK6/ElVdeGVVVVak+Z4sWLQZOnDhxbEVFhZ9vAQDA33CBDADADjdz5syftW3b9sA0n3HdunVx2WWXxRtvvCFwAHa5xYsXx2WXXRZr165N9TnbtWt34MyZM38mcQAA+P8UwAAA7FBPPvnkqB49eoxK8xlXrVoVV1xxRSxfvlzgADQZK1asiCuuuCJWrlyZ6nP26NFj1KRJk0ZJHAAA/lth5cDo+lFfaKyKqFmRtSEAALbZjTfeuO+XvvSlCzOZTGp/8fDdd9+NK6+8MtatWydwAJqc6urqmDVrVvTv3z8qKytTe85u3boN79mz57yHH37Yb2MBAJAXyjrXR0H5R39NAQwAwA5x3nnn9T311FOvLigoKE7rGV977bW45pprorq6WuAANFmbN2+OGTNmRI8ePaJ9+/apPGMmkykYNGjQIYWFhVOee+45v5UFAEDqKYABANipjjrqqFb/8R//8YdmzZq1SusZX3755bj++uujrq5O4AA0eQ0NDfHSSy9Fp06dolOnTqk8Y0FBQXbYsGGfnz9//qQ33nhjk9QBAEizLRXA3gEMAMB21a1bt+wNN9zwnyUlJR3TesYpU6bEjTfeGPX19QIHIDHq6+vjxhtvjKlTp6b2jCUlJZ1uuOGG/+jWrZu7GgAAyFsKYAAAtqunn376XyorK/dI6/kmTZoUd955ZzQ2NgobgMRpbGyMO+64I5588snUnrGysnLw008//UtpAwCQrxTAAABsN88///w3unTp8k9pPd/EiRNj3LhxkcvlhA1AYuVyubj//vtTXQJ36dLlS88+++yJ0gYAIB95BzAAANvFTTfdNPzII488L5PJZNJ4vsceeyzGjx8vaABS47XXXotmzZpF3759U3m+zp0779urV695Dz/88DJpAwCQNlt6B7ACGACAz+wXv/hF7x/+8IdXFxQUNEvj+R544IF45JFHBA1A6ixYsCDq6upiwIABqTtbJpPJDBo06ODGxsbnpk6dul7aAACkiQIYAIAd5oADDqj43e9+99tmzZq1TeP5xo0bF5MmTRI0AKm1aNGi2Lx5cwwcODB1ZysoKMjut99+w5577rnH33vvvc3SBgAgLbZUAHsHMAAA2yybzWbuvPPOfyktLe2exvNNmDBB+QtAXpg0aVI89NBDqTxbaWlpj3vuuedfstlsRtIAAOQDBTAAANvs+eef/06HDh0OTePZHn744Xj44YeFDEDeeOSRR+LPf/5zKs/WoUOHw5599tlvSxkAgHygAAYAYJvcd999hw8ePPjUNJ7t/vvvjwkTJggZgLzz2GOPxX333ZfKs+25554/+NOf/nSYlAEASDvvAAYAYKudccYZ3ceMGXN5QUFBcdrONnHixHjkkUeEDEDeWrx4cWSz2ejbt2/ajpbp06fPvn/5y18mz5w5c4OkAQBIsi29A1gBDADAVhk2bFjzG2+88Q/NmjVrl7azPfroo6l99CUAbI2FCxdGUVFR9OvXL1XnKigoKD7wwAP3e/LJJx9dsWJFnaQBAEiqLRXAHgENAMBWuffee39ZWlraPW3nevzxx+PBBx8UMAD8jz//+c8xceLE1J2rtLS0x3333fdLCQMAkFYKYAAAPrVJkyaN6tSp0xEpPFeMHz9ewADwD8aPHx+TJ09O3bk6der0hUmTJn1VwgAApJFHQAMA8Kmcd955fU888cR/z2QyRWk618yZM+Puu+8WMAB8jNdeey06duwYnTt3TtW5unbtuk9BQcHzzz333DopAwCQNN4BDADAZ3LEEUe0vOKKK67NZrOVaTrXyy+/HDfffHPkcjkhA8DHyOVyMXv27OjSpUt06tQpNefKZDJF++yzz/CpU6c+9u67726WNAAASeIdwAAAbLNsNpu55ZZb/q2kpKRjms61YMGCuPnmm6OxsVHIAPAJGhsb4+abb44333wzVecqLS3tcvfdd5+fzWYzUgYAIC0UwAAAbNGkSZO+3rZt2wPTdKZly5bFDTfcEPX19QIGgE+prq4urr322li+fHmqztWuXbsDH3/88VESBgAgLTwCGgCAj3XZZZcN+upXvzo2k8mk5hcH33///bjyyiujqqpKwACwlerq6uLll1+OIUOGRHl5eWrO1aVLl31LS0unPvPMM2ukDABAEngENAAAW61///4lJ5988q8ymUxRWs60YcOG+P3vfx8bN24UMABso40bN8bvfve7VP15WlBQkP3hD394ft++fYslDABA4q9vrQAAgI/y4IMPnl1WVtYrLeeprq6O3/zmN7Fq1SrhAsBntGrVqrjyyiujuro6NWcqKyvr8/DDD58lXQAAkk4BDADA/zF+/Pgju3XrdnRazlNfX5/KdxYCwK60fPnyuO6666K+vj41Z+rRo8dXx40bd4R0AQBIMgUwAAB/53vf+16nI4444py0nCeXy8Vtt90Wb7zxhnABYDt7/fXX47bbbotcLpeaMx155JHnfvvb3+4oXQAAkkoBDADA/6qoqCi4+OKLLywsLCxPy5nGjx8fs2bNEi4A7CCzZs2Khx56KDXnKSwsrPj1r399QUVFhZ+bAQCQSC5kAQD4XxMnThxdWVk5OC3nef7552PixImCBYAd7LHHHosXXnghNeeprKzc89FHHz1RsgAAJFFh5cDo+lFfaKyKqFmRtSEAgDxxwQUX7DZq1KgLM5lMYRrO8+qrr8Ytt9ySqkdSAkBT/7O3d+/e0a5du1Scp2PHjkNzudxzU6ZMWSddAACamrLO9VHwMc/wcwcwAADRt2/f4h//+McXZzKZVPwG4HvvvRc33HBDNDY2ChcAdpKGhoa47rrrYtmyZak4T0FBQfbss88e27dv32LpAgCQqGtZKwAAYPz48T8qKyvrkYazbNiwIX7/+99HbW2tYAFgJ9u0aVP8/ve/j40bN6biPGVlZb3GjRs3RrIAACSJAhgAIM/ddNNNw/v06XN8Gs5SX18f1157baxdu1awALCLrFmzJq699tqor69PxXn69ev3jd///vdDJQsAQFIogAEA8thBBx1Uceyxx/5rRGTScJ477rgjFi9eLFgA2MUWLVoUd955Z1qOU/CNb3zj34YNG9ZcsgAAJOIC1goAAPLXzTfffFZxcXG7NJxl4sSJMX36dKECQBMxbdq0mDRpUirOUlJS0unOO+88Q6oAACSBAhgAIE/913/914FdunT5UhrO8tprr8Wf//xnoQJAEzN+/PhYsGBBKs7SrVu3o2+66abhUgUAoKlTAAMA5KEvfelLlV/5yld+lYazrFixIq6//vpobGwULAA0MY2NjXHdddfFihUr0nCczHHHHfergw46qEKyAAA0ZQpgAIA8dPXVV5+ezWYrk36OmpqauPbaa2PTpk1CBYAmatOmTXHttddGTU1N4s+SzWbb3njjjT+RKgAATZkCGAAgz9x6663Du3Tp8uWknyOXy8Utt9wSK1euFCoANHErV66MW2+9NXK5XOLP4lHQAAA0dQpgAIA8MnTo0LKvfvWrv0jDWSZMmBBz584VKgAkxJw5c+Kxxx5LxVlGjRr1i6FDh5ZJFQCApkgBDACQR+64444fFRcXd0z6OV5++eV45JFHBAoACfPQQw+l4he4SkpKOt5xxx0/lCgAAE2RAhgAIE9cfvnlu/fs2XNU0s/xwQcfxB133JGKR0gCQL7J5XJx2223xZo1axJ/lp49ex57+eWX7y5VAACaGgUwAEAe6NatW/a73/3uv2YymURf/9XX18cNN9wQ1dXVQgWAhKqqqoobb7wx6uvrE32OTCZT8N3vfvdX3bp1y0oVAICmRAEMAJAHxo8ff0pZWVmvpJ/jnnvuiaVLlwoUABJuyZIlcd999yX+HGVlZT3Hjx9/ikQBAGhKFMAAACn385//vOeAAQNGJ/0c06dPjylTpggUAFJi8uTJMWPGjMSfY8CAAaN//vOf95QoAABNhQIYACDFstls5qyzzjo3k8kk+tGEK1asiLvvvlugAJAyd911V6xYsSLRZ8hkMtmzzjrr3Gw2m5EoAABNgQIYACDFxo0b98XKysq9knyG2trauOGGG6K2tlagAJAyf/1zfvPmzYk+R2Vl5V7jxo0bKVEAAJoCBTAAQEoNHz68/OCDDz4t6ee4//77Y/ny5QIFgJRavnx5jBs3LvHnGDFixI+HDRvWXKIAAOxqCmAAgJS69dZbT8tms22TfIYZM2bEc889J0wASLnJkyfHzJkzE32G4uLitrfffvtp0gQAYFdTAAMApNBVV121R48ePb6S5DOsXLky7rrrLmECQJ64++6744MPPkj0GXr27PnVK664Yg9pAgCwKymAAQBSprS0tOAb3/jGT5N8rdfY2Bi333679/4CQB6pqamJ2267LRobG5N8jIJvfvObZ5aWlvqZGwAAu+6i1AoAANJlwoQJX6uoqBiQ5DOMHz8+Fi1aJEwAyDNvvvlmPPjgg4k+Q4sWLQY9+OCDx0gTAIBdRQEMAJAiRx55ZOWwYcN+kOQzzJ07N5544glhAkCemjhxYixYsCDRZxg+fPiPjjjiiJbSBABgV1AAAwCkyBVXXHFyUVFRRVLnr6qqijvvvDNyuZwwASBP5XK5uP3226O6ujqxZygqKmrxm9/85mRpAgCwKyiAAQBS4vzzz+/Xu3fv45J8httvvz0+/PBDYQJAnlu3bl3cfvvtiT5D7969jz///PP7SRMAgJ1NAQwAkALZbDZz6qmn/jyTyST2+m769OkxZ84cYQIAERExe/bsmDFjRmLnz2QyBaeeeurPs9lsRpoAAOxMCmAAgBT44x//eERlZeXgpM6/du3auPfeewUJAPyde+65J9atW5fY+SsrKwf/6U9/+oIkAQDYmRTAAAAJ17dv3+JDDjnkR0k+w9133x01NTXCBAD+Tk1NTdx9992JPsPBBx/8o759+xZLEwCAnUUBDACQcHfdddc3S0pKOiV1/smTJ8e8efMECQB8pLlz58azzz6b2PlLSko63nPPPd+SJAAAO4sCGAAgwb70pS9VDhw48KSkzr9mzZoYP368IAGALXrggQdizZo1iZ2/f//+Jx111FGtJAkAwM6gAAYASLArrrji1MLCwvIkzp7L5eK2226LTZs2CRIA2KJNmzbFbbfdFrlcLpHzFxYWll1++eU/kCQAADuDAhgAIKF+8Ytf9O7evftXkjr/s88+G2+88YYgAYBP5Y033ojnn38+sfN369bt6F/96ld9JAkAwI6mAAYASKgf/vCHP8pkMom8nvvggw/igQceECIAsFXGjRsX69atS+TsmUym4NRTT/2xFAEA2NEUwAAACXTdddcNa9eu3YFJnD2Xy8Udd9wRtbW1ggQAtsqmTZvizjvvTOz8bdq02f+mm27aT5IAAOxICmAAgIQpLS0t+NrXvnZ6Uud/4YUXYuHChYIEALbJq6++GjNmzEjs/Mccc8zpFRUVfiYHAMAO42ITACBhbr/99oMrKip2S+LsGzZsiHHjxgkRAPhM7r///qiqqkrk7OXl5X3/67/+6wgpAgCwoyiAAQASpKKiouCwww47Nanz33vvvYn9YS0A0HRs2LAh7r///sTOf9BBB/1zq1atCiUJAMCOoAAGAEiQ+++//+iysrKeSZx9zpw58dJLLwkRANguXnjhhViwYEEiZy8tLe32xz/+8StSBABgR1AAAwAkRN++fYv33Xfff07i7LW1tXHvvfcKEQDYru6+++6oq6tL5Oz77bffKf379y+RIgAA25sCGAAgIW6++eZRxcXFbZM4+yOPPBJr164VIgCwXa1atSoee+yxRM6ezWbb3njjjV+TIgAA25sCGAAgAYYOHVq21157fSeJs7/33nvxxBNPCBEA2CEmTpwYK1asSOTsQ4YM+c7QoUPLpAgAwPakAAYASIBrr732xKKiosqkzZ3L5eKee+6JxsZGIQIAO0R9fX3cddddkcvlEjd7UVFR5bXXXnuCFAEA2J4UwAAATdwBBxxQMWDAgG8kcfYZM2bEW2+9JUQAYId6880348UXX0zk7AMGDPjG8OHDy6UIAMD2ogAGAGjirrrqqhOKiooqkjb3pk2b4oEHHhAgALBT3H///VFbW5u4uYuKilpcffXV7gIGAGC7UQADADRhBx10UEX//v0Teffvww8/HB9++KEQAYCdYv369TFhwoREzj5w4MBvHHDAARVSBABge1AAAwA0Yf/5n/95bGFhYfOkzb1q1aqYPHmyAAGAnerpp5+O1atXJ27uwsLC8ssuu2yUBAEA2B4UwAAATdQBBxxQMWjQoNFJnP3uu++O+vp6IQIAO1V9fX3cddddiZx99913/+bQoUPLpAgAwGelAAYAaKIuv/zyYwsLC8uTNvfcuXNjwYIFAgQAdokFCxbE3LlzEzd3UVFRi9/97ndflyAAAJ+VAhgAoAkaOnRo2aBBgxL37t+6urr405/+JEAAYJf605/+FHV1dYmbe/fdd//mkCFDSiUIAMBnoQAGAGiCfvOb33ylqKioZdLmfu655xL53j0AIF1Wr14dzz33XOLmLioqann11VcfLUEAAD4LBTAAQBPTq1evZoMHD/5m0uaurq6ORx55RIAAQJPwyCOPRHV1deLmHjJkyLe6deuWlSAAANtKAQwA0MTcdNNNxxQXF7dN2twTJkyIqqoqAQIATUJVVVUifzmtuLi43a233uouYAAAtpkCGACgCWnVqlXhXnvtdVLS5l61alU8++yzAgQAmpTJkyfHqlWrEjf30KFDR7dq1apQggAAbAsFMABAE3LLLbccXlJS0jlpcz/44INRX18vQACgSamvr48HH3wwcXOXlJR0vummmw6VIAAA20IBDADQRGSz2cwBBxzw7aTNvWjRonjppZcECAA0SS+99FIsXrw4cXOPGDHiO9lsNiNBAAC2lgIYAKCJuOaaa/YuLy/vm7S5H3roocjlcgIEAJqkXC6XyLuAy8vL+1111VV7SRAAgK2lAAYAaCK+8pWvfDdpM8+bNy8WLlwoPACgSVu4cGG8+uqrrg8BAMgLCmAAgCbgsssuG1RZWblPkmbO5XIxfvx44QEAifDAAw8k7qklrVu33veSSy7pLz0AALaGAhgAoAkYNWrUCUmbefbs2bFs2TLhAQCJsGzZsnjllVcSN/cJJ5xwovQAANgaCmAAgF3sn//5nzt37NjxiCTN3NjYGA888IDwAIBEGT9+fDQ0NCRq5k6dOn1h9OjRHaQHAMCnpQAGANjFfvSjH30tk8kk6rps2rRpsWrVKuEBAImyatWqeOGFFxI1cyaTKfzpT386SnoAAHxaCmAAgF1o0KBBpX369Plqkmaur6+PCRMmCA8ASKQJEyZEXV1dombu27fvV/r27VssPQAAPg0FMADALnTZZZcdXlRUVJGkmadOnRpr164VHgCQSOvXr48pU6YkauaioqLK3/zmN0dIDwCAT0MBDACwi2Sz2cy+++57UpJmrqurc/cvAJB4jz76aOLuAt5///1PymazGekBAPBJFMAAALvI1VdfPbSsrKx3kmaeMmVKbNiwQXgAQKJt2LAhnn/++UTNXFZW1ueqq67aS3oAAHwSBTAAwC7y5S9/+bgkzVtfXx8TJ04UHACQCo8//nji7gL+0pe+dLzkAAD4JApgAIBdYPTo0R3atm07IkkzT5s2LdatWyc8ACAVPvzww5g+fXqiZm7fvv2I0aNHd5AeAABbogAGANgFfvrTn47KZDKFSZm3vr4+HnnkEcEBAKnyyCOPRH19fWLmzWQyhT/96U+/JjkAALZEAQwAsJN16tQp26dPn6OTNLO7fwGANFq3bl1MmzYtUTP36dPnmE6dOmWlBwDAx1EAAwDsZFddddUB2Wy2dVLmbWxsjEmTJgmOVOvYsWN06OCJmgD5aNKkSdHY2JiYebPZbOurrrrqAMkBAPBxFMAAADvZiBEjvp6keV988cVYtWqV4Ei1Ll26xIUXXhinnXZadO3a1UIA8siqVavipZdecj0JAEBqKIABAHaiM844o3tlZeXeSZk3l8vFxIkTBUdeyGQyMXjw4PjVr34VY8aMcUcwQB55/PHHI5fLJWbeysrKvc8444zukgMA4KMogAEAdqJTTjnlqxGRScq8CxYsiPfee09w5JVMJhN77713XHjhhTFmzJho3769pQCk3HvvvRcLFy5M1B9X/3NdCQAA/4cCGABgJ+nVq1ezXr16fTlJM3v3L/nsr0XwBRdcECeffHK0bdvWUgBSLGnXPb169fpyr169mkkOAIB/pAAGANhJrrjiioOLiopaJmXeBN4JAztEYWFhDB8+PC688MIYPXp0VFZWWgpACi1YsCCWLVuWmHmLiopaXnnllYdIDgCAf6QABgDYSYYPH/6VJM2btHfhwY5WVFQUI0aMiEsuuSRGjx4dLVu2tBSAFMnlcvH4448naub99tvvK5IDAOAfKYABAHaC0aNHd6isrByalHnXrl0bL7/8suDgI/y1CL744ovjhBNOiBYtWlgKQEq89NJLsW7dusTMW1lZudfo0aM7SA4AgL+lAAYA2AlOP/30o5J07fXMM89EQ0OD4GALiouL47DDDouxY8fGqFGjoqyszFIAEq6hoSGeeeaZJI1c8D/XmQAA8P8vEq0AAGDHymazmX79+n05KfPW1tbGlClTBAefUnFxcYwcOTIuvfTSGDVqVJSWlloKQII9//zzUVtbm5h5+/Xr9+VsNpuRHAAAf6UABgDYwX7zm9/sWVJS0jkp886YMSOqq6sFB1uppKQkRo4cGZdcckkcffTRUVJSYikACVRdXR0zZ85M0p8/na+44orBkgMA4K8UwAAAO9gXvvCFLyVl1lwuF08//bTQ4DNo3rx5HHXUUXHJJZfEyJEjI5vNWgpAwjz11FORy+USM++RRx75ZakBAPBXCmAAgB1oyJAhpZ07dz4iKfO+/vrrsWLFCsHBdlBeXh6jRo2KSy+9VBEMkDArVqyI119/PTHzdunS5fD+/ft79AQAABGhAAYA2KHGjh17aGFhYWJeCOruX9j+KioqYtSoUXHxxRfH4YcfHkVFRZYC4LpouyosLGz+H//xHwdLDQCACAUwAMAOteeeex6ZlFnXrl0b8+bNExrsIK1atYrjjz8+LrroohgxYkQUFPh2DKApmzdvXqxZsyYx8+61115HSg0AgAgFMADADnPMMce0bt269b5Jmfe5556LxsZGwcEO1qZNmxg9enRcfPHFimCAJqyxsTGee+65JP35MvyYY45pLTkAAPykAQBgBznzzDMPz2Qyibje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FpafffbZvaQFAJA+CmAAgE9p//3375+EOV9//XVhAXmjvLw8Ro0aFf/+7/8eI0eOjGw2aymA67FP6fOf//xAaQEApI8CGADgU+rRo8fuTX3G2traePvtt4UF5J2KiooYNWpUXHzxxXH44Yc3+XdvAum2ePHiRLySo1evXoOkBQCQPgpgAIBPqXXr1k3+DoklS5ZEQ0ODsIC81apVqzj++OPj4osvjhEjRkRBgW97gZ2voaEhlixZ0uTnbNOmjQIYACCFfCcMAPApDBkypLR58+a9m/qcb775prAAIqJ169YxevToGDt2rCIYcF32MZo3b95n0KBBpdICAEgX3wEDAHwKp556av9MJtPkr53eeustYQH8jTZt2sTo0aPjvPPOi+HDhyuCAddlfyOTyRT84Ac/+Jy0AADSxXe+AACfwuDBg5v84/Hq6+tj0aJFwgL4CJ06dYqTTz45/u3f/i323nvvyGQylgLsUIsXL07Eqzn23HPPgdICAEgXBTAAwKfQpUuXAU19xnfeeSfq6uqEBbAFnTt3jjFjxiiCgR2utrY23n333SRc53oPMABAyiiAAQA+hZYtW/Zv6jN6/DPAp9elS5cYM2ZMnHvuuTF48GALAfL2+iwJ17kAAGwdBTAAwCcYPnx4eUlJSeemPqfHPwNsvV69esVpp50W55xzTvTvrwMB8u/6rLS0tPPw4cPLpQUAkB4KYACAT/Ctb31rt4ho8s8IXbx4sbAAtlGfPn3irLPOinPOOSd22203CwG2i4T8gl7m29/+dj9pAQCkhwIYAOAT7L777k2+CVi7dm1s2LBBWACfUZ8+feLss8+OM888M3r27GkhwGfy4Ycfxrp165r8nAMHDlQAAwCkSJEVAABsWadOnZr8D8TefvttQQFsRwMGDIgBAwbEggULYvz48bF06VJLAbb5Oq1Vq1audwEA2GkUwAAAn6CyslIBDJCnBgwYEP3794958+bFQw89FO+++66lAFtlyZIlsddeezX1613PvgcASBEFMADAFnTq1CnbvHnzXk19ziVLlggLYAfJZDIxePDg2GOPPeLll1+OBx98MFauXGkxQGqu05o3b967Xbt2RatXr66XGABA8nkHMADAFowZM6ZnJpPJNuUZGxsbPZoUYCfIZDKx9957x4UXXhhjxoyJ9u3bWwrwiZYuXRqNjY1NesaCgoLsqaee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k8nE7NmzY+LEibFmzRoNAXif/VNSzJo161kTAwBongTAAAC70O233/7KOeecszU/P78022sdMGBAtG/f3tWu0My9+uqrMWnSpFi1apVmALyPjh07xgEHHJCIWhsbG6u/8Y1v/NXUAACaJwEwAMAutHLlyob169fP6tat2+hsrzWVSsXhhx8ekyZNMjhohubMmROTJ0+OJUuWaAbAh3D44YdHKpVKRK3r16+fVVFR0WhqAADNU54WAADsWq+++mpirts77LDDDAyamTfffDP+/d//PX7yk58IfwE+pFQqFSNGjEhMva+88sozpgYA0Hw5AQwAsIvddNNNz5166qnV+fn5rbO91m7dukX//v1j4cKFBgc5btmyZfHYY4/F/PnzNQPgIzrggAOic+fOiai1sbFxy4033jjL1AAAmi8BMADALrZy5cqG8vLyGT179vx0Euo9+uijBcCQw5YvXx7jxo0T/AJ8wv1SUpSXlz9TXl7eYGoAAM2XABgAYDeYMWPGUxdccEEiAuAhQ4ZE69ato7q62uAgh6xcuTImTJgQc+fOjUwmoyEAH1ObNm3i0EMPTUy9zz777FOmBgDQvPkGMADAbvDVr371lYaGhk1JqLWgoCCOOOIIQ4McsW7duvj5z38ed955Z8yZM0f4C/AJjRw5MvLz8xNRa0NDw8abbrrpVVMDAGjenAAGANgNKisrm1atWjW9T58+5ySh3qOPPjqefPJJQREk2Pr16+Oxxx6L2bNn+3cZYBdJpVIxatSoxNS7cuXKadXV1WmTAwBo3pwABgDYTaZNm5aY6/c6duwYAwYMMDRIoA0bNsSvf/3r+Nd//dd49dVXhb8Au9CAAQOiY8eOian36aefftLUAAAQAAMA7CZf+9rX5tTX11ckpd6jjjrK0CBBqqqq4v77749vf/vb8eKLL0ZTU5OmAOxiRx55ZGJqra+vX/eNb3zjdVMDAMAV0AAAu0ltbW16xYoVT++///6fT0K9hx56aLRt2zaqqqoMD7LYli1bYsqUKfHcc8/F9u3bNQRgN2nbtm0MGTIkMfUuX778qdraWtc/AwDgBDAAwO70xz/+cUJSas3Pz4/jjjvO0CBL1dTUxKOPPhq33XZb/OUvfxH+Auxmxx57bOTn5yel3Myvf/3rCaYGAECEABgAYLe66667llZXV89PSr1HH310tGjRwuAgi7wd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qATAAABZ4umnn968fPnyxF3xN3r06GjXrp0BAgBZqUOHDnH88ccnru5ly5ZNmD59+hYTBADgoxIAAwBkkbvvvvtPmUymKUk1FxYWxumnn254AEBWOvPMMxP3yYpMJtN41113/dH0AAD4OATAAABZ5Be/+MWatWvXTkta3UcccUT06tXLAAGArNK7d+847LDDEld3eXn5X+6///51JggAwMchAAYAyDIPP/zwn5JWcyqVijPPPNPwAICsMmbMmEilUomr+8EHH/yT6QEA8HEJgAEAsszXv/71NzZs2DAraXUPHDgwDj74YAMEALLCoEGDYsCAAYmru6KiYuatt966wAQBAPi4BMAAAFlo3Lhxv01i3WPGjIm8PFtMAGDvysvLizFjxiSy9j//+c+/NUEAAD7RflgLAACyz/XXXz+nqqrqr0mru1evXnHUUUcZIACwVx1zzDHRvXv3xNW9adOmV2666aa5JggAwCchAAYAyFJ/+tOf7k1i3WPGjInS0lIDBAD2ijZt2sRZZ52VyNofeOCBe00QAIBPSgAMAJClbrrpprlJPAVcXFwcZ555pgECAHvFWWedFS1btkxc3Zs2bXrl5ptvnmeCAAB8UgJgAIAsNm7cuF8lse5Ro0ZF7969DRAA2KP69OkTI0eOTGTtjz322K9MEACAXUEADACQxcaOHTu7qqrqb0mrO5VKxfnnnx+pVMoQAYA9tv/4/Oc/n8j9R2Vl5d+uueaav5kiAAC7ggAYACDLTZ069bdJrLtv374xZMgQAwQA9ojDDjsssTeQTJ48+TcmCADAriIABgDIcpdeeumLVVVVryax9s9//vNRXFxsiADAblVaWhrnn39+Imuvqqr66+WXX/6SKQIAsKsIgAEAEuChhx76WRLrLisri9NPP90AAYDd6qyzzoqSkpIklp753e9+91MTBABgVxIAAwAkwA033DB3w4YNs5JY+3HHHRd9+vQxRABgt+jbt28cddRRiay9oqJi1te//vU3TBEAgF1JAAwAkBA///nPfxoR6aTVnUql4vOf/3zk5dl6AgC7Vn5+flx00UWRSqWSWH76F7/4xX+bIgAAu5q3cAAACXHHHXe8tW7duulJrL13795xzDHHGCIAsEudcMIJ0b1790TWXl5e/pc77rjjLVMEAGBXEwADACTI3XfffW8mk2lKYu1nnXVWtG3b1hAB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onv37vEv//IvccABBzSrdWcymaYnnnji9s997nPPeAoAAODDEQADAJBYDz/88NJjjjlmY+/evY+KiGZzFKq0tDSGDx8eq1evjvXr13sQAHLcwIED45prrol27do1t6VnZsyY8YOzzjprqqcAAAA+PAEwAACJdv/99795zDHHVO6zzz5HRDMKgVu0aBGHHXZYtGzZMhYuXBjpdNrDAJBjCgoK4pxzzonzzz8/WrRo0dyWn37uuefuOuWUUyZ4EgAA4KMRAAMAkHi///3v5w8bNmxF3759j2lO10GnUqno27dvDB06NBYvXhxbtmzxMADkiJ49e8YNN9wQgwcPblbf+434x7XPU6ZM+fbpp58+2ZMAAAAfnQAYAICc8OCDDy4ZNmzYin79+jWrEDgionXr1nHEEUdEfX19LF261MMAkGCpVCpOOOGEuOKKK6JNmzbNbv07wt9vnXPOOX/xNAAAwMcjAAYAIGc89NBDS4YNG7a8X79+xza3EDg/Pz8OPvjg6NWrV8yfPz8aGho8EAAJU1JSEpdffnmccMIJkZ+f3+zWn8lkGp944olvffazn53maQAAgI9PAAwAQE556KGHlo4cOXJdnz59RqWa252ZEdG1a9cYNmxYLFu2LCorKz0QAAnRt2/fuO6662K//fZrluvPZDLpp59++rvnnHPO054GAAD4ZATAAADknD/96U+LDj300KX7779/s7sOOiKiuLg4jjzyyCgpKYk333wz0um0hwIgSxUUFMRnP/vZuPDCC6OkpKRZ9iCdTjc88sgjt5x//vnTPREAAPDJCYABAMhJDz/88NId3wQelUqlmt09mqlUKvr06RMHH3xwLFy4MLZt2+ahAMgynTt3jrFjx8bQoUOjGV5aERH/CH+feOKJ2y+88MLnPBEAALBrCIABAMhZDz300JIePXq8NmjQoGPz8vJaNMcetG3bNkaNGhVNTU2xZMkSDwVAFkilUjF69Oi48soro0OHDs22D01NTdt++9vf/stll132oqcCAAB2HQEwAAA5bdKkSeXdu3efM3jw4GYbAufn58eAAQOiV69esWDBgti+fbsHA2Avad26dVx66aVx/PHHR35+frPtQ2NjY/WvfvWrf7nuuute81QAAMCuJQAGACDnTZ48eW1jY+OMI4444uiCgoKS5tqHrl27xqhRo2Lbtm2xcuVKDwbAHpRKpWLUqFExduzY6NWrV7PuRX19/fo77rjjK9/61rcWejIAAGDXEwADANAsPP/881UbNmx45rjjjjuqsLCwrLn2obCwMAYNGhT77bdfLF68OGpraz0cALtZ+/bt44orrogTTzwxCgub9/uU2tralTfffPPVP/nJT1Z7MgAAYPcQAAMA0Gz87W9/27p27doZJ5xwwsjCwsK2zbkXnTp1ipEjR0Z1dbXTwAC70ciRI+Pqq6+OHj16NPte1NTULL/hhhuu/e1vf7vOkwEAALuPABgAgGbltdde2zpv3rynTznllKFFRUWdmnMvCgsL49BDD4399tsvFi1a5DQwwC709qnfk08+udmf+o2I2Lx587yLL774unHjxm30dAAAwO71fgFwat9zYsTOftC4LmLjq610DwCAxOrVq1fhs88++69du3Y9QTciGhoaYurUqTF58uRobGzUEICPqaCgIE499dQYPXq04HeH8vLyp4499tjvrly5skE3AABg9+swrDYKuuz8ZwJgAAByWuvWrfNeeumlm/fdd9+zdeMfVq9eHffff38sWbJEMwA+ov322y8uuugi1z2/w8KFC38/fPjwnzY0NGR0AwAA9gwBMAAAzd7zzz9/8aGHHnp1RKR0IyKTycTMmTPjz3/+c9TV1WkIwAcoKSmJ8847L0aMGBGplL9Kdki//PLL9xx77LEPagUAAOxZ7xcA+wYwAADNwn333Tfn+OOP39yrV6/DQwgcqVQqevfuHcOHD49169ZFRUWFhwTgPRxyyCExduzY6N+/v/B3h0wm0/jMM8/828knnzxONwAAYM97v28AC4ABAGg2fve7370xePDgJf369Ts6lUrl60hEcXFxjBgxInr06BFLly6N2tpaTQHYoUOHDvGFL3whzjzzzCguLtaQHZqammoeeOCBWz73uc9N1w0AANg7BMAAALDDww8/vKygoOC54cOHH1FQUFCqI//QrVu3OO6446K0tDQWL14cTU1NmgI0Wy1btoxzzjknLr300ujevbuGvENtbe2K22677arbbrvtDd0AAIC9RwAMAADv8Oyzz25avHjx0yeddNKQoqKiTjryD3l5edGnT58YOXJkbN26NVatWqUpQLNz+OGHx1e+8pUYMGBA5OXlacg7VFVV/e3CCy+8/oEHHvDdAAAA2MsEwAAA8C7z58+vefbZZ58+44wz+hcX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OigOOOCAKCoq0pQs0djYuGXmzJk/PvPMMyc2NDS4XgAAAPjQBMAAAJADjjzyyNY/+9nPrujbt++5EZGnI9mluro6Fi5cGPPnz4958+ZFZWWlprBXtGvXLg455JAYMGBA9O/fP1q39nsjWSj91ltv/fmqq676xaxZs6q1AwAA+KgEwAAAkEN++ctfjhgzZsyNrVq16q0b2SmdTsfKlStj0aJFsXDhwli8eHFs27ZNY9gtSkpKol+/ftG/f//Yf//9o1evXpGX53dEslVNTc3yRx999EdXXnnlK7oBAAB8XAJgAADIMd26dSt8+OGHPzd48ODL8vPzbdyzXCaTiTVr1sTChQtj0aJFsWjRotiyZYvG8LGUlZVF//79/xn6du/ePVKplMZkuaampprXXnvtV2PGjHmgoqKiUUcAAIBPQgAMAAA56rjjjiv7r//6r8tdC508mzdvjuXLl8eKFSti+fLlsXjx4qipqdEY/pfi4uLo169f9O7dO/bZZ5/Yd999o6ysTGMSJJPJpJcsWfLn66677pfTp0/3mx8AAMAuIQAGAIAc99Of/nTIueeee3NpaWlf3UimxsbGWLlyZSxdujSWLVsWy5cvj3Xr1kUmk9GcZiKVSkWXLl1in332iT59+kSfPn2iV69eUVBQoDkJtXXr1sUPPvjgj6699tq/6wYAALArCYABAKAZaNeuXf64cePOGTp06BUFBQWtdST56uvrY+XKlf88KbxixYpYu3ZtpNNpzUm4vLy86Nq1a+yzzz7//KdXr17RsmVLzckBjY2NW/7617/+4pxzznm0srKySUcAAIBdTQAMAADNyIknntjmnnvuuXzfffcdk0qlHB3MMdu3b481a9bEmjVrYu3atbF27dooLy+PDRs2CIazUF5eXnTs2DG6desWXbt2jW7dukW3bt2iR48eUVhYqEE5JpPJNC5ZsuRR1z0DAAC7mwAYAACaoa9+9av7Xnfdddd16NDhCN3IfY2Njf8MhNetWxfr1q2LioqKWLduXWzbtk2DdrPS0tLo3LnzP//p0qVLdO3aNbp27eoK52Ziw4YNM//rv/7rxz/60Y+W6wYAALC7CYABAKAZ++UvfznirLPOuq64uNj3gZupmpqaWL9+/T//qaioiMrKyqisrIxNmzZFY2OjJn2AgoKCaN++fbRr1y7at28fHTt2jC5dukSnTp2ic+fOUVxcrEnN1NatW98aP378PVdcccXLugEAAOwpAmAAAGjm2rVrl//ggw+eOWLEiCsLCwvb6gjvtHnz5n+GwZs2bYrKysqorq6OLVu2xJYtW6K6ujqqq6sjk8nk3NpTqVS0bt06WrduHWVlZdGmTZsoLS39X2Fvu3btok2bNh4U/peGhobKl1566efnnHPO+OrqavevAwAAe5QAGAAAiIiIoUOHFt97770XHHjggRfk5+c7ssiHlk6n/xkEV1dXR01Nzfv+k8lkora2NtLpdNTX10dTU1PU1dXt0u8U5+XlRcuWLSM/Pz+KiooiLy8vWrVqFalUKoqLi3f6T0lJSbRq1eqfgW9paWnk5eUZMB9aU1NTzYIFC/745S9/+Y+zZ8+u0REAAGBvEAADAAD/y2mnndb2rrvuurRPnz5n5+XlFeoIe9LbgfDbtm/f/r7XUBcUFESLFi3++b/fDnxhT8pkMg1LliwZd8stt/xq0qRJVToCAADsTQJgAABgp0477bS2d95554X777//5wTBAP9XJpNpWLhw4QO33nrrHwS/AABAtni/ADi/7UHRc2c/SG+LqC33/gcAAHLZokWL6u69995X0un0swcddFCnkpKS3roC8A8VFRUz/7//7/+77aKLLpq6aNGiOh0BAACyRXH3xsgr3fnPBMAAAEDMnDmz8u67736qqalpWr9+/Ypbt27dN5VKpXQGaG4ymUx6zZo1U+65555/Peeccx6cOXNmpa4AAADZRgAMAAB8KDNnzqz88Y9//Gw6nZ4uCAaak7eD37vvvvtfzz///McEvwAAQDYTAAMAAB+JIBhoLgS/AABAEgmAAQCAj+XtILisrOyFvn37diwuLu4VEYJgIBdkKioqnrv33nu/PWbMmEcEvwAAQJK8XwCc2vecGLGzHzSui9j4aivdAwAA/unqq6/u8ZWvfOX8Pn36nJWXl9dCR4CkSafT9UuXLh3/X//1Xw/84he/WKMjAABAEnUYVhsFXXb+MwEwAADwkZ1xxhntb7/99rMPPPDA8/Pz81vrCJDtmpqaqhcsWPDgd77znUcmTpzotC8AAJBoAmAAAGC3GD58eMkPf/jDzwwePPjioqKijjoCZJv6+voNr7322u9vvPHGx2fPnl2jIwAAQC4QAAMAALvV0KFDi++5556zDjnkkPOKioq66giwt9XV1a19/fXXH7z++uvHC34BAIBcIwAGAAD2iFatWuXdc889w0455ZTzO3bseJSOAHtYZsOGDbOmTJny4PXXX/9qbW1tWksAAIBcJAAGAAD2uH/913/tf/7555/dq1ev0/Ly8lroCLC7pNPp+pUrV05+4IEHHvnOd76zSEcAAIBcJwAGAAD2mjPOOKP97bfffvYBBxzw2YKCgjY6AuwqjY2NVW+++eafv/e97z06YcKETToCAAA0FwJgAABgrxs+fHjJ97///dGDBg0aU1paur+OAB9XdXX1wtdee+3Rr371q1Nfe+21Wh0BAACaGwEwAACQVb761a/ue8EFF3y6b9++ZxUUFLTWEeCDNDY2bnnrrbfG/+EPf5j4ox/9aLmOAAAAzZkAGAAAyEpDhw4t/sEPfnDy4MGDz27dunV/HQHerbq6+s2XX375weuuu+7ppUuXbtcRAAAAATAAAJDlCgsLU3ffffeQ0aNHn9G1a9fj8vLyinQFmq90Ol1XXl4+bcqUKROuvfbav+sIAADA/yYABgAAEqNPnz4t/v3f/33UyJEjz2rfvv2nIiKlK9AsZDZt2vTXF1544bGvfe1rzzntCwAA8N4EwAAAQCJddNFFXa6++uqTDzzwwLNbtmzZTUcg99TV1a1ZsGDBuJ/+9KdP3n///et0BAAA4IMJgAEAgETr1q1b4d13333k4YcffmqHDh2OyMvLK9QVSK50Ot2wcePGWbNmzZp8/fXXz6qoqGjUFQAAgA9PAAwAAOSMo48+uvWtt956/CGHHHJKu3btBkdEnq5AIqQrKytfmzdv3uQ777xz+owZM6q1BAAA4OMRAAMAADnpuOOOK/vGN75x/CGHHHJa27ZtB4bvBUO2yVRVVc2dN2/epO9///vTpk+fvkVLAAAAPjkBMAAAkPNuu+22vmedddZJffr0Oa5Vq1a9dQT2ntra2uXLly+f/uijjz51xx13vKUjAAAAu5YAGAAAaFa+9KUvdbv44ouP7t+//wlOBsMekamqqpq7cOHCv/z+97+fcd9995VrCQAAwO4jAAYAAJqtSy65pOtll112jDAYdrl0VVXVvIULF/7lV7/61bO/+93v1moJAADAniEABgAAiIirr766x/nnnz9q//33P7JNmzZDUqlUga7Ah5fJZBoqKyv/vnjx4uf+9Kc/zbr33ntX6woAAMCeJwAGAAB4l379+hV97WtfGzRy5MhRPXv2PLaoqKizrsD/VV9fv37VqlXPvPDCC8/9+7//+5zFixfX6woAAMDeJQAGAAB4H61bt8678847Bx599NFHde/efURpaen+4apomq/M1q1bF69Zs+bF5557btY3v/nNOdXV1WltAQAAyB4CYAAAgI9g8ODBrcaOHXvI8OHDD+vevfvw1q1bHxACYXJXprq6+s01a9a88sorr7z8k5/8ZN5rr71Wqy0AAADZSwAMAADwCVx//fX7nHHGGYf169dvePv27Yfm5+e31hWSrLGxsXrjxo2vvvXWWy+PHz/+lR//+McrdQUAACA5BMAAAAC70Je//OUe55xzzvA+ffoM7tix45CioqKuukI2q6+vX7t27doXFy9ePGfixImv3Xvvvat1BQAAILkEwAAAALvRl7/85R6f+cxnBvfr129Qly5dDm/ZsqVAmL2qrq5u7bp16wS+AAAAOUoADAAAsIe0atUq75prrtnn2GOPPXi//fY7uH379oeUlpb2TaVS+brDbpKuqalZtnHjxnlvvfXW3GeffXbef/3Xfy2vra1Naw0AAEBuEgADAADsRQceeGDLq6+++oAhQ4Yc3LNnz4Pbtm17sGuj+biampqqq6qqXi8vL583Z86ceb/85S/nvfjii1t1BgAAoPkQAAMAAGSZoUOHFn/xi1/cf9CgQQf26NHjwHbt2h3YqlWr3qlUKk932CFdU1OzvLKycsHq1asXzJkzZ8Gf//znJTNmzKjWGgAAgOZNAAwAAJAAxx13XNkFF1xw4MEHH3xA165dDywtLd23pKRkn1QqVag7uS2TyTRs27ZtRXV19dJ169YtfOONNxZOmDBh4YQJEzbpDgAAAO8mAAYAAEiw8847r+OJJ57Yp3///vt16dKlT5s2bfZr3bp1v/z8/GLdSZampqaa6urqxZs3b16ybt26pQsXLlzy9NNPL33ooYc26A4AAAAflgAYAAAgx7Rr1y7/kksu6TFkyJCe++67b89OnTr1Kisr61VcXNyzZcuW3VKpVL4u7R2ZTKaprq6uvKamZtWWLVtWVlRUrFy6dOnK2bNnr7r//vvXVFZWNukSAAAAn4QAGAAAoBnp1KlTwec+97luQ4cO7bnPPvv0aNeuXafWrVt3Li4u7tqyZcvORUVFnfLy8lro1MeTTqe319fXr6+rq6uoqalZW11dvb6y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AACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoIQADAAAAAAAAlBCAAQAAAAAAAEoIwAAAAAAAAAAlBGAAAAAAAACAEgIwAAAAAAAAQAkBGAAAAAAAAKCEAAwAAAAAAABQQgAGAAAAAAAAKCEAAwAAAAAAAJQQgAEAAAAAAABKCMAAAAAAAAAAJQRgAAAAAAAAgBICMAAAAAAAAEAJARgAAAAAAACghAAMAAAAAAAAUEIABgAAAAAAACghAAMAAAAAAACUEIABAAAAAAAASgjAAAAAAAAAACUEYAAAAAAAAIASAjAAAAAAAABACQEYAAAAAAAAoIQADAAAAAAAAFBCAAYAAAAAAAAoMY3JP//9K8My551lzpAk293a7QAAAAAAAAD8iIw3r+SL7TpPxuSt4/ibJJuz20QDBgAAAAAAAHhzrJPp6NvHYcyyjLl7dJTPbl/Jp6sPLufSsM+5ec7q5YOrF0OeP/SVaAAAAAAAAIA3weH7u6zfm79dWDIMSw6225w785vsx3XyzasOby5ssz69c4sAAAAAAAAAr9n69C6bC9sT93a7rKdnORin5Otxynzi1JQcfiwCAwAAAAAAALxO69O7HH68TaaT98cp85R8vbp3Nfszv8q0LHn7xMF1sjk/ZzUPmZ/MmWefhAYAAAAAAAD4oayPdjn7223Gn7x6ZrXPnetX8mBKkpv3cufDd/Pufn5FL56SzcVtNheT3dNt8tQlAwAAAAAAAHyv1sm0ySv/9XtsNWZ3I/lHkgzHi7/4XY7WYz5yiwAAAAAAAAA/Ivtcv/3nPEyS1fHa42t5+vNLyX7JoRsCAAAAAAAAePPt1/nbV3/KvePn1Xc371/NYxEYAAAAAAAA4M0yTHk2rvNo2WdzvHZqzN9v/SF3vju3evng/at5fPDLPDu1yuGyZHSVAAAAAAAAAK/XOGe4ueSvZ4e8M8xZbVe5ceuPufvy3Oqkw4+v5emDX+fuuTlLhhwsEYIBAAAAAAAAXpclGR8OuXP0TR4cPM/9z/+SRyfNDf/1TZezOp/87KdjzizJZkhO7edMy/w/nAUAAAAAAADg/zaMWcZkP8/ZLsmLacr2/MV8+cnvs/tP5/4FmLjAq1ifcioAAAAASUVORK5CYII=";

// vendor/pptxgenjs/src/gen-utils.ts
function getSmartParseNumber(size, xyDir, layout) {
  if (typeof size === "string" && !isNaN(Number(size))) size = Number(size);
  if (typeof size === "number" && size < 100) return inch2Emu(size);
  if (typeof size === "number" && size >= 100) return size;
  if (typeof size === "string" && size.includes("%")) {
    if (xyDir && xyDir === "X") return Math.round(parseFloat(size) / 100 * layout.width);
    if (xyDir && xyDir === "Y") return Math.round(parseFloat(size) / 100 * layout.height);
    return Math.round(parseFloat(size) / 100 * layout.width);
  }
  return 0;
}
function getUuid(uuidFormat) {
  return uuidFormat.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === "x" ? r : r & 3 | 8;
    return v.toString(16);
  });
}
function encodeXmlEntities(xml) {
  if (typeof xml === "undefined" || xml == null) return "";
  return xml.toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function inch2Emu(inches) {
  if (typeof inches === "number" && inches > 100) return inches;
  if (typeof inches === "string") inches = Number(inches.replace(/in*/gi, ""));
  return Math.round(EMU * inches);
}
function valToPts(pt) {
  const points = Number(pt) || 0;
  return isNaN(points) ? 0 : Math.round(points * ONEPT);
}
function convertRotationDegrees(d) {
  d = d || 0;
  return Math.round((d > 360 ? d - 360 : d) * 6e4);
}
function componentToHex(c) {
  const hex = c.toString(16);
  return hex.length === 1 ? "0" + hex : hex;
}
function rgbToHex(r, g, b) {
  return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
}
function createColorElement(colorStr, innerElements) {
  let colorVal = (colorStr || "").replace("#", "");
  if (!REGEX_HEX_COLOR.test(colorVal) && colorVal !== "bg1" /* background1 */ && colorVal !== "bg2" /* background2 */ && colorVal !== "tx1" /* text1 */ && colorVal !== "tx2" /* text2 */ && colorVal !== "accent1" /* accent1 */ && colorVal !== "accent2" /* accent2 */ && colorVal !== "accent3" /* accent3 */ && colorVal !== "accent4" /* accent4 */ && colorVal !== "accent5" /* accent5 */ && colorVal !== "accent6" /* accent6 */) {
    console.warn(`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'pptx.SchemeColor' values!`);
    colorVal = DEF_FONT_COLOR;
  }
  const tagName = REGEX_HEX_COLOR.test(colorVal) ? "srgbClr" : "schemeClr";
  const colorAttr = 'val="' + (REGEX_HEX_COLOR.test(colorVal) ? colorVal.toUpperCase() : colorVal) + '"';
  return innerElements ? `<a:${tagName} ${colorAttr}>${innerElements}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`;
}
function createGlowElement(options, defaults) {
  let strXml = "";
  const opts = { ...defaults, ...options };
  const size = Math.round(opts.size * ONEPT);
  const color = opts.color;
  const opacity = Math.round(opts.opacity * 1e5);
  strXml += `<a:glow rad="${size}">`;
  strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`);
  strXml += "</a:glow>";
  return strXml;
}
function genXmlColorSelection(props) {
  let fillType = "solid";
  let colorVal = "";
  let internalElements = "";
  let outText = "";
  if (props) {
    if (typeof props === "string") colorVal = props;
    else {
      if (props.type) fillType = props.type;
      if (props.color) colorVal = props.color;
      if (props.alpha) internalElements += `<a:alpha val="${Math.round((100 - props.alpha) * 1e3)}"/>`;
      if (props.transparency) internalElements += `<a:alpha val="${Math.round((100 - props.transparency) * 1e3)}"/>`;
    }
    switch (fillType) {
      case "solid":
        outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`;
        break;
      default:
        outText += "";
        break;
    }
  }
  return outText;
}
function getNewRelId(target) {
  return target._rels.length + target._relsChart.length + target._relsMedia.length + 1;
}
function correctShadowOptions(ShadowProps4) {
  if (!ShadowProps4 || typeof ShadowProps4 !== "object") {
    return;
  }
  if (ShadowProps4.type !== "outer" && ShadowProps4.type !== "inner" && ShadowProps4.type !== "none") {
    console.warn("Warning: shadow.type options are `outer`, `inner` or `none`.");
    ShadowProps4.type = "outer";
  }
  if (ShadowProps4.angle) {
    if (isNaN(Number(ShadowProps4.angle)) || ShadowProps4.angle < 0 || ShadowProps4.angle > 359) {
      console.warn("Warning: shadow.angle can only be 0-359");
      ShadowProps4.angle = 270;
    }
    ShadowProps4.angle = Math.round(Number(ShadowProps4.angle));
  }
  if (ShadowProps4.opacity) {
    if (isNaN(Number(ShadowProps4.opacity)) || ShadowProps4.opacity < 0 || ShadowProps4.opacity > 1) {
      console.warn("Warning: shadow.opacity can only be 0-1");
      ShadowProps4.opacity = 0.75;
    }
    ShadowProps4.opacity = Number(ShadowProps4.opacity);
  }
  if (ShadowProps4.color) {
    if (ShadowProps4.color.startsWith("#")) {
      console.warn('Warning: shadow.color should not include hash (#) character, , e.g. "FF0000"');
      ShadowProps4.color = ShadowProps4.color.replace("#", "");
    }
  }
  return ShadowProps4;
}

// vendor/pptxgenjs/src/gen-tables.ts
function parseTextToLines(cell, colWidth, verbose) {
  const FOCO = 2.3 + (cell.options?.autoPageCharWeight ? cell.options.autoPageCharWeight : 0);
  const CPL = Math.floor(colWidth / ONEPT * EMU) / ((cell.options?.fontSize ? cell.options.fontSize : DEF_FONT_SIZE) / FOCO);
  const parsedLines = [];
  let inputCells = [];
  const inputLines1 = [];
  const inputLines2 = [];
  if (cell.text && cell.text.toString().trim().length === 0) {
    inputCells.push({ _type: "tablecell" /* tablecell */, text: " " });
  } else if (typeof cell.text === "number" || typeof cell.text === "string") {
    inputCells.push({ _type: "tablecell" /* tablecell */, text: (cell.text || "").toString().trim() });
  } else if (Array.isArray(cell.text)) {
    inputCells = cell.text;
  }
  if (verbose) {
    console.log("[1/4] inputCells");
    inputCells.forEach((cell2, idx) => console.log(`[1/4] [${idx + 1}] cell: ${JSON.stringify(cell2)}`));
  }
  let newLine = [];
  inputCells.forEach((cell2) => {
    if (typeof cell2.text === "string") {
      if (cell2.text.split("\n").length > 1) {
        cell2.text.split("\n").forEach((textLine) => {
          newLine.push({
            _type: "tablecell" /* tablecell */,
            text: textLine,
            options: { ...cell2.options, ...{ breakLine: true } }
          });
        });
      } else {
        newLine.push({
          _type: "tablecell" /* tablecell */,
          text: cell2.text.trim(),
          options: cell2.options
        });
      }
      if (cell2.options?.breakLine) {
        if (verbose) console.log(`inputCells: new line > ${JSON.stringify(newLine)}`);
        inputLines1.push(newLine);
        newLine = [];
      }
    }
    if (newLine.length > 0) {
      inputLines1.push(newLine);
      newLine = [];
    }
  });
  if (verbose) {
    console.log(`[2/4] inputLines1 (${inputLines1.length})`);
    inputLines1.forEach((line, idx) => console.log(`[2/4] [${idx + 1}] line: ${JSON.stringify(line)}`));
  }
  inputLines1.forEach((line) => {
    line.forEach((cell2) => {
      const lineCells = [];
      const cellTextStr = String(cell2.text);
      const lineWords = cellTextStr.split(" ");
      lineWords.forEach((word, idx) => {
        const cellProps = { ...cell2.options };
        if (cellProps?.breakLine) cellProps.breakLine = idx + 1 === lineWords.length;
        lineCells.push({ _type: "tablecell" /* tablecell */, text: word + (idx + 1 < lineWords.length ? " " : ""), options: cellProps });
      });
      inputLines2.push(lineCells);
    });
  });
  if (verbose) {
    console.log(`[3/4] inputLines2 (${inputLines2.length})`);
    inputLines2.forEach((line) => console.log(`[3/4] line: ${JSON.stringify(line)}`));
  }
  inputLines2.forEach((line) => {
    let lineCells = [];
    let strCurrLine = "";
    line.forEach((word) => {
      if (strCurrLine.length + word.text.length > CPL) {
        parsedLines.push(lineCells);
        lineCells = [];
        strCurrLine = "";
      }
      lineCells.push(word);
      strCurrLine += word.text.toString();
    });
    if (lineCells.length > 0) parsedLines.push(lineCells);
  });
  if (verbose) {
    console.log(`[4/4] parsedLines (${parsedLines.length})`);
    parsedLines.forEach((line, idx) => console.log(`[4/4] [Line ${idx + 1}]:
${JSON.stringify(line)}`));
    console.log("...............................................\n\n");
  }
  return parsedLines;
}
function getSlidesForTableRows(tableRows = [], tableProps = {}, presLayout, masterSlide) {
  let arrInchMargins = DEF_SLIDE_MARGIN_IN;
  let emuSlideTabW = EMU * 1;
  let emuSlideTabH = EMU * 1;
  let emuTabCurrH = 0;
  let numCols = 0;
  const tableRowSlides = [];
  const tablePropX = getSmartParseNumber(tableProps.x, "X", presLayout);
  const tablePropY = getSmartParseNumber(tableProps.y, "Y", presLayout);
  const tablePropW = getSmartParseNumber(tableProps.w, "X", presLayout);
  const tablePropH = getSmartParseNumber(tableProps.h, "Y", presLayout);
  let tableCalcW = tablePropW;
  function calcSlideTabH() {
    let emuStartY = 0;
    if (tableRowSlides.length === 0) emuStartY = tablePropY || inch2Emu(arrInchMargins[0]);
    if (tableRowSlides.length > 0) emuStartY = inch2Emu(tableProps.autoPageSlideStartY || tableProps.newSlideStartY || arrInchMargins[0]);
    emuSlideTabH = (tablePropH || presLayout.height) - emuStartY - inch2Emu(arrInchMargins[2]);
    if (tableRowSlides.length > 1) {
      if (typeof tableProps.autoPageSlideStartY === "number") {
        emuSlideTabH = (tablePropH || presLayout.height) - inch2Emu(tableProps.autoPageSlideStartY + arrInchMargins[2]);
      } else if (typeof tableProps.newSlideStartY === "number") {
        emuSlideTabH = (tablePropH || presLayout.height) - inch2Emu(tableProps.newSlideStartY + arrInchMargins[2]);
      } else if (tablePropY) {
        emuSlideTabH = (tablePropH || presLayout.height) - inch2Emu((tablePropY / EMU < arrInchMargins[0] ? tablePropY / EMU : arrInchMargins[0]) + arrInchMargins[2]);
        if (emuSlideTabH < tablePropH) emuSlideTabH = tablePropH;
      }
    }
  }
  if (tableProps.verbose) {
    console.log("[[VERBOSE MODE]]");
    console.log("|-- TABLE PROPS --------------------------------------------------------|");
    console.log(`| presLayout.width ................................ = ${(presLayout.width / EMU).toFixed(1)}`);
    console.log(`| presLayout.height ............................... = ${(presLayout.height / EMU).toFixed(1)}`);
    console.log(`| tableProps.x .................................... = ${typeof tableProps.x === "number" ? (tableProps.x / EMU).toFixed(1) : tableProps.x}`);
    console.log(`| tableProps.y .................................... = ${typeof tableProps.y === "number" ? (tableProps.y / EMU).toFixed(1) : tableProps.y}`);
    console.log(`| tableProps.w .................................... = ${typeof tableProps.w === "number" ? (tableProps.w / EMU).toFixed(1) : tableProps.w}`);
    console.log(`| tableProps.h .................................... = ${typeof tableProps.h === "number" ? (tableProps.h / EMU).toFixed(1) : tableProps.h}`);
    console.log(`| tableProps.slideMargin .......................... = ${tableProps.slideMargin ? String(tableProps.slideMargin) : ""}`);
    console.log(`| tableProps.margin ............................... = ${String(tableProps.margin)}`);
    console.log(`| tableProps.colW ................................. = ${String(tableProps.colW)}`);
    console.log(`| tableProps.autoPageSlideStartY .................. = ${tableProps.autoPageSlideStartY}`);
    console.log(`| tableProps.autoPageCharWeight ................... = ${tableProps.autoPageCharWeight}`);
    console.log("|-- CALCULATIONS -------------------------------------------------------|");
    console.log(`| tablePropX ...................................... = ${tablePropX / EMU}`);
    console.log(`| tablePropY ...................................... = ${tablePropY / EMU}`);
    console.log(`| tablePropW ...................................... = ${tablePropW / EMU}`);
    console.log(`| tablePropH ...................................... = ${tablePropH / EMU}`);
    console.log(`| tableCalcW ...................................... = ${tableCalcW / EMU}`);
  }
  {
    if (!tableProps.slideMargin && tableProps.slideMargin !== 0) tableProps.slideMargin = DEF_SLIDE_MARGIN_IN[0];
    if (masterSlide && typeof masterSlide._margin !== "undefined") {
      if (Array.isArray(masterSlide._margin)) arrInchMargins = masterSlide._margin;
      else if (!isNaN(Number(masterSlide._margin))) {
        arrInchMargins = [Number(masterSlide._margin), Number(masterSlide._margin), Number(masterSlide._margin), Number(masterSlide._margin)];
      }
    } else if (tableProps.slideMargin || tableProps.slideMargin === 0) {
      if (Array.isArray(tableProps.slideMargin)) arrInchMargins = tableProps.slideMargin;
      else if (!isNaN(tableProps.slideMargin)) arrInchMargins = [tableProps.slideMargin, tableProps.slideMargin, tableProps.slideMargin, tableProps.slideMargin];
    }
    if (tableProps.verbose) console.log(`| arrInchMargins .................................. = [${arrInchMargins.join(", ")}]`);
  }
  {
    const firstRow = tableRows[0] || [];
    firstRow.forEach((cell) => {
      if (!cell) cell = { _type: "tablecell" /* tablecell */ };
      const cellOpts = cell.options || null;
      numCols += Number(cellOpts?.colspan ? cellOpts.colspan : 1);
    });
    if (tableProps.verbose) console.log(`| numCols ......................................... = ${numCols}`);
  }
  if (!tablePropW && tableProps.colW) {
    tableCalcW = Array.isArray(tableProps.colW) ? tableProps.colW.reduce((p, n) => p + n) * EMU : tableProps.colW * numCols || 0;
    if (tableProps.verbose) console.log(`| tableCalcW ...................................... = ${tableCalcW / EMU}`);
  }
  {
    emuSlideTabW = tableCalcW || inch2Emu((tablePropX ? tablePropX / EMU : arrInchMargins[1]) + arrInchMargins[3]);
    if (tableProps.verbose) console.log(`| emuSlideTabW .................................... = ${(emuSlideTabW / EMU).toFixed(1)}`);
  }
  if (!tableProps.colW || !Array.isArray(tableProps.colW)) {
    if (tableProps.colW && !isNaN(Number(tableProps.colW))) {
      const arrColW = [];
      const firstRow = tableRows[0] || [];
      firstRow.forEach(() => arrColW.push(tableProps.colW));
      tableProps.colW = [];
      arrColW.forEach((val) => {
        if (Array.isArray(tableProps.colW)) tableProps.colW.push(val);
      });
    } else {
      tableProps.colW = [];
      for (let iCol = 0; iCol < numCols; iCol++) {
        tableProps.colW.push(emuSlideTabW / EMU / numCols);
      }
    }
  }
  let newTableRowSlide = { rows: [] };
  tableRows.forEach((row, iRow) => {
    const rowCellLines = [];
    let maxCellMarTopEmu = 0;
    let maxCellMarBtmEmu = 0;
    let currTableRow = [];
    row.forEach((cell) => {
      currTableRow.push({
        _type: "tablecell" /* tablecell */,
        text: [],
        options: cell.options
      });
      if (cell.options.margin && cell.options.margin[0] >= 1) {
        if (cell.options?.margin && cell.options.margin[0] && valToPts(cell.options.margin[0]) > maxCellMarTopEmu) maxCellMarTopEmu = valToPts(cell.options.margin[0]);
        else if (tableProps?.margin && tableProps.margin[0] && valToPts(tableProps.margin[0]) > maxCellMarTopEmu) maxCellMarTopEmu = valToPts(tableProps.margin[0]);
        if (cell.options?.margin && cell.options.margin[2] && valToPts(cell.options.margin[2]) > maxCellMarBtmEmu) maxCellMarBtmEmu = valToPts(cell.options.margin[2]);
        else if (tableProps?.margin && tableProps.margin[2] && valToPts(tableProps.margin[2]) > maxCellMarBtmEmu) maxCellMarBtmEmu = valToPts(tableProps.margin[2]);
      } else {
        if (cell.options?.margin && cell.options.margin[0] && inch2Emu(cell.options.margin[0]) > maxCellMarTopEmu) maxCellMarTopEmu = inch2Emu(cell.options.margin[0]);
        else if (tableProps?.margin && tableProps.margin[0] && inch2Emu(tableProps.margin[0]) > maxCellMarTopEmu) maxCellMarTopEmu = inch2Emu(tableProps.margin[0]);
        if (cell.options?.margin && cell.options.margin[2] && inch2Emu(cell.options.margin[2]) > maxCellMarBtmEmu) maxCellMarBtmEmu = inch2Emu(cell.options.margin[2]);
        else if (tableProps?.margin && tableProps.margin[2] && inch2Emu(tableProps.margin[2]) > maxCellMarBtmEmu) maxCellMarBtmEmu = inch2Emu(tableProps.margin[2]);
      }
    });
    calcSlideTabH();
    emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu;
    if (tableProps.verbose && iRow === 0) console.log(`| SLIDE [${tableRowSlides.length}]: emuSlideTabH ...... = ${(emuSlideTabH / EMU).toFixed(1)} `);
    row.forEach((cell, iCell) => {
      const newCell = {
        _type: "tablecell" /* tablecell */,
        _lines: null,
        _lineHeight: inch2Emu(
          (cell.options?.fontSize ? cell.options.fontSize : tableProps.fontSize ? tableProps.fontSize : DEF_FONT_SIZE) * (LINEH_MODIFIER + (tableProps.autoPageLineWeight ? tableProps.autoPageLineWeight : 0)) / 100
        ),
        text: [],
        options: cell.options
      };
      if (newCell.options.rowspan) newCell._lineHeight = 0;
      newCell.options.autoPageCharWeight = tableProps.autoPageCharWeight ? tableProps.autoPageCharWeight : null;
      let totalColW = tableProps.colW[iCell];
      if (cell.options.colspan && Array.isArray(tableProps.colW)) {
        totalColW = tableProps.colW.filter((_cell, idx) => idx >= iCell && idx < idx + cell.options.colspan).reduce((prev, curr) => prev + curr);
      }
      newCell._lines = parseTextToLines(cell, totalColW, false);
      rowCellLines.push(newCell);
    });
    if (tableProps.verbose) console.log(`
| SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: START...`);
    let currCellIdx = 0;
    let emuLineMaxH = 0;
    let isDone = false;
    while (!isDone) {
      const srcCell = rowCellLines[currCellIdx];
      let tgtCell = currTableRow[currCellIdx];
      rowCellLines.forEach((cell) => {
        if (cell._lineHeight >= emuLineMaxH) emuLineMaxH = cell._lineHeight;
      });
      if (emuTabCurrH + emuLineMaxH > emuSlideTabH) {
        if (tableProps.verbose) {
          console.log("\n|-----------------------------------------------------------------------|");
          console.log(`|-- NEW SLIDE CREATED (currTabH+currLineH > maxH) => ${(emuTabCurrH / EMU).toFixed(2)} + ${(srcCell._lineHeight / EMU).toFixed(2)} > ${emuSlideTabH / EMU}`);
          console.log("|-----------------------------------------------------------------------|\n\n");
        }
        if (currTableRow.length > 0 && currTableRow.map((cell) => cell.text.length).reduce((p, n) => p + n) > 0) newTableRowSlide.rows.push(currTableRow);
        tableRowSlides.push(newTableRowSlide);
        const newRows = [];
        newTableRowSlide = { rows: newRows };
        currTableRow = [];
        row.forEach((cell) => currTableRow.push({ _type: "tablecell" /* tablecell */, text: [], options: cell.options }));
        calcSlideTabH();
        emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu;
        if (tableProps.verbose) console.log(`| SLIDE [${tableRowSlides.length}]: emuSlideTabH ...... = ${(emuSlideTabH / EMU).toFixed(1)} `);
        emuTabCurrH = 0;
        if ((tableProps.addHeaderToEach || tableProps.autoPageRepeatHeader) && tableProps._arrObjTabHeadRows) {
          tableProps._arrObjTabHeadRows.forEach((row2) => {
            const newHeadRow = [];
            let maxLineHeight = 0;
            row2.forEach((cell) => {
              newHeadRow.push(cell);
              if (cell._lineHeight > maxLineHeight) maxLineHeight = cell._lineHeight;
            });
            newTableRowSlide.rows.push(newHeadRow);
            emuTabCurrH += maxLineHeight;
          });
        }
        tgtCell = currTableRow[currCellIdx];
      }
      const currLine = srcCell._lines.shift();
      if (Array.isArray(tgtCell.text)) {
        if (currLine) tgtCell.text = tgtCell.text.concat(currLine);
        else if (tgtCell.text.length === 0) tgtCell.text = tgtCell.text.concat({ _type: "tablecell" /* tablecell */, text: "" });
      }
      if (currCellIdx === rowCellLines.length - 1) emuTabCurrH += emuLineMaxH;
      currCellIdx = currCellIdx < rowCellLines.length - 1 ? currCellIdx + 1 : 0;
      const brent = rowCellLines.map((cell) => cell._lines.length).reduce((prev, next) => prev + next);
      if (brent === 0) isDone = true;
    }
    if (currTableRow.length > 0) newTableRowSlide.rows.push(currTableRow);
    if (tableProps.verbose) {
      console.log(
        `- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: ...COMPLETE ...... emuTabCurrH = ${(emuTabCurrH / EMU).toFixed(2)} ( emuSlideTabH = ${(emuSlideTabH / EMU).toFixed(2)} )`
      );
    }
  });
  tableRowSlides.push(newTableRowSlide);
  if (tableProps.verbose) {
    console.log("\n|================================================|");
    console.log(`| FINAL: tableRowSlides.length = ${tableRowSlides.length}`);
    tableRowSlides.forEach((slide) => console.log(slide));
    console.log("|================================================|\n\n");
  }
  return tableRowSlides;
}
function genTableToSlides(pptx, tabEleId, options = {}, masterSlide) {
  const opts = options || {};
  opts.slideMargin = opts.slideMargin || opts.slideMargin === 0 ? opts.slideMargin : 0.5;
  let emuSlideTabW = opts.w || pptx.presLayout.width;
  const arrObjTabHeadRows = [];
  const arrObjTabBodyRows = [];
  const arrObjTabFootRows = [];
  const arrColW = [];
  const arrTabColW = [];
  let arrInchMargins = [0.5, 0.5, 0.5, 0.5];
  let intTabW = 0;
  if (!document.getElementById(tabEleId)) throw new Error('tableToSlides: Table ID "' + tabEleId + '" does not exist!');
  if (masterSlide?._margin) {
    if (Array.isArray(masterSlide._margin)) arrInchMargins = masterSlide._margin;
    else if (!isNaN(masterSlide._margin)) arrInchMargins = [masterSlide._margin, masterSlide._margin, masterSlide._margin, masterSlide._margin];
    opts.slideMargin = arrInchMargins;
  } else if (opts?.slideMargin) {
    if (Array.isArray(opts.slideMargin)) arrInchMargins = opts.slideMargin;
    else if (!isNaN(opts.slideMargin)) arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin];
  }
  emuSlideTabW = (opts.w ? inch2Emu(opts.w) : pptx.presLayout.width) - inch2Emu(arrInchMargins[1] + arrInchMargins[3]);
  if (opts.verbose) {
    console.log("[[VERBOSE MODE]]");
    console.log("|-- `tableToSlides` ----------------------------------------------------|");
    console.log(`| tableProps.h .................................... = ${opts.h}`);
    console.log(`| tableProps.w .................................... = ${opts.w}`);
    console.log(`| pptx.presLayout.width ........................... = ${(pptx.presLayout.width / EMU).toFixed(1)}`);
    console.log(`| pptx.presLayout.height .......................... = ${(pptx.presLayout.height / EMU).toFixed(1)}`);
    console.log(`| emuSlideTabW .................................... = ${(emuSlideTabW / EMU).toFixed(1)}`);
  }
  let firstRowCells = document.querySelectorAll(`#${tabEleId} tr:first-child th`);
  if (firstRowCells.length === 0) firstRowCells = document.querySelectorAll(`#${tabEleId} tr:first-child td`);
  firstRowCells.forEach((cellEle) => {
    const cell = cellEle;
    if (cell.getAttribute("colspan")) {
      for (let idxc = 0; idxc < Number(cell.getAttribute("colspan")); idxc++) {
        arrTabColW.push(Math.round(cell.offsetWidth / Number(cell.getAttribute("colspan"))));
      }
    } else {
      arrTabColW.push(cell.offsetWidth);
    }
  });
  arrTabColW.forEach((colW) => {
    intTabW += colW;
  });
  arrTabColW.forEach((colW, idxW) => {
    const intCalcWidth = Number((Number(emuSlideTabW) * (colW / intTabW * 100) / 100 / EMU).toFixed(2));
    let intMinWidth = 0;
    const colSelectorMin = document.querySelector(`#${tabEleId} thead tr:first-child th:nth-child(${idxW + 1})`);
    if (colSelectorMin) intMinWidth = Number(colSelectorMin.getAttribute("data-pptx-min-width"));
    const intSetWidth = 0;
    const colSelectorSet = document.querySelector(`#${tabEleId} thead tr:first-child th:nth-child(${idxW + 1})`);
    if (colSelectorSet) intMinWidth = Number(colSelectorSet.getAttribute("data-pptx-width"));
    arrColW.push(intSetWidth || (intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth));
  });
  if (opts.verbose) {
    console.log(`| arrColW ......................................... = [${arrColW.join(", ")}]`);
  }
  const tableParts = ["thead", "tbody", "tfoot"];
  tableParts.forEach((part) => {
    document.querySelectorAll(`#${tabEleId} ${part} tr`).forEach((row) => {
      const htmlRow = row;
      const arrObjTabCells = [];
      Array.from(htmlRow.cells).forEach((cell) => {
        const arrRGB1 = window.getComputedStyle(cell).getPropertyValue("color").replace(/\s+/gi, "").replace("rgba(", "").replace("rgb(", "").replace(")", "").split(",");
        let arrRGB2 = window.getComputedStyle(cell).getPropertyValue("background-color").replace(/\s+/gi, "").replace("rgba(", "").replace("rgb(", "").replace(")", "").split(",");
        if (
          // NOTE: (ISSUE#57): Default for unstyled tables is black bkgd, so use white instead
          window.getComputedStyle(cell).getPropertyValue("background-color") === "rgba(0, 0, 0, 0)" || window.getComputedStyle(cell).getPropertyValue("transparent")
        ) {
          arrRGB2 = ["255", "255", "255"];
        }
        const cellOpts = {
          align: null,
          bold: !!(window.getComputedStyle(cell).getPropertyValue("font-weight") === "bold" || Number(window.getComputedStyle(cell).getPropertyValue("font-weight")) >= 500),
          border: null,
          color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
          fill: { color: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])) },
          fontFace: (window.getComputedStyle(cell).getPropertyValue("font-family") || "").split(",")[0].replace(/"/g, "").replace("inherit", "").replace("initial", "") || null,
          fontSize: Number(window.getComputedStyle(cell).getPropertyValue("font-size").replace(/[a-z]/gi, "")),
          margin: null,
          colspan: Number(cell.getAttribute("colspan")) || null,
          rowspan: Number(cell.getAttribute("rowspan")) || null,
          valign: null
        };
        if (["left", "center", "right", "start", "end"].includes(window.getComputedStyle(cell).getPropertyValue("text-align"))) {
          const align = window.getComputedStyle(cell).getPropertyValue("text-align").replace("start", "left").replace("end", "right");
          cellOpts.align = align === "center" ? "center" : align === "left" ? "left" : align === "right" ? "right" : null;
        }
        if (["top", "middle", "bottom"].includes(window.getComputedStyle(cell).getPropertyValue("vertical-align"))) {
          const valign = window.getComputedStyle(cell).getPropertyValue("vertical-align");
          cellOpts.valign = valign === "top" ? "top" : valign === "middle" ? "middle" : valign === "bottom" ? "bottom" : null;
        }
        if (window.getComputedStyle(cell).getPropertyValue("padding-left")) {
          cellOpts.margin = [0, 0, 0, 0];
          const sidesPad = ["padding-top", "padding-right", "padding-bottom", "padding-left"];
          sidesPad.forEach((val, idxs) => {
            cellOpts.margin[idxs] = Math.round(Number(window.getComputedStyle(cell).getPropertyValue(val).replace(/\D/gi, "")));
          });
        }
        if (window.getComputedStyle(cell).getPropertyValue("border-top-width") || window.getComputedStyle(cell).getPropertyValue("border-right-width") || window.getComputedStyle(cell).getPropertyValue("border-bottom-width") || window.getComputedStyle(cell).getPropertyValue("border-left-width")) {
          cellOpts.border = [null, null, null, null];
          const sidesBor = ["top", "right", "bottom", "left"];
          sidesBor.forEach((val, idxb) => {
            const intBorderW = Math.round(
              Number(
                window.getComputedStyle(cell).getPropertyValue("border-" + val + "-width").replace("px", "")
              )
            );
            let arrRGB = [];
            arrRGB = window.getComputedStyle(cell).getPropertyValue("border-" + val + "-color").replace(/\s+/gi, "").replace("rgba(", "").replace("rgb(", "").replace(")", "").split(",");
            const strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
            cellOpts.border[idxb] = { pt: intBorderW, color: strBorderC };
          });
        }
        arrObjTabCells.push({
          _type: "tablecell" /* tablecell */,
          text: cell.innerText,
          // `innerText` returns <br> as "\n", so linebreak etc. work later!
          options: cellOpts
        });
      });
      switch (part) {
        case "thead":
          arrObjTabHeadRows.push(arrObjTabCells);
          break;
        case "tbody":
          arrObjTabBodyRows.push(arrObjTabCells);
          break;
        case "tfoot":
          arrObjTabFootRows.push(arrObjTabCells);
          break;
        default:
          console.log(`table parsing: unexpected table part: ${part}`);
          break;
      }
    });
  });
  opts._arrObjTabHeadRows = arrObjTabHeadRows || null;
  opts.colW = arrColW;
  getSlidesForTableRows([...arrObjTabHeadRows, ...arrObjTabBodyRows, ...arrObjTabFootRows], opts, pptx.presLayout, masterSlide).forEach((slide, idxTr) => {
    const newSlide = pptx.addSlide({ masterName: opts.masterSlideName || null });
    if (idxTr === 0) opts.y = opts.y || arrInchMargins[0];
    if (idxTr > 0) opts.y = opts.autoPageSlideStartY || opts.newSlideStartY || arrInchMargins[0];
    if (opts.verbose) console.log(`| opts.autoPageSlideStartY: ${opts.autoPageSlideStartY} / arrInchMargins[0]: ${arrInchMargins[0]} => opts.y = ${opts.y}`);
    newSlide.addTable(slide.rows, { x: opts.x || arrInchMargins[3], y: opts.y, w: Number(emuSlideTabW) / EMU, colW: arrColW, autoPage: false });
    if (opts.addImage) {
      opts.addImage.options = opts.addImage.options || {};
      if (!opts.addImage.image || !opts.addImage.image.path && !opts.addImage.image.data) {
        console.warn("Warning: tableToSlides.addImage requires either `path` or `data`");
      } else {
        newSlide.addImage({
          path: opts.addImage.image.path,
          data: opts.addImage.image.data,
          x: opts.addImage.options.x,
          y: opts.addImage.options.y,
          w: opts.addImage.options.w,
          h: opts.addImage.options.h
        });
      }
    }
    if (opts.addShape) newSlide.addShape(opts.addShape.shapeName, opts.addShape.options || {});
    if (opts.addTable) newSlide.addTable(opts.addTable.rows, opts.addTable.options || {});
    if (opts.addText) newSlide.addText(opts.addText.text, opts.addText.options || {});
  });
}

// vendor/pptxgenjs/src/gen-objects.ts
var _chartCounter = 0;
function createSlideMaster(props, target) {
  if (props.bkgd) target.bkgd = props.bkgd;
  if (props.objects && Array.isArray(props.objects) && props.objects.length > 0) {
    props.objects.forEach((object, idx) => {
      const key = Object.keys(object)[0];
      const tgt = target;
      if (MASTER_OBJECTS[key] && key === "chart") addChartDefinition(tgt, object[key].type, object[key].data, object[key].opts);
      else if (MASTER_OBJECTS[key] && key === "image") addImageDefinition(tgt, object[key]);
      else if (MASTER_OBJECTS[key] && key === "line") addShapeDefinition(tgt, "line" /* LINE */, object[key]);
      else if (MASTER_OBJECTS[key] && key === "rect") addShapeDefinition(tgt, "rect" /* RECTANGLE */, object[key]);
      else if (MASTER_OBJECTS[key] && key === "text") addTextDefinition(tgt, [{ text: object[key].text }], object[key].options, false);
      else if (MASTER_OBJECTS[key] && key === "placeholder") {
        object[key].options.placeholder = object[key].options.name;
        delete object[key].options.name;
        object[key].options._placeholderType = object[key].options.type;
        delete object[key].options.type;
        object[key].options._placeholderIdx = 100 + idx;
        addTextDefinition(tgt, [{ text: object[key].text }], object[key].options, true);
      }
    });
  }
  if (props.slideNumber && typeof props.slideNumber === "object") target._slideNumberProps = props.slideNumber;
}
function addChartDefinition(target, type, data, opt) {
  function correctGridLineOptions(glOpts) {
    if (!glOpts || glOpts.style === "none") return;
    if (glOpts.size !== void 0 && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
      console.warn("Warning: chart.gridLine.size must be greater than 0.");
      delete glOpts.size;
    }
    if (glOpts.style && !["solid", "dash", "dot"].includes(glOpts.style)) {
      console.warn("Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.");
      delete glOpts.style;
    }
    if (glOpts.cap && !["flat", "square", "round"].includes(glOpts.cap)) {
      console.warn("Warning: chart.gridLine.cap options: `flat`, `square`, `round`.");
      delete glOpts.cap;
    }
  }
  const chartId = ++_chartCounter;
  const resultObject = {
    _type: null,
    text: null,
    options: null,
    chartRid: null
  };
  let tmpOpt = null;
  let tmpData = [];
  if (Array.isArray(type)) {
    type.forEach((obj) => {
      tmpData = tmpData.concat(obj.data);
    });
    tmpOpt = data || opt;
  } else {
    tmpData = data;
    tmpOpt = opt;
  }
  tmpData.forEach((item, i) => {
    item._dataIndex = i;
    if (item.labels !== void 0 && !Array.isArray(item.labels[0])) {
      item.labels = [item.labels];
    }
  });
  const options = tmpOpt && typeof tmpOpt === "object" ? tmpOpt : {};
  options._type = type;
  options.x = typeof options.x !== "undefined" && options.x != null && !isNaN(Number(options.x)) ? options.x : 1;
  options.y = typeof options.y !== "undefined" && options.y != null && !isNaN(Number(options.y)) ? options.y : 1;
  options.w = options.w || "50%";
  options.h = options.h || "50%";
  options.objectName = options.objectName ? encodeXmlEntities(options.objectName) : `Chart ${target._slideObjects.filter((obj) => obj._type === "chart" /* chart */).length}`;
  if (!["bar", "col"].includes(options.barDir || "")) options.barDir = "col";
  if (options._type === "area" /* AREA */) {
    if (!["stacked", "standard", "percentStacked"].includes(options.barGrouping || "")) options.barGrouping = "standard";
  }
  if (options._type === "bar" /* BAR */) {
    if (!["clustered", "stacked", "percentStacked"].includes(options.barGrouping || "")) options.barGrouping = "clustered";
  }
  if (options._type === "bar3D" /* BAR3D */) {
    if (!["clustered", "stacked", "standard", "percentStacked"].includes(options.barGrouping || "")) options.barGrouping = "standard";
  }
  if (options.barGrouping?.includes("tacked")) {
    if (!options.barGapWidthPct) options.barGapWidthPct = 50;
  }
  if (options.dataLabelPosition) {
    if (options._type === "area" /* AREA */ || options._type === "bar3D" /* BAR3D */ || options._type === "doughnut" /* DOUGHNUT */ || options._type === "radar" /* RADAR */) {
      delete options.dataLabelPosition;
    }
    if (options._type === "pie" /* PIE */) {
      if (!["bestFit", "ctr", "inEnd", "outEnd"].includes(options.dataLabelPosition)) delete options.dataLabelPosition;
    }
    if (options._type === "bubble" /* BUBBLE */ || options._type === "bubble3D" /* BUBBLE3D */ || options._type === "line" /* LINE */ || options._type === "scatter" /* SCATTER */) {
      if (!["b", "ctr", "l", "r", "t"].includes(options.dataLabelPosition)) delete options.dataLabelPosition;
    }
    if (options._type === "bar" /* BAR */) {
      if (!["stacked", "percentStacked"].includes(options.barGrouping || "")) {
        if (!["ctr", "inBase", "inEnd"].includes(options.dataLabelPosition)) delete options.dataLabelPosition;
      }
      if (!["clustered"].includes(options.barGrouping || "")) {
        if (!["ctr", "inBase", "inEnd", "outEnd"].includes(options.dataLabelPosition)) delete options.dataLabelPosition;
      }
    }
  }
  options.dataLabelBkgrdColors = options.dataLabelBkgrdColors || !options.dataLabelBkgrdColors ? options.dataLabelBkgrdColors : false;
  if (!["b", "l", "r", "t", "tr"].includes(options.legendPos || "")) options.legendPos = "r";
  if (!["cone", "coneToMax", "box", "cylinder", "pyramid", "pyramidToMax"].includes(options.bar3DShape || "")) options.bar3DShape = "box";
  if (!["circle", "dash", "diamond", "dot", "none", "square", "triangle"].includes(options.lineDataSymbol || "")) options.lineDataSymbol = "circle";
  if (!["gap", "span"].includes(options.displayBlanksAs || "")) options.displayBlanksAs = "span";
  if (!["standard", "marker", "filled"].includes(options.radarStyle || "")) options.radarStyle = "standard";
  options.lineDataSymbolSize = options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize) ? options.lineDataSymbolSize : 6;
  options.lineDataSymbolLineSize = options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? valToPts(options.lineDataSymbolLineSize) : valToPts(0.75);
  if (options.layout) {
    ["x", "y", "w", "h"].forEach((key) => {
      const val = options.layout[key];
      if (isNaN(Number(val)) || val < 0 || val > 1) {
        console.warn("Warning: chart.layout." + key + " can only be 0-1");
        delete options.layout[key];
      }
    });
  }
  options.catGridLine = options.catGridLine || (options._type === "scatter" /* SCATTER */ ? { color: "D9D9D9", size: 1 } : { style: "none" });
  options.valGridLine = options.valGridLine || (options._type === "scatter" /* SCATTER */ ? { color: "D9D9D9", size: 1 } : {});
  options.serGridLine = options.serGridLine || (options._type === "scatter" /* SCATTER */ ? { color: "D9D9D9", size: 1 } : { style: "none" });
  correctGridLineOptions(options.catGridLine);
  correctGridLineOptions(options.valGridLine);
  correctGridLineOptions(options.serGridLine);
  correctShadowOptions(options.shadow);
  options.showDataTable = options.showDataTable || !options.showDataTable ? options.showDataTable : false;
  options.showDataTableHorzBorder = options.showDataTableHorzBorder || !options.showDataTableHorzBorder ? options.showDataTableHorzBorder : true;
  options.showDataTableVertBorder = options.showDataTableVertBorder || !options.showDataTableVertBorder ? options.showDataTableVertBorder : true;
  options.showDataTableOutline = options.showDataTableOutline || !options.showDataTableOutline ? options.showDataTableOutline : true;
  options.showDataTableKeys = options.showDataTableKeys || !options.showDataTableKeys ? options.showDataTableKeys : true;
  options.showLabel = options.showLabel || !options.showLabel ? options.showLabel : false;
  options.showLegend = options.showLegend || !options.showLegend ? options.showLegend : false;
  options.showPercent = options.showPercent || !options.showPercent ? options.showPercent : true;
  options.showTitle = options.showTitle || !options.showTitle ? options.showTitle : false;
  options.showValue = options.showValue || !options.showValue ? options.showValue : false;
  options.showLeaderLines = options.showLeaderLines || !options.showLeaderLines ? options.showLeaderLines : false;
  options.catAxisLineShow = typeof options.catAxisLineShow !== "undefined" ? options.catAxisLineShow : true;
  options.valAxisLineShow = typeof options.valAxisLineShow !== "undefined" ? options.valAxisLineShow : true;
  options.serAxisLineShow = typeof options.serAxisLineShow !== "undefined" ? options.serAxisLineShow : true;
  options.v3DRotX = !isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90 ? options.v3DRotX : 30;
  options.v3DRotY = !isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360 ? options.v3DRotY : 30;
  options.v3DRAngAx = options.v3DRAngAx || !options.v3DRAngAx ? options.v3DRAngAx : true;
  options.v3DPerspective = !isNaN(options.v3DPerspective) && options.v3DPerspective >= 0 && options.v3DPerspective <= 240 ? options.v3DPerspective : 30;
  options.barGapWidthPct = !isNaN(options.barGapWidthPct) && options.barGapWidthPct >= 0 && options.barGapWidthPct <= 1e3 ? options.barGapWidthPct : 150;
  options.barGapDepthPct = !isNaN(options.barGapDepthPct) && options.barGapDepthPct >= 0 && options.barGapDepthPct <= 1e3 ? options.barGapDepthPct : 150;
  options.chartColors = Array.isArray(options.chartColors) ? options.chartColors : options._type === "pie" /* PIE */ || options._type === "doughnut" /* DOUGHNUT */ ? PIECHART_COLORS : BARCHART_COLORS;
  options.chartColorsOpacity = options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : null;
  options.border = options.border && typeof options.border === "object" ? options.border : null;
  if (options.border && (!options.border.pt || isNaN(options.border.pt))) options.border.pt = DEF_CHART_BORDER.pt;
  if (options.border && (!options.border.color || typeof options.border.color !== "string")) options.border.color = DEF_CHART_BORDER.color;
  options.plotArea = options.plotArea || {};
  options.plotArea.border = options.plotArea.border && typeof options.plotArea.border === "object" ? options.plotArea.border : null;
  if (options.plotArea.border && (!options.plotArea.border.pt || isNaN(options.plotArea.border.pt))) options.plotArea.border.pt = DEF_CHART_BORDER.pt;
  if (options.plotArea.border && (!options.plotArea.border.color || typeof options.plotArea.border.color !== "string")) {
    options.plotArea.border.color = DEF_CHART_BORDER.color;
  }
  if (options.border) options.plotArea.border = options.border;
  options.plotArea.fill = options.plotArea.fill || { color: null, transparency: null };
  if (options.fill) options.plotArea.fill.color = options.fill;
  options.chartArea = options.chartArea || {};
  options.chartArea.border = options.chartArea.border && typeof options.chartArea.border === "object" ? options.chartArea.border : null;
  if (options.chartArea.border) {
    options.chartArea.border = {
      color: options.chartArea.border.color || DEF_CHART_BORDER.color,
      pt: options.chartArea.border.pt || DEF_CHART_BORDER.pt
    };
  }
  options.chartArea.roundedCorners = typeof options.chartArea.roundedCorners === "boolean" ? options.chartArea.roundedCorners : true;
  options.dataBorder = options.dataBorder && typeof options.dataBorder === "object" ? options.dataBorder : null;
  if (options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt))) options.dataBorder.pt = 0.75;
  if (options.dataBorder && (!options.dataBorder.color || typeof options.dataBorder.color !== "string" || options.dataBorder.color.length !== 6)) {
    options.dataBorder.color = "F9F9F9";
  }
  if (!options.dataLabelFormatCode && options._type === "scatter" /* SCATTER */) options.dataLabelFormatCode = "General";
  if (!options.dataLabelFormatCode && (options._type === "pie" /* PIE */ || options._type === "doughnut" /* DOUGHNUT */)) {
    options.dataLabelFormatCode = options.showPercent ? "0%" : "General";
  }
  options.dataLabelFormatCode = options.dataLabelFormatCode && typeof options.dataLabelFormatCode === "string" ? options.dataLabelFormatCode : "#,##0";
  if (!options.dataLabelFormatScatter && options._type === "scatter" /* SCATTER */) options.dataLabelFormatScatter = "custom";
  options.lineSize = typeof options.lineSize === "number" ? options.lineSize : 2;
  options.valAxisMajorUnit = typeof options.valAxisMajorUnit === "number" ? options.valAxisMajorUnit : null;
  if (options._type === "area" /* AREA */ || options._type === "bar" /* BAR */ || options._type === "bar3D" /* BAR3D */ || options._type === "line" /* LINE */) {
    options.catAxisMultiLevelLabels = !!options.catAxisMultiLevelLabels;
  } else {
    delete options.catAxisMultiLevelLabels;
  }
  resultObject._type = "chart";
  resultObject.options = options;
  resultObject.chartRid = getNewRelId(target);
  target._relsChart.push({
    rId: getNewRelId(target),
    data: tmpData,
    opts: options,
    type: options._type,
    globalId: chartId,
    fileName: `chart${chartId}.xml`,
    Target: `/ppt/charts/chart${chartId}.xml`
  });
  target._slideObjects.push(resultObject);
  return resultObject;
}
function addImageDefinition(target, opt) {
  const newObject = {
    _type: null,
    text: null,
    options: null,
    image: null,
    imageRid: null,
    hyperlink: null
  };
  const intPosX = opt.x || 0;
  const intPosY = opt.y || 0;
  const intWidth = opt.w || 0;
  const intHeight = opt.h || 0;
  const sizing = opt.sizing || null;
  const objHyperlink = opt.hyperlink || "";
  const strImageData = opt.data || "";
  const strImagePath = opt.path || "";
  let imageRelId = getNewRelId(target);
  const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Image ${target._slideObjects.filter((obj) => obj._type === "image" /* image */).length}`;
  if (!strImagePath && !strImageData) {
    console.error("ERROR: addImage() requires either 'data' or 'path' parameter!");
    return null;
  } else if (strImagePath && typeof strImagePath !== "string") {
    console.error(`ERROR: addImage() 'path' should be a string, ex: {path:'/img/sample.png'} - you sent ${String(strImagePath)}`);
    return null;
  } else if (strImageData && typeof strImageData !== "string") {
    console.error(`ERROR: addImage() 'data' should be a string, ex: {data:'image/png;base64,NMP[...]'} - you sent ${String(strImageData)}`);
    return null;
  } else if (strImageData && typeof strImageData === "string" && !strImageData.toLowerCase().includes("base64,")) {
    console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')");
    return null;
  }
  let strImgExtn = (strImagePath.substring(strImagePath.lastIndexOf("/") + 1).split("?")[0].split(".").pop().split("#")[0] || "png").toLowerCase();
  if (strImageData && /image\/(\w+);/.exec(strImageData) && /image\/(\w+);/.exec(strImageData).length > 0) {
    strImgExtn = /image\/(\w+);/.exec(strImageData)[1];
  } else if (strImageData?.toLowerCase().includes("image/svg+xml")) {
    strImgExtn = "svg";
  }
  newObject._type = "image" /* image */;
  newObject.image = strImagePath || "preencoded.png";
  newObject.options = {
    x: intPosX || 0,
    y: intPosY || 0,
    w: intWidth || 1,
    h: intHeight || 1,
    altText: opt.altText || "",
    rounding: typeof opt.rounding === "boolean" ? opt.rounding : false,
    sizing,
    placeholder: opt.placeholder,
    rotate: opt.rotate || 0,
    flipV: opt.flipV || false,
    flipH: opt.flipH || false,
    transparency: opt.transparency || 0,
    objectName,
    shadow: correctShadowOptions(opt.shadow)
  };
  if (strImgExtn === "svg") {
    target._relsMedia.push({
      path: strImagePath || strImageData + "png",
      type: "image/png",
      extn: "png",
      data: strImageData || "",
      rId: imageRelId,
      Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`,
      isSvgPng: true,
      svgSize: { w: getSmartParseNumber(newObject.options.w, "X", target._presLayout), h: getSmartParseNumber(newObject.options.h, "Y", target._presLayout) }
    });
    newObject.imageRid = imageRelId;
    target._relsMedia.push({
      path: strImagePath || strImageData,
      type: "image/svg+xml",
      extn: strImgExtn,
      data: strImageData || "",
      rId: imageRelId + 1,
      Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.${strImgExtn}`
    });
    newObject.imageRid = imageRelId + 1;
  } else {
    const dupeItem = target._relsMedia.filter((item) => item.path && item.path === strImagePath && item.type === "image/" + strImgExtn && !item.isDuplicate)[0];
    target._relsMedia.push({
      path: strImagePath || "preencoded." + strImgExtn,
      type: "image/" + strImgExtn,
      extn: strImgExtn,
      data: strImageData || "",
      rId: imageRelId,
      isDuplicate: !!dupeItem?.Target,
      Target: dupeItem?.Target ? dupeItem.Target : `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.${strImgExtn}`
    });
    newObject.imageRid = imageRelId;
  }
  if (typeof objHyperlink === "object") {
    if (!objHyperlink.url && !objHyperlink.slide) throw new Error("ERROR: `hyperlink` option requires either: `url` or `slide`");
    else {
      imageRelId++;
      target._rels.push({
        type: "hyperlink" /* hyperlink */,
        data: objHyperlink.slide ? "slide" : "dummy",
        rId: imageRelId,
        Target: objHyperlink.url || objHyperlink.slide.toString()
      });
      objHyperlink._rId = imageRelId;
      newObject.hyperlink = objHyperlink;
    }
  }
  target._slideObjects.push(newObject);
}
function addMediaDefinition(target, opt) {
  const intPosX = opt.x || 0;
  const intPosY = opt.y || 0;
  const intSizeX = opt.w || 2;
  const intSizeY = opt.h || 2;
  const strData = opt.data || "";
  const strLink = opt.link || "";
  const strPath = opt.path || "";
  const strType = opt.type || "audio";
  let strExtn = "";
  const strCover = opt.cover || IMG_PLAYBTN;
  const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Media ${target._slideObjects.filter((obj) => obj._type === "media" /* media */).length}`;
  const slideData = { _type: "media" /* media */ };
  if (!strPath && !strData && strType !== "online") {
    throw new Error("addMedia() error: either `data` or `path` are required!");
  } else if (strData && !strData.toLowerCase().includes("base64,")) {
    throw new Error("addMedia() error: `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')");
  } else if (strCover && !strCover.toLowerCase().includes("base64,")) {
    throw new Error("addMedia() error: `cover` value lacks a base64 header! Ex: 'data:image/png;base64,iV[...]')");
  }
  if (strType === "online" && !strLink) {
    throw new Error("addMedia() error: online videos require `link` value");
  }
  strExtn = opt.extn || (strData ? strData.split(";")[0].split("/")[1] : strPath.split(".").pop()) || "mp3";
  slideData.mtype = strType;
  slideData.media = strPath || "preencoded.mov";
  slideData.options = {};
  slideData.options.x = intPosX;
  slideData.options.y = intPosY;
  slideData.options.w = intSizeX;
  slideData.options.h = intSizeY;
  slideData.options.objectName = objectName;
  if (strType === "online") {
    const relId1 = getNewRelId(target);
    target._relsMedia.push({
      path: strPath || "preencoded" + strExtn,
      data: "dummy",
      type: "online",
      extn: strExtn,
      rId: relId1,
      Target: strLink
    });
    slideData.mediaRid = relId1;
    target._relsMedia.push({
      path: "preencoded.png",
      data: strCover,
      type: "image/png",
      extn: "png",
      rId: getNewRelId(target),
      Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`
    });
  } else {
    const dupeItem = target._relsMedia.filter((item) => item.path && item.path === strPath && item.type === strType + "/" + strExtn && !item.isDuplicate)[0];
    const relId1 = getNewRelId(target);
    target._relsMedia.push({
      path: strPath || "preencoded" + strExtn,
      type: strType + "/" + strExtn,
      extn: strExtn,
      data: strData || "",
      rId: relId1,
      isDuplicate: !!dupeItem?.Target,
      Target: dupeItem?.Target ? dupeItem.Target : `../media/media-${target._slideNum}-${target._relsMedia.length + 1}.${strExtn}`
    });
    slideData.mediaRid = relId1;
    target._relsMedia.push({
      path: strPath || "preencoded" + strExtn,
      type: strType + "/" + strExtn,
      extn: strExtn,
      data: strData || "",
      rId: getNewRelId(target),
      isDuplicate: !!dupeItem?.Target,
      Target: dupeItem?.Target ? dupeItem.Target : `../media/media-${target._slideNum}-${target._relsMedia.length + 0}.${strExtn}`
    });
    target._relsMedia.push({
      path: "preencoded.png",
      type: "image/png",
      extn: "png",
      data: strCover,
      rId: getNewRelId(target),
      Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`
    });
  }
  target._slideObjects.push(slideData);
}
function addNotesDefinition(target, notes) {
  target._slideObjects.push({
    _type: "notes" /* notes */,
    text: [{ text: notes }]
  });
}
function addShapeDefinition(target, shapeName, opts) {
  const options = typeof opts === "object" ? opts : {};
  options.line = options.line || { type: "none" };
  const newObject = {
    _type: "text" /* text */,
    shape: shapeName || "rect" /* RECTANGLE */,
    options,
    text: null
  };
  if (!shapeName) throw new Error("Missing/Invalid shape parameter! Example: `addShape(pptxgen.shapes.LINE, {x:1, y:1, w:1, h:1});`");
  const newLineOpts = {
    type: options.line.type || "solid",
    color: options.line.color || DEF_SHAPE_LINE_COLOR,
    transparency: options.line.transparency || 0,
    width: options.line.width || 1,
    dashType: options.line.dashType || "solid",
    beginArrowType: options.line.beginArrowType || null,
    endArrowType: options.line.endArrowType || null
  };
  if (typeof options.line === "object" && options.line.type !== "none") options.line = newLineOpts;
  options.x = options.x || (options.x === 0 ? 0 : 1);
  options.y = options.y || (options.y === 0 ? 0 : 1);
  options.w = options.w || (options.w === 0 ? 0 : 1);
  options.h = options.h || (options.h === 0 ? 0 : 1);
  options.objectName = options.objectName ? encodeXmlEntities(options.objectName) : `Shape ${target._slideObjects.filter((obj) => obj._type === "text" /* text */).length}`;
  if (typeof options.line === "string") {
    const tmpOpts = newLineOpts;
    tmpOpts.color = String(options.line);
    options.line = tmpOpts;
  }
  if (typeof options.lineSize === "number") options.line.width = options.lineSize;
  if (typeof options.lineDash === "string") options.line.dashType = options.lineDash;
  if (typeof options.lineHead === "string") options.line.beginArrowType = options.lineHead;
  if (typeof options.lineTail === "string") options.line.endArrowType = options.lineTail;
  createHyperlinkRels(target, newObject);
  target._slideObjects.push(newObject);
}
function addTableDefinition(target, tableRows, options, slideLayout, presLayout, addSlide, getSlide) {
  const slides = [target];
  const opt = options && typeof options === "object" ? options : {};
  opt.objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Table ${target._slideObjects.filter((obj) => obj._type === "table" /* table */).length}`;
  {
    if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
      throw new Error("addTable: Array expected! EX: 'slide.addTable( [rows], {options} );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)");
    }
    if (!tableRows[0] || !Array.isArray(tableRows[0])) {
      throw new Error(
        "addTable: 'rows' should be an array of cells! EX: 'slide.addTable( [ ['A'], ['B'], {text:'C',options:{align:'center'}} ] );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)"
      );
    }
  }
  const arrRows = [];
  tableRows.forEach((row) => {
    const newRow = [];
    if (Array.isArray(row)) {
      row.forEach((cell) => {
        const newCell = {
          _type: "tablecell" /* tablecell */,
          text: "",
          options: typeof cell === "object" && cell.options ? cell.options : {}
        };
        if (typeof cell === "string" || typeof cell === "number") newCell.text = cell.toString();
        else if (cell.text) {
          if (typeof cell.text === "string" || typeof cell.text === "number") newCell.text = cell.text.toString();
          else if (cell.text) newCell.text = cell.text;
          if (cell.options && typeof cell.options === "object") newCell.options = cell.options;
        }
        newCell.options.border = newCell.options.border || opt.border || [{ type: "none" }, { type: "none" }, { type: "none" }, { type: "none" }];
        const cellBorder = newCell.options.border;
        if (!Array.isArray(cellBorder) && typeof cellBorder === "object") newCell.options.border = [cellBorder, cellBorder, cellBorder, cellBorder];
        if (!newCell.options.border[0]) newCell.options.border[0] = { type: "none" };
        if (!newCell.options.border[1]) newCell.options.border[1] = { type: "none" };
        if (!newCell.options.border[2]) newCell.options.border[2] = { type: "none" };
        if (!newCell.options.border[3]) newCell.options.border[3] = { type: "none" };
        const arrSides = [0, 1, 2, 3];
        arrSides.forEach((idx) => {
          newCell.options.border[idx] = {
            type: newCell.options.border[idx].type || DEF_CELL_BORDER.type,
            color: newCell.options.border[idx].color || DEF_CELL_BORDER.color,
            pt: typeof newCell.options.border[idx].pt === "number" ? newCell.options.border[idx].pt : DEF_CELL_BORDER.pt
          };
        });
        newRow.push(newCell);
      });
    } else {
      console.log("addTable: tableRows has a bad row. A row should be an array of cells. You provided:");
      console.log(row);
    }
    arrRows.push(newRow);
  });
  opt.x = getSmartParseNumber(opt.x || (opt.x === 0 ? 0 : EMU / 2), "X", presLayout);
  opt.y = getSmartParseNumber(opt.y || (opt.y === 0 ? 0 : EMU / 2), "Y", presLayout);
  if (opt.h) opt.h = getSmartParseNumber(opt.h, "Y", presLayout);
  opt.fontSize = opt.fontSize || DEF_FONT_SIZE;
  opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_IN;
  if (typeof opt.margin === "number") opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)];
  if (!opt.color) opt.color = opt.color || DEF_FONT_COLOR;
  if (typeof opt.border === "string") {
    console.warn("addTable `border` option must be an object. Ex: `{border: {type:'none'}}`");
    opt.border = null;
  } else if (Array.isArray(opt.border)) {
    [0, 1, 2, 3].forEach((idx) => {
      opt.border[idx] = opt.border[idx] ? { type: opt.border[idx].type || DEF_CELL_BORDER.type, color: opt.border[idx].color || DEF_CELL_BORDER.color, pt: opt.border[idx].pt || DEF_CELL_BORDER.pt } : { type: "none" };
    });
  }
  opt.autoPage = typeof opt.autoPage === "boolean" ? opt.autoPage : false;
  opt.autoPageRepeatHeader = typeof opt.autoPageRepeatHeader === "boolean" ? opt.autoPageRepeatHeader : false;
  opt.autoPageHeaderRows = typeof opt.autoPageHeaderRows !== "undefined" && !isNaN(Number(opt.autoPageHeaderRows)) ? Number(opt.autoPageHeaderRows) : 1;
  opt.autoPageLineWeight = typeof opt.autoPageLineWeight !== "undefined" && !isNaN(Number(opt.autoPageLineWeight)) ? Number(opt.autoPageLineWeight) : 0;
  if (opt.autoPageLineWeight) {
    if (opt.autoPageLineWeight > 1) opt.autoPageLineWeight = 1;
    else if (opt.autoPageLineWeight < -1) opt.autoPageLineWeight = -1;
  }
  let arrTableMargin = DEF_SLIDE_MARGIN_IN;
  if (slideLayout && typeof slideLayout._margin !== "undefined") {
    if (Array.isArray(slideLayout._margin)) arrTableMargin = slideLayout._margin;
    else if (!isNaN(Number(slideLayout._margin))) {
      arrTableMargin = [Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin)];
    }
  }
  if (opt.colW) {
    const firstRowColCnt = arrRows[0].reduce((totalLen, c) => {
      if (c?.options?.colspan && typeof c.options.colspan === "number") {
        totalLen += c.options.colspan;
      } else {
        totalLen += 1;
      }
      return totalLen;
    }, 0);
    if (typeof opt.colW === "string" || typeof opt.colW === "number") {
      opt.w = Math.floor(Number(opt.colW) * firstRowColCnt);
      opt.colW = null;
    } else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length === 1 && firstRowColCnt > 1) {
      opt.w = Math.floor(Number(opt.colW) * firstRowColCnt);
      opt.colW = null;
    } else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== firstRowColCnt) {
      console.warn("addTable: mismatch: (colW.length != data.length) Therefore, defaulting to evenly distributed col widths.");
      opt.colW = null;
    }
  } else if (opt.w) {
    opt.w = getSmartParseNumber(opt.w, "X", presLayout);
  } else {
    opt.w = Math.floor(presLayout._sizeW / EMU - arrTableMargin[1] - arrTableMargin[3]);
  }
  if (opt.x && opt.x < 20) opt.x = inch2Emu(opt.x);
  if (opt.y && opt.y < 20) opt.y = inch2Emu(opt.y);
  if (opt.w && typeof opt.w === "number" && opt.w < 20) opt.w = inch2Emu(opt.w);
  if (opt.h && typeof opt.h === "number" && opt.h < 20) opt.h = inch2Emu(opt.h);
  arrRows.forEach((row) => {
    row.forEach((cell, idy) => {
      if (typeof cell === "number" || typeof cell === "string") {
        row[idy] = { _type: "tablecell" /* tablecell */, text: String(row[idy]), options: opt };
      } else if (typeof cell === "object") {
        if (typeof cell.text === "number") row[idy].text = row[idy].text.toString();
        else if (typeof cell.text === "undefined" || cell.text === null) row[idy].text = "";
        row[idy].options = cell.options || {};
        row[idy]._type = "tablecell" /* tablecell */;
      }
    });
  });
  const newAutoPagedSlides = [];
  if (opt && !opt.autoPage) {
    createHyperlinkRels(target, arrRows);
    target._slideObjects.push({
      _type: "table" /* table */,
      arrTabRows: arrRows,
      options: Object.assign({}, opt)
    });
  } else {
    if (opt.autoPageRepeatHeader) opt._arrObjTabHeadRows = arrRows.filter((_row, idx) => idx < opt.autoPageHeaderRows);
    getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach((slide, idx) => {
      if (!getSlide(target._slideNum + idx)) slides.push(addSlide({ masterName: slideLayout?._name || null }));
      if (idx > 0) opt.y = inch2Emu(opt.autoPageSlideStartY || opt.newSlideStartY || arrTableMargin[0]);
      {
        const newSlide = getSlide(target._slideNum + idx);
        opt.autoPage = false;
        createHyperlinkRels(newSlide, slide.rows);
        newSlide.addTable(slide.rows, Object.assign({}, opt));
        if (idx > 0) newAutoPagedSlides.push(newSlide);
      }
    });
  }
  return newAutoPagedSlides;
}
function addTextDefinition(target, text, opts, isPlaceholder) {
  const newObject = {
    _type: isPlaceholder ? "placeholder" /* placeholder */ : "text" /* text */,
    shape: opts?.shape || "rect" /* RECTANGLE */,
    text: !text || text.length === 0 ? [{ text: "", options: null }] : text,
    options: opts || {}
  };
  function cleanOpts(itemOpts) {
    {
      if (!itemOpts.placeholder) {
        itemOpts.color = itemOpts.color || newObject.options.color || target.color || DEF_FONT_COLOR;
      }
      if (itemOpts.placeholder || isPlaceholder) {
        itemOpts.bullet = itemOpts.bullet || false;
      }
      if (itemOpts.placeholder && target._slideLayout && target._slideLayout._slideObjects) {
        const placeHold = target._slideLayout._slideObjects.filter(
          (item) => item._type === "placeholder" && item.options && item.options.placeholder && item.options.placeholder === itemOpts.placeholder
        )[0];
        if (placeHold?.options) itemOpts = { ...itemOpts, ...placeHold.options };
      }
      itemOpts.objectName = itemOpts.objectName ? encodeXmlEntities(itemOpts.objectName) : `Text ${target._slideObjects.filter((obj) => obj._type === "text" /* text */).length}`;
      if (itemOpts.shape === "line" /* LINE */) {
        const newLineOpts = {
          type: itemOpts.line.type || "solid",
          color: itemOpts.line.color || DEF_SHAPE_LINE_COLOR,
          transparency: itemOpts.line.transparency || 0,
          width: itemOpts.line.width || 1,
          dashType: itemOpts.line.dashType || "solid",
          beginArrowType: itemOpts.line.beginArrowType || null,
          endArrowType: itemOpts.line.endArrowType || null
        };
        if (typeof itemOpts.line === "object") itemOpts.line = newLineOpts;
        if (typeof itemOpts.line === "string") {
          const tmpOpts = newLineOpts;
          if (typeof itemOpts.line === "string") tmpOpts.color = itemOpts.line;
          itemOpts.line = tmpOpts;
        }
        if (typeof itemOpts.lineSize === "number") itemOpts.line.width = itemOpts.lineSize;
        if (typeof itemOpts.lineDash === "string") itemOpts.line.dashType = itemOpts.lineDash;
        if (typeof itemOpts.lineHead === "string") itemOpts.line.beginArrowType = itemOpts.lineHead;
        if (typeof itemOpts.lineTail === "string") itemOpts.line.endArrowType = itemOpts.lineTail;
      }
      itemOpts.line = itemOpts.line || {};
      itemOpts.lineSpacing = itemOpts.lineSpacing && !isNaN(itemOpts.lineSpacing) ? itemOpts.lineSpacing : null;
      itemOpts.lineSpacingMultiple = itemOpts.lineSpacingMultiple && !isNaN(itemOpts.lineSpacingMultiple) ? itemOpts.lineSpacingMultiple : null;
      itemOpts._bodyProp = itemOpts._bodyProp || {};
      itemOpts._bodyProp.autoFit = itemOpts.autoFit || false;
      itemOpts._bodyProp.anchor = !itemOpts.placeholder ? "ctr" /* ctr */ : null;
      itemOpts._bodyProp.vert = itemOpts.vert || null;
      itemOpts._bodyProp.wrap = typeof itemOpts.wrap === "boolean" ? itemOpts.wrap : true;
      if (itemOpts.inset && !isNaN(Number(itemOpts.inset)) || itemOpts.inset === 0) {
        itemOpts._bodyProp.lIns = inch2Emu(itemOpts.inset);
        itemOpts._bodyProp.rIns = inch2Emu(itemOpts.inset);
        itemOpts._bodyProp.tIns = inch2Emu(itemOpts.inset);
        itemOpts._bodyProp.bIns = inch2Emu(itemOpts.inset);
      }
      if (typeof itemOpts.underline === "boolean" && itemOpts.underline === true) itemOpts.underline = { style: "sng" };
    }
    {
      if ((itemOpts.align || "").toLowerCase().indexOf("c") === 0) itemOpts._bodyProp.align = "center" /* center */;
      else if ((itemOpts.align || "").toLowerCase().indexOf("l") === 0) itemOpts._bodyProp.align = "left" /* left */;
      else if ((itemOpts.align || "").toLowerCase().indexOf("r") === 0) itemOpts._bodyProp.align = "right" /* right */;
      else if ((itemOpts.align || "").toLowerCase().indexOf("j") === 0) itemOpts._bodyProp.align = "justify" /* justify */;
      if ((itemOpts.valign || "").toLowerCase().indexOf("b") === 0) itemOpts._bodyProp.anchor = "b" /* b */;
      else if ((itemOpts.valign || "").toLowerCase().indexOf("m") === 0) itemOpts._bodyProp.anchor = "ctr" /* ctr */;
      else if ((itemOpts.valign || "").toLowerCase().indexOf("t") === 0) itemOpts._bodyProp.anchor = "t" /* t */;
    }
    correctShadowOptions(itemOpts.shadow);
    return itemOpts;
  }
  newObject.options = cleanOpts(newObject.options);
  newObject.text.forEach((item) => item.options = cleanOpts(item.options || {}));
  createHyperlinkRels(target, newObject.text || "");
  target._slideObjects.push(newObject);
}
function addPlaceholdersToSlideLayouts(slide) {
  (slide._slideLayout._slideObjects || []).forEach((slideLayoutObj) => {
    if (slideLayoutObj._type === "placeholder" /* placeholder */) {
      if (slide._slideObjects.filter((slideObj) => slideObj.options && slideObj.options.placeholder === slideLayoutObj.options.placeholder).length === 0) {
        addTextDefinition(slide, [{ text: "" }], slideLayoutObj.options, false);
      }
    }
  });
}
function addBackgroundDefinition(props, target) {
  if (target.bkgd) {
    if (!target.background) target.background = {};
    if (typeof target.bkgd === "string") target.background.color = target.bkgd;
    else {
      if (target.bkgd.data) target.background.data = target.bkgd.data;
      if (target.bkgd.path) target.background.path = target.bkgd.path;
      if (target.bkgd.src) target.background.path = target.bkgd.src;
    }
  }
  if (target.background?.fill) target.background.color = target.background.fill;
  if (props && (props.path || props.data)) {
    props.path = props.path || "preencoded.png";
    let strImgExtn = (props.path.split(".").pop() || "png").split("?")[0];
    if (strImgExtn === "jpg") strImgExtn = "jpeg";
    target._relsMedia = target._relsMedia || [];
    const intRels = target._relsMedia.length + 1;
    target._relsMedia.push({
      path: props.path,
      type: "image" /* image */,
      extn: strImgExtn,
      data: props.data || null,
      rId: intRels,
      Target: `../media/${(target._name || "").replace(/\s+/gi, "-")}-image-${target._relsMedia.length + 1}.${strImgExtn}`
    });
    target._bkgdImgRid = intRels;
  }
}
function createHyperlinkRels(target, text) {
  let textObjs = [];
  if (typeof text === "string" || typeof text === "number") return;
  else if (Array.isArray(text)) textObjs = text;
  else if (typeof text === "object") textObjs = [text];
  textObjs.forEach((text2) => {
    if (Array.isArray(text2)) {
      createHyperlinkRels(target, text2);
    } else if (Array.isArray(text2.text)) {
      createHyperlinkRels(target, text2.text);
    } else if (text2 && typeof text2 === "object" && text2.options && text2.options.hyperlink && !text2.options.hyperlink._rId) {
      if (typeof text2.options.hyperlink !== "object") console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ");
      else if (!text2.options.hyperlink.url && !text2.options.hyperlink.slide) console.log("ERROR: 'hyperlink requires either: `url` or `slide`'");
      else {
        const relId = getNewRelId(target);
        target._rels.push({
          type: "hyperlink" /* hyperlink */,
          data: text2.options.hyperlink.slide ? "slide" : "dummy",
          rId: relId,
          Target: encodeXmlEntities(text2.options.hyperlink.url) || text2.options.hyperlink.slide.toString()
        });
        text2.options.hyperlink._rId = relId;
      }
    }
  });
}

// vendor/pptxgenjs/src/slide.ts
var Slide = class {
  _setSlideNum;
  addSlide;
  getSlide;
  _name;
  _presLayout;
  _rels;
  _relsChart;
  _relsMedia;
  _rId;
  _slideId;
  _slideLayout;
  _slideNum;
  _slideNumberProps;
  _slideObjects;
  _newAutoPagedSlides;
  constructor(params) {
    this.addSlide = params.addSlide;
    this.getSlide = params.getSlide;
    this._name = `Slide ${params.slideNumber}`;
    this._presLayout = params.presLayout;
    this._rId = params.slideRId;
    this._rels = [];
    this._relsChart = [];
    this._relsMedia = [];
    this._setSlideNum = params.setSlideNum;
    this._slideId = params.slideId;
    this._slideLayout = params.slideLayout || null;
    this._slideNum = params.slideNumber;
    this._slideObjects = [];
    this._slideNumberProps = this._slideLayout?._slideNumberProps ? this._slideLayout._slideNumberProps : null;
  }
  /**
   * Background color
   * @type {string|BackgroundProps}
   * @deprecated in v3.3.0 - use `background` instead
   */
  _bkgd;
  set bkgd(value) {
    this._bkgd = value;
    if (!this._background || !this._background.color) {
      if (!this._background) this._background = {};
      if (typeof value === "string") this._background.color = value;
    }
  }
  get bkgd() {
    return this._bkgd;
  }
  /**
   * Background color or image
   * @type {BackgroundProps}
   * @example solid color `background: { color:'FF0000' }`
   * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
   * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
   * @example url `background: { path:'https://some.url/image.jpg'}`
   * @since v3.3.0
   */
  _background;
  set background(props) {
    this._background = props;
    if (props) addBackgroundDefinition(props, this);
  }
  get background() {
    return this._background;
  }
  /**
   * Default font color
   * @type {HexColor}
   */
  _color;
  set color(value) {
    this._color = value;
  }
  get color() {
    return this._color;
  }
  /**
   * @type {boolean}
   */
  _hidden;
  set hidden(value) {
    this._hidden = value;
  }
  get hidden() {
    return this._hidden;
  }
  /**
   * @type {SlideNumberProps}
   */
  set slideNumber(value) {
    this._slideNumberProps = value;
    this._setSlideNum(value);
  }
  get slideNumber() {
    return this._slideNumberProps;
  }
  get newAutoPagedSlides() {
    return this._newAutoPagedSlides;
  }
  /**
   * Add chart to Slide
   * @param {CHART_NAME|IChartMulti[]} type - chart type
   * @param {object[]} data - data object
   * @param {IChartOpts} options - chart options
   * @return {Slide} this Slide
   */
  addChart(type, data, options) {
    const optionsWithType = options || {};
    optionsWithType._type = type;
    addChartDefinition(this, type, data, options);
    return this;
  }
  /**
   * Add image to Slide
   * @param {ImageProps} options - image options
   * @return {Slide} this Slide
   */
  addImage(options) {
    addImageDefinition(this, options);
    return this;
  }
  /**
   * Add media (audio/video) to Slide
   * @param {MediaProps} options - media options
   * @return {Slide} this Slide
   */
  addMedia(options) {
    addMediaDefinition(this, options);
    return this;
  }
  /**
   * Add speaker notes to Slide
   * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
   * @param {string} notes - notes to add to slide
   * @return {Slide} this Slide
   */
  addNotes(notes) {
    addNotesDefinition(this, notes);
    return this;
  }
  /**
   * Add shape to Slide
   * @param {SHAPE_NAME} shapeName - shape name
   * @param {ShapeProps} options - shape options
   * @return {Slide} this Slide
   */
  addShape(shapeName, options) {
    addShapeDefinition(this, shapeName, options);
    return this;
  }
  /**
   * Add table to Slide
   * @param {TableRow[]} tableRows - table rows
   * @param {TableProps} options - table options
   * @return {Slide} this Slide
   */
  addTable(tableRows, options) {
    this._newAutoPagedSlides = addTableDefinition(this, tableRows, options, this._slideLayout, this._presLayout, this.addSlide, this.getSlide);
    return this;
  }
  /**
   * Add text to Slide
   * @param {string|TextProps[]} text - text string or complex object
   * @param {TextPropsOptions} options - text options
   * @return {Slide} this Slide
   */
  addText(text, options) {
    const textParam = typeof text === "string" || typeof text === "number" ? [{ text, options }] : text;
    addTextDefinition(this, textParam, options, false);
    return this;
  }
};

// vendor/pptxgenjs/src/gen-charts.ts
import JSZip from "jszip";
async function createExcelWorksheet(chartObject, zip) {
  const data = chartObject.data;
  return await new Promise((resolve3, reject) => {
    const zipExcel = new JSZip();
    const intBubbleCols = (data.length - 1) * 2 + 1;
    const IS_MULTI_CAT_AXES = data[0]?.labels?.length > 1;
    zipExcel.folder("_rels");
    zipExcel.folder("docProps");
    zipExcel.folder("xl/_rels");
    zipExcel.folder("xl/tables");
    zipExcel.folder("xl/theme");
    zipExcel.folder("xl/worksheets");
    zipExcel.folder("xl/worksheets/_rels");
    {
      zipExcel.file(
        "[Content_Types].xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>  <Default Extension="xml" ContentType="application/xml"/>  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>\n'
      );
      zipExcel.file(
        "_rels/.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>\n'
      );
      zipExcel.file(
        "docProps/app.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Macintosh Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>\n'
      );
      zipExcel.file(
        "docProps/core.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>PptxGenJS</dc:creator><cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">' + (/* @__PURE__ */ new Date()).toISOString() + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + (/* @__PURE__ */ new Date()).toISOString() + "</dcterms:modified></cp:coreProperties>"
      );
      zipExcel.file(
        "xl/_rels/workbook.xml.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
      );
      zipExcel.file(
        "xl/styles.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/><rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
      );
      zipExcel.file(
        "xl/theme/theme1.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="\uB9D1\uC740 \uACE0\uB515"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="\u65B0\u7D30\u660E\u9AD4"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="\uB9D1\uC740 \uACE0\uB515"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="\u65B0\u7D30\u660E\u9AD4"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'
      );
      zipExcel.file(
        "xl/workbook.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="10507"/><workbookPr/><bookViews><workbookView xWindow="0" yWindow="500" windowWidth="20960" windowHeight="15960"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0" concurrentCalc="0"/></workbook>\n'
      );
      zipExcel.file(
        "xl/worksheets/_rels/sheet1.xml.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>\n'
      );
    }
    {
      let strSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      if (chartObject.opts._type === "bubble" /* BUBBLE */ || chartObject.opts._type === "bubble3D" /* BUBBLE3D */) {
        strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${intBubbleCols}" uniqueCount="${intBubbleCols}">`;
      } else if (chartObject.opts._type === "scatter" /* SCATTER */) {
        strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${data.length}" uniqueCount="${data.length}">`;
      } else if (IS_MULTI_CAT_AXES) {
        let totCount = data.length;
        data[0].labels.forEach((arrLabel) => totCount += arrLabel.filter((label) => label && label !== "").length);
        strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totCount}" uniqueCount="${totCount}">`;
        strSharedStrings += "<si><t/></si>";
      } else {
        const totCount = data.length + data[0].labels.length * data[0].labels[0].length + data[0].labels.length;
        const unqCount = data.length + data[0].labels.length * data[0].labels[0].length + 1;
        strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totCount}" uniqueCount="${unqCount}">`;
        strSharedStrings += '<si><t xml:space="preserve"></t></si>';
      }
      if (chartObject.opts._type === "bubble" /* BUBBLE */ || chartObject.opts._type === "bubble3D" /* BUBBLE3D */) {
        data.forEach((objData, idx) => {
          if (idx === 0) strSharedStrings += "<si><t>X-Axis</t></si>";
          else {
            strSharedStrings += `<si><t>${encodeXmlEntities(objData.name || `Y-Axis${idx}`)}</t></si>`;
            strSharedStrings += `<si><t>${encodeXmlEntities(`Size${idx}`)}</t></si>`;
          }
        });
      } else {
        data.forEach((objData) => {
          strSharedStrings += `<si><t>${encodeXmlEntities((objData.name || " ").replace("X-Axis", "X-Values"))}</t></si>`;
        });
      }
      if (chartObject.opts._type !== "bubble" /* BUBBLE */ && chartObject.opts._type !== "bubble3D" /* BUBBLE3D */ && chartObject.opts._type !== "scatter" /* SCATTER */) {
        data[0].labels.slice().reverse().forEach((labelsGroup) => {
          labelsGroup.filter((label) => label && label !== "").forEach((label) => {
            strSharedStrings += `<si><t>${encodeXmlEntities(label)}</t></si>`;
          });
        });
      }
      strSharedStrings += "</sst>\n";
      zipExcel.file("xl/sharedStrings.xml", strSharedStrings);
    }
    {
      let strTableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      if (chartObject.opts._type === "bubble" /* BUBBLE */ || chartObject.opts._type === "bubble3D" /* BUBBLE3D */) {
        strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(intBubbleCols)}${intBubbleCols}" totalsRowShown="0">`;
        strTableXml += `<tableColumns count="${intBubbleCols}">`;
        let idxColLtr = 1;
        data.forEach((obj, idx) => {
          if (idx === 0) {
            strTableXml += `<tableColumn id="${idx + 1}" name="X-Values"/>`;
          } else {
            strTableXml += `<tableColumn id="${idx + idxColLtr}" name="${obj.name}"/>`;
            idxColLtr++;
            strTableXml += `<tableColumn id="${idx + idxColLtr}" name="Size${idx}"/>`;
          }
        });
      } else if (chartObject.opts._type === "scatter" /* SCATTER */) {
        strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(data.length)}${data[0].values.length + 1}" totalsRowShown="0">`;
        strTableXml += `<tableColumns count="${data.length}">`;
        data.forEach((_obj, idx) => {
          strTableXml += `<tableColumn id="${idx + 1}" name="${idx === 0 ? "X-Values" : "Y-Value "}${idx}"/>`;
        });
      } else {
        strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(data.length + data[0].labels.length)}${data[0].labels[0].length + 1}'" totalsRowShown="0">`;
        strTableXml += `<tableColumns count="${data.length + data[0].labels.length}">`;
        data[0].labels.forEach((_labelsGroup, idx) => {
          strTableXml += `<tableColumn id="${idx + 1}" name="Column${idx + 1}"/>`;
        });
        data.forEach((obj, idx) => {
          strTableXml += `<tableColumn id="${idx + data[0].labels.length + 1}" name="${encodeXmlEntities(obj.name)}"/>`;
        });
      }
      strTableXml += "</tableColumns>";
      strTableXml += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>';
      strTableXml += "</table>";
      zipExcel.file("xl/tables/table1.xml", strTableXml);
    }
    {
      let strSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      strSheetXml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
      if (chartObject.opts._type === "bubble" /* BUBBLE */ || chartObject.opts._type === "bubble3D" /* BUBBLE3D */) {
        strSheetXml += `<dimension ref="A1:${getExcelColName(intBubbleCols)}${data[0].values.length + 1}"/>`;
      } else if (chartObject.opts._type === "scatter" /* SCATTER */) {
        strSheetXml += `<dimension ref="A1:${getExcelColName(data.length)}${data[0].values.length + 1}"/>`;
      } else {
        strSheetXml += `<dimension ref="A1:${getExcelColName(data.length + 1)}${data[0].values.length + 1}"/>`;
      }
      strSheetXml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>';
      strSheetXml += '<sheetFormatPr baseColWidth="10" defaultRowHeight="16"/>';
      if (chartObject.opts._type === "bubble" /* BUBBLE */ || chartObject.opts._type === "bubble3D" /* BUBBLE3D */) {
        strSheetXml += "<sheetData>";
        strSheetXml += `<row r="1" spans="1:${intBubbleCols}">`;
        strSheetXml += '<c r="A1" t="s"><v>0</v></c>';
        for (let idx = 1; idx < intBubbleCols; idx++) {
          strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>${idx}</v></c>`;
        }
        strSheetXml += "</row>";
        data[0].values.forEach((val, idx) => {
          strSheetXml += `<row r="${idx + 2}" spans="1:${intBubbleCols}">`;
          strSheetXml += `<c r="A${idx + 2}"><v>${val}</v></c>`;
          let idxColLtr = 2;
          for (let idy = 1; idy < data.length; idy++) {
            strSheetXml += `<c r="${getExcelColName(idxColLtr)}${idx + 2}"><v>${data[idy].values[idx] || ""}</v></c>`;
            idxColLtr++;
            strSheetXml += `<c r="${getExcelColName(idxColLtr)}${idx + 2}"><v>${data[idy].sizes[idx] || ""}</v></c>`;
            idxColLtr++;
          }
          strSheetXml += "</row>";
        });
      } else if (chartObject.opts._type === "scatter" /* SCATTER */) {
        strSheetXml += "<sheetData>";
        strSheetXml += `<row r="1" spans="1:${data.length}">`;
        for (let idx = 0; idx < data.length; idx++) {
          strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>${idx}</v></c>`;
        }
        strSheetXml += "</row>";
        data[0].values.forEach((val, idx) => {
          strSheetXml += `<row r="${idx + 2}" spans="1:${data.length}">`;
          strSheetXml += `<c r="A${idx + 2}"><v>${val}</v></c>`;
          for (let idy = 1; idy < data.length; idy++) {
            strSheetXml += `<c r="${getExcelColName(idy + 1)}${idx + 2}"><v>${data[idy].values[idx] || data[idy].values[idx] === 0 ? data[idy].values[idx] : ""}</v></c>`;
          }
          strSheetXml += "</row>";
        });
      } else {
        strSheetXml += "<sheetData>";
        if (!IS_MULTI_CAT_AXES) {
          strSheetXml += `<row r="1" spans="1:${data.length + data[0].labels.length}">`;
          data[0].labels.forEach((_labelsGroup, idx) => {
            strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>0</v></c>`;
          });
          for (let idx = 0; idx < data.length; idx++) {
            strSheetXml += `<c r="${getExcelColName(idx + 1 + data[0].labels.length)}1" t="s"><v>${idx + 1}</v></c>`;
          }
          strSheetXml += "</row>";
          data[0].labels[0].forEach((_cat, idx) => {
            strSheetXml += `<row r="${idx + 2}" spans="1:${data.length + data[0].labels.length}">`;
            for (let idx2 = data[0].labels.length - 1; idx2 >= 0; idx2--) {
              strSheetXml += `<c r="${getExcelColName(data[0].labels.length - idx2)}${idx + 2}" t="s">`;
              strSheetXml += `<v>${data.length + idx + 1}</v>`;
              strSheetXml += "</c>";
            }
            for (let idy = 0; idy < data.length; idy++) {
              strSheetXml += `<c r="${getExcelColName(data[0].labels.length + idy + 1)}${idx + 2}"><v>${data[idy].values[idx] || ""}</v></c>`;
            }
            strSheetXml += "</row>";
          });
        } else {
          strSheetXml += `<row r="1" spans="1:${data.length + data[0].labels.length}">`;
          for (let idx = 0; idx < data[0].labels.length; idx++) {
            strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>0</v></c>`;
          }
          for (let idx = data[0].labels.length - 1; idx < data.length + data[0].labels.length - 1; idx++) {
            strSheetXml += `<c r="${getExcelColName(idx + data[0].labels.length)}1" t="s"><v>${idx}</v></c>`;
          }
          strSheetXml += "</row>";
          const TOT_SER = data.length;
          const TOT_CAT = data[0].labels[0].length;
          const TOT_LVL = data[0].labels.length;
          for (let idx = 0; idx < TOT_CAT; idx++) {
            strSheetXml += `<row r="${idx + 2}" spans="1:${TOT_SER + TOT_LVL}">`;
            let totLabels = TOT_SER;
            const revLabelGroups = data[0].labels.slice().reverse();
            revLabelGroups.forEach((labelsGroup, idy) => {
              const colLabel = labelsGroup[idx];
              if (colLabel) {
                const totGrpLbls = idy === 0 ? 1 : revLabelGroups[idy - 1].filter((label) => label && label !== "").length;
                totLabels += totGrpLbls;
                strSheetXml += `<c r="${getExcelColName(idx + 1 + idy)}${idx + 2}" t="s"><v>${totLabels}</v></c>`;
              }
            });
            for (let idy = 0; idy < TOT_SER; idy++) {
              strSheetXml += `<c r="${getExcelColName(TOT_LVL + idy + 1)}${idx + 2}"><v>${data[idy].values[idx] || 0}</v></c>`;
            }
            strSheetXml += "</row>";
          }
        }
      }
      strSheetXml += "</sheetData>";
      strSheetXml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      strSheetXml += "</worksheet>\n";
      zipExcel.file("xl/worksheets/sheet1.xml", strSheetXml);
    }
    zipExcel.generateAsync({ type: "base64" }).then((content) => {
      zip.file(`ppt/embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx`, content, { base64: true });
      zip.file(
        "ppt/charts/_rels/" + chartObject.fileName + ".rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx"/></Relationships>`
      );
      zip.file(`ppt/charts/${chartObject.fileName}`, makeXmlCharts(chartObject));
      resolve3("");
    }).catch((strErr) => {
      reject(strErr);
    });
  });
}
function makeXmlCharts(rel) {
  let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  let usesSecondaryValAxis = false;
  {
    strXml += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    strXml += '<c:date1904 val="0"/>';
    strXml += `<c:roundedCorners val="${rel.opts.chartArea.roundedCorners ? "1" : "0"}"/>`;
    strXml += "<c:chart>";
    if (rel.opts.showTitle) {
      strXml += genXmlTitle(
        {
          title: rel.opts.title || "Chart Title",
          color: rel.opts.titleColor,
          fontFace: rel.opts.titleFontFace,
          fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
          titleAlign: rel.opts.titleAlign,
          titleBold: rel.opts.titleBold,
          titlePos: rel.opts.titlePos,
          titleRotate: rel.opts.titleRotate
        },
        rel.opts.x,
        rel.opts.y
      );
      strXml += '<c:autoTitleDeleted val="0"/>';
    } else {
      strXml += '<c:autoTitleDeleted val="1"/>';
    }
    if (rel.opts._type === "bar3D" /* BAR3D */) {
      strXml += `<c:view3D><c:rotX val="${rel.opts.v3DRotX}"/><c:rotY val="${rel.opts.v3DRotY}"/><c:rAngAx val="${!rel.opts.v3DRAngAx ? 0 : 1}"/><c:perspective val="${rel.opts.v3DPerspective}"/></c:view3D>`;
    }
    strXml += "<c:plotArea>";
    if (rel.opts.layout) {
      strXml += "<c:layout>";
      strXml += " <c:manualLayout>";
      strXml += '  <c:layoutTarget val="inner" />';
      strXml += '  <c:xMode val="edge" />';
      strXml += '  <c:yMode val="edge" />';
      strXml += '  <c:x val="' + (rel.opts.layout.x || 0) + '" />';
      strXml += '  <c:y val="' + (rel.opts.layout.y || 0) + '" />';
      strXml += '  <c:w val="' + (rel.opts.layout.w || 1) + '" />';
      strXml += '  <c:h val="' + (rel.opts.layout.h || 1) + '" />';
      strXml += " </c:manualLayout>";
      strXml += "</c:layout>";
    } else {
      strXml += "<c:layout/>";
    }
  }
  if (Array.isArray(rel.opts._type)) {
    rel.opts._type.forEach((type) => {
      const options = { ...rel.opts, ...type.options };
      const valAxisId = options.secondaryValAxis ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY;
      const catAxisId = options.secondaryCatAxis ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY;
      usesSecondaryValAxis = usesSecondaryValAxis || options.secondaryValAxis;
      strXml += makeChartType(type.type, type.data, options, valAxisId, catAxisId, true);
    });
  } else {
    strXml += makeChartType(rel.opts._type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY, false);
  }
  if (rel.opts._type !== "pie" /* PIE */ && rel.opts._type !== "doughnut" /* DOUGHNUT */) {
    if (rel.opts.valAxes && rel.opts.valAxes.length > 1 && !usesSecondaryValAxis) {
      throw new Error("Secondary axis must be used by one of the multiple charts");
    }
    if (rel.opts.catAxes) {
      if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
        throw new Error("There must be the same number of value and category axes.");
      }
      strXml += makeCatAxis({ ...rel.opts, ...rel.opts.catAxes[0] }, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
    } else {
      strXml += makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
    }
    if (rel.opts.valAxes) {
      strXml += makeValAxis({ ...rel.opts, ...rel.opts.valAxes[0] }, AXIS_ID_VALUE_PRIMARY);
      if (rel.opts.valAxes[1]) {
        strXml += makeValAxis({ ...rel.opts, ...rel.opts.valAxes[1] }, AXIS_ID_VALUE_SECONDARY);
      }
    } else {
      strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY);
      if (rel.opts._type === "bar3D" /* BAR3D */) {
        strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY);
      }
    }
    if (rel.opts?.catAxes && rel.opts?.catAxes[1]) {
      strXml += makeCatAxis({ ...rel.opts, ...rel.opts.catAxes[1] }, AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_SECONDARY);
    }
  }
  {
    if (rel.opts.showDataTable) {
      strXml += "<c:dTable>";
      strXml += `  <c:showHorzBorder val="${!rel.opts.showDataTableHorzBorder ? 0 : 1}"/>`;
      strXml += `  <c:showVertBorder val="${!rel.opts.showDataTableVertBorder ? 0 : 1}"/>`;
      strXml += `  <c:showOutline    val="${!rel.opts.showDataTableOutline ? 0 : 1}"/>`;
      strXml += `  <c:showKeys       val="${!rel.opts.showDataTableKeys ? 0 : 1}"/>`;
      strXml += "  <c:spPr>";
      strXml += "    <a:noFill/>";
      strXml += '    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>';
      strXml += "    <a:effectLst/>";
      strXml += "  </c:spPr>";
      strXml += "  <c:txPr>";
      strXml += '   <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>';
      strXml += "   <a:lstStyle/>";
      strXml += "   <a:p>";
      strXml += '     <a:pPr rtl="0">';
      strXml += `       <a:defRPr sz="${Math.round((rel.opts.dataTableFontSize || DEF_FONT_SIZE) * 100)}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">`;
      strXml += '         <a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>';
      strXml += '         <a:latin typeface="+mn-lt"/>';
      strXml += '         <a:ea typeface="+mn-ea"/>';
      strXml += '         <a:cs typeface="+mn-cs"/>';
      strXml += "       </a:defRPr>";
      strXml += "     </a:pPr>";
      strXml += '    <a:endParaRPr lang="en-US"/>';
      strXml += "   </a:p>";
      strXml += " </c:txPr>";
      strXml += "</c:dTable>";
    }
    strXml += "  <c:spPr>";
    strXml += rel.opts.plotArea.fill?.color ? genXmlColorSelection(rel.opts.plotArea.fill) : "<a:noFill/>";
    strXml += rel.opts.plotArea.border ? `<a:ln w="${valToPts(rel.opts.plotArea.border.pt)}" cap="flat">${genXmlColorSelection(rel.opts.plotArea.border.color)}</a:ln>` : "<a:ln><a:noFill/></a:ln>";
    strXml += "    <a:effectLst/>";
    strXml += "  </c:spPr>";
    strXml += "</c:plotArea>";
    if (rel.opts.showLegend) {
      strXml += "<c:legend>";
      strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>';
      strXml += '<c:overlay val="0"/>';
      if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
        strXml += "<c:txPr>";
        strXml += "  <a:bodyPr/>";
        strXml += "  <a:lstStyle/>";
        strXml += "  <a:p>";
        strXml += "    <a:pPr>";
        strXml += rel.opts.legendFontSize ? `<a:defRPr sz="${Math.round(Number(rel.opts.legendFontSize) * 100)}">` : "<a:defRPr>";
        if (rel.opts.legendColor) strXml += genXmlColorSelection(rel.opts.legendColor);
        if (rel.opts.legendFontFace) strXml += '<a:latin typeface="' + rel.opts.legendFontFace + '"/>';
        if (rel.opts.legendFontFace) strXml += '<a:cs    typeface="' + rel.opts.legendFontFace + '"/>';
        strXml += "      </a:defRPr>";
        strXml += "    </a:pPr>";
        strXml += '    <a:endParaRPr lang="en-US"/>';
        strXml += "  </a:p>";
        strXml += "</c:txPr>";
      }
      strXml += "</c:legend>";
    }
  }
  strXml += '  <c:plotVisOnly val="1"/>';
  strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>';
  if (rel.opts._type === "scatter" /* SCATTER */) strXml += '<c:showDLblsOverMax val="1"/>';
  strXml += "</c:chart>";
  strXml += "<c:spPr>";
  strXml += rel.opts.chartArea.fill?.color ? genXmlColorSelection(rel.opts.chartArea.fill) : "<a:noFill/>";
  strXml += rel.opts.chartArea.border ? `<a:ln w="${valToPts(rel.opts.chartArea.border.pt)}" cap="flat">${genXmlColorSelection(rel.opts.chartArea.border.color)}</a:ln>` : "<a:ln><a:noFill/></a:ln>";
  strXml += "  <a:effectLst/>";
  strXml += "</c:spPr>";
  strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>';
  strXml += "</c:chartSpace>";
  return strXml;
}
function makeChartType(chartType, data, opts, valAxisId, catAxisId, isMultiTypeChart) {
  let colorIndex = -1;
  let idxColLtr = 1;
  let optsChartData = null;
  let strXml = "";
  switch (chartType) {
    case "area" /* AREA */:
    case "bar" /* BAR */:
    case "bar3D" /* BAR3D */:
    case "line" /* LINE */:
    case "radar" /* RADAR */:
      strXml += `<c:${chartType}Chart>`;
      if (chartType === "area" /* AREA */ && opts.barGrouping === "stacked") {
        strXml += '<c:grouping val="' + opts.barGrouping + '"/>';
      }
      if (chartType === "bar" /* BAR */ || chartType === "bar3D" /* BAR3D */) {
        strXml += '<c:barDir val="' + opts.barDir + '"/>';
        strXml += '<c:grouping val="' + (opts.barGrouping || "clustered") + '"/>';
      }
      if (chartType === "radar" /* RADAR */) {
        strXml += '<c:radarStyle val="' + opts.radarStyle + '"/>';
      }
      strXml += '<c:varyColors val="0"/>';
      data.forEach((obj) => {
        colorIndex++;
        strXml += "<c:ser>";
        strXml += `  <c:idx val="${obj._dataIndex}"/><c:order val="${obj._dataIndex}"/>`;
        strXml += "  <c:tx>";
        strXml += "    <c:strRef>";
        strXml += "      <c:f>Sheet1!$" + getExcelColName(obj._dataIndex + obj.labels.length + 1) + "$1</c:f>";
        strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + "</c:v></c:pt></c:strCache>";
        strXml += "    </c:strRef>";
        strXml += "  </c:tx>";
        const seriesColor = opts.chartColors ? opts.chartColors[colorIndex % opts.chartColors.length] : null;
        strXml += "  <c:spPr>";
        if (seriesColor === "transparent") {
          strXml += "<a:noFill/>";
        } else if (opts.chartColorsOpacity) {
          strXml += "<a:solidFill>" + createColorElement(seriesColor, `<a:alpha val="${Math.round(opts.chartColorsOpacity * 1e3)}"/>`) + "</a:solidFill>";
        } else {
          strXml += "<a:solidFill>" + createColorElement(seriesColor) + "</a:solidFill>";
        }
        if (chartType === "line" /* LINE */ || chartType === "radar" /* RADAR */) {
          if (opts.lineSize === 0) {
            strXml += "<a:ln><a:noFill/></a:ln>";
          } else {
            strXml += `<a:ln w="${valToPts(opts.lineSize)}" cap="${createLineCap(opts.lineCap)}"><a:solidFill>${createColorElement(seriesColor)}</a:solidFill>`;
            strXml += '<a:prstDash val="' + (opts.lineDash || "solid") + '"/><a:round/></a:ln>';
          }
        } else if (opts.dataBorder) {
          strXml += `<a:ln w="${valToPts(opts.dataBorder.pt)}" cap="${createLineCap(opts.lineCap)}"><a:solidFill>${createColorElement(opts.dataBorder.color)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`;
        }
        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
        strXml += "  </c:spPr>";
        strXml += '  <c:invertIfNegative val="0"/>';
        if (chartType !== "radar" /* RADAR */) {
          strXml += "<c:dLbls>";
          strXml += `<c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || "General"}" sourceLinked="0"/>`;
          if (opts.dataLabelBkgrdColors) strXml += `<c:spPr><a:solidFill>${createColorElement(seriesColor)}</a:solidFill></c:spPr>`;
          strXml += "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>";
          strXml += `<a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${Math.round(
            (opts.dataLabelFontSize || DEF_FONT_SIZE) * 100
          )}" u="none">`;
          strXml += `<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>`;
          strXml += `<a:latin typeface="${opts.dataLabelFontFace || "Arial"}"/>`;
          strXml += "</a:defRPr></a:pPr></a:p></c:txPr>";
          if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`;
          strXml += '<c:showLegendKey val="0"/>';
          strXml += `<c:showVal val="${opts.showValue ? "1" : "0"}"/>`;
          strXml += `<c:showCatName val="0"/><c:showSerName val="${opts.showSerName ? "1" : "0"}"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`;
          strXml += `<c:showLeaderLines val="${opts.showLeaderLines ? "1" : "0"}"/>`;
          strXml += "</c:dLbls>";
        }
        if (chartType === "line" /* LINE */ || chartType === "radar" /* RADAR */) {
          strXml += "<c:marker>";
          strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
          if (opts.lineDataSymbolSize) strXml += `<c:size val="${opts.lineDataSymbolSize}"/>`;
          strXml += "  <c:spPr>";
          strXml += `    <a:solidFill>${createColorElement(opts.chartColors[obj._dataIndex + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : obj._dataIndex])}</a:solidFill>`;
          strXml += `    <a:ln w="${opts.lineDataSymbolLineSize}" cap="flat"><a:solidFill>${createColorElement(opts.lineDataSymbolLineColor || seriesColor)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`;
          strXml += "    <a:effectLst/>";
          strXml += "  </c:spPr>";
          strXml += "</c:marker>";
        }
        if ((chartType === "bar" /* BAR */ || chartType === "bar3D" /* BAR3D */) && data.length === 1 && (opts.chartColors && opts.chartColors !== BARCHART_COLORS && opts.chartColors.length > 1 || opts.invertedColors?.length)) {
          obj.values.forEach((value, index) => {
            const arrColors = value < 0 ? opts.invertedColors || opts.chartColors || BARCHART_COLORS : opts.chartColors || [];
            strXml += "  <c:dPt>";
            strXml += `    <c:idx val="${index}"/>`;
            strXml += '      <c:invertIfNegative val="0"/>';
            strXml += '    <c:bubble3D val="0"/>';
            strXml += "    <c:spPr>";
            if (opts.lineSize === 0) {
              strXml += "<a:ln><a:noFill/></a:ln>";
            } else if (chartType === "bar" /* BAR */) {
              strXml += "<a:solidFill>";
              strXml += '  <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
              strXml += "</a:solidFill>";
            } else {
              strXml += "<a:ln>";
              strXml += "  <a:solidFill>";
              strXml += '   <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
              strXml += "  </a:solidFill>";
              strXml += "</a:ln>";
            }
            strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
            strXml += "    </c:spPr>";
            strXml += "  </c:dPt>";
          });
        }
        {
          strXml += "<c:cat>";
          if (opts.catLabelFormatCode) {
            strXml += "  <c:numRef>";
            strXml += `    <c:f>Sheet1!$A$2:$A$${obj.labels[0].length + 1}</c:f>`;
            strXml += "    <c:numCache>";
            strXml += "      <c:formatCode>" + (opts.catLabelFormatCode || "General") + "</c:formatCode>";
            strXml += `      <c:ptCount val="${obj.labels[0].length}"/>`;
            obj.labels[0].forEach((label, idx) => strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`);
            strXml += "    </c:numCache>";
            strXml += "  </c:numRef>";
          } else {
            strXml += "  <c:multiLvlStrRef>";
            strXml += `    <c:f>Sheet1!$A$2:$${getExcelColName(obj.labels.length)}$${obj.labels[0].length + 1}</c:f>`;
            strXml += "    <c:multiLvlStrCache>";
            strXml += `      <c:ptCount val="${obj.labels[0].length}"/>`;
            obj.labels.forEach((labelsGroup) => {
              strXml += "<c:lvl>";
              labelsGroup.forEach((label, idx) => strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`);
              strXml += "</c:lvl>";
            });
            strXml += "    </c:multiLvlStrCache>";
            strXml += "  </c:multiLvlStrRef>";
          }
          strXml += "</c:cat>";
        }
        {
          strXml += "<c:val>";
          strXml += "  <c:numRef>";
          strXml += `<c:f>Sheet1!$${getExcelColName(obj._dataIndex + obj.labels.length + 1)}$2:$${getExcelColName(obj._dataIndex + obj.labels.length + 1)}$${obj.labels[0].length + 1}</c:f>`;
          strXml += "    <c:numCache>";
          strXml += "      <c:formatCode>" + (opts.valLabelFormatCode || opts.dataTableFormatCode || "General") + "</c:formatCode>";
          strXml += `      <c:ptCount val="${obj.labels[0].length}"/>`;
          obj.values.forEach((value, idx) => strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ""}</c:v></c:pt>`);
          strXml += "    </c:numCache>";
          strXml += "  </c:numRef>";
          strXml += "</c:val>";
        }
        if (chartType === "line" /* LINE */) strXml += '<c:smooth val="' + (opts.lineSmooth ? "1" : "0") + '"/>';
        strXml += "</c:ser>";
      });
      {
        strXml += "  <c:dLbls>";
        strXml += `    <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || "General"}" sourceLinked="0"/>`;
        strXml += "    <c:txPr>";
        strXml += "      <a:bodyPr/>";
        strXml += "      <a:lstStyle/>";
        strXml += "      <a:p><a:pPr>";
        strXml += `        <a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" u="none">`;
        strXml += "          <a:solidFill>" + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + "</a:solidFill>";
        strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || "Arial") + '"/>';
        strXml += "        </a:defRPr>";
        strXml += "      </a:pPr></a:p>";
        strXml += "    </c:txPr>";
        if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
        strXml += '    <c:showLegendKey val="0"/>';
        strXml += '    <c:showVal val="' + (opts.showValue ? "1" : "0") + '"/>';
        strXml += '    <c:showCatName val="0"/>';
        strXml += '    <c:showSerName val="' + (opts.showSerName ? "1" : "0") + '"/>';
        strXml += '    <c:showPercent val="0"/>';
        strXml += '    <c:showBubbleSize val="0"/>';
        strXml += `    <c:showLeaderLines val="${opts.showLeaderLines ? "1" : "0"}"/>`;
        strXml += "  </c:dLbls>";
      }
      if (chartType === "bar" /* BAR */) {
        strXml += `  <c:gapWidth val="${opts.barGapWidthPct}"/>`;
        strXml += `  <c:overlap val="${(opts.barGrouping || "").includes("tacked") ? 100 : opts.barOverlapPct ? opts.barOverlapPct : 0}"/>`;
      } else if (chartType === "bar3D" /* BAR3D */) {
        strXml += `  <c:gapWidth val="${opts.barGapWidthPct}"/>`;
        strXml += `  <c:gapDepth val="${opts.barGapDepthPct}"/>`;
        strXml += '  <c:shape val="' + opts.bar3DShape + '"/>';
      } else if (chartType === "line" /* LINE */) {
        strXml += '  <c:marker val="1"/>';
      }
      strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/><c:axId val="${AXIS_ID_SERIES_PRIMARY}"/>`;
      strXml += `</c:${chartType}Chart>`;
      break;
    case "scatter" /* SCATTER */:
      strXml += "<c:" + chartType + "Chart>";
      strXml += '<c:scatterStyle val="lineMarker"/>';
      strXml += '<c:varyColors val="0"/>';
      colorIndex = -1;
      data.filter((_obj, idx) => idx > 0).forEach((obj, idx) => {
        colorIndex++;
        strXml += "<c:ser>";
        strXml += `  <c:idx val="${idx}"/>`;
        strXml += `  <c:order val="${idx}"/>`;
        strXml += "  <c:tx>";
        strXml += "    <c:strRef>";
        strXml += `      <c:f>Sheet1!$${getExcelColName(idx + 2)}$1</c:f>`;
        strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + "</c:v></c:pt></c:strCache>";
        strXml += "    </c:strRef>";
        strXml += "  </c:tx>";
        strXml += "  <c:spPr>";
        {
          const tmpSerColor = opts.chartColors[colorIndex % opts.chartColors.length];
          if (tmpSerColor === "transparent") {
            strXml += "<a:noFill/>";
          } else if (opts.chartColorsOpacity) {
            strXml += "<a:solidFill>" + createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * 1e3).toString() + '"/>') + "</a:solidFill>";
          } else {
            strXml += "<a:solidFill>" + createColorElement(tmpSerColor) + "</a:solidFill>";
          }
          if (opts.lineSize === 0) {
            strXml += "<a:ln><a:noFill/></a:ln>";
          } else {
            strXml += `<a:ln w="${valToPts(opts.lineSize)}" cap="${createLineCap(opts.lineCap)}"><a:solidFill>${createColorElement(tmpSerColor)}</a:solidFill>`;
            strXml += `<a:prstDash val="${opts.lineDash || "solid"}"/><a:round/></a:ln>`;
          }
          strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
        }
        strXml += "  </c:spPr>";
        {
          strXml += "<c:marker>";
          strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
          if (opts.lineDataSymbolSize) {
            strXml += `<c:size val="${opts.lineDataSymbolSize}"/>`;
          }
          strXml += "<c:spPr>";
          strXml += `<a:solidFill>${createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx])}</a:solidFill>`;
          strXml += `<a:ln w="${opts.lineDataSymbolLineSize}" cap="flat"><a:solidFill>${createColorElement(opts.lineDataSymbolLineColor || opts.chartColors[colorIndex % opts.chartColors.length])}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`;
          strXml += "<a:effectLst/>";
          strXml += "</c:spPr>";
          strXml += "</c:marker>";
        }
        if (opts.showLabel) {
          const chartUuid = getUuid("-xxxx-xxxx-xxxx-xxxxxxxxxxxx");
          if (obj.labels[0] && (opts.dataLabelFormatScatter === "custom" || opts.dataLabelFormatScatter === "customXY")) {
            strXml += "<c:dLbls>";
            obj.labels[0].forEach((label, idx2) => {
              if (opts.dataLabelFormatScatter === "custom" || opts.dataLabelFormatScatter === "customXY") {
                strXml += "  <c:dLbl>";
                strXml += `    <c:idx val="${idx2}"/>`;
                strXml += "    <c:tx>";
                strXml += "      <c:rich>";
                strXml += "            <a:bodyPr>";
                strXml += "                <a:spAutoFit/>";
                strXml += "            </a:bodyPr>";
                strXml += "            <a:lstStyle/>";
                strXml += "            <a:p>";
                strXml += "                <a:pPr>";
                strXml += "                    <a:defRPr/>";
                strXml += "                </a:pPr>";
                strXml += "              <a:r>";
                strXml += '                    <a:rPr lang="' + (opts.lang || "en-US") + '" dirty="0"/>';
                strXml += "                    <a:t>" + encodeXmlEntities(label) + "</a:t>";
                strXml += "              </a:r>";
                if (opts.dataLabelFormatScatter === "customXY" && !/^ *$/.test(label)) {
                  strXml += "              <a:r>";
                  strXml += '                  <a:rPr lang="' + (opts.lang || "en-US") + '" baseline="0" dirty="0"/>';
                  strXml += "                  <a:t> (</a:t>";
                  strXml += "              </a:r>";
                  strXml += '              <a:fld id="{' + getUuid("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx") + '}" type="XVALUE">';
                  strXml += '                  <a:rPr lang="' + (opts.lang || "en-US") + '" baseline="0"/>';
                  strXml += "                  <a:pPr>";
                  strXml += "                      <a:defRPr/>";
                  strXml += "                  </a:pPr>";
                  strXml += "                  <a:t>[" + encodeXmlEntities(obj.name) + "</a:t>";
                  strXml += "              </a:fld>";
                  strXml += "              <a:r>";
                  strXml += '                  <a:rPr lang="' + (opts.lang || "en-US") + '" baseline="0" dirty="0"/>';
                  strXml += "                  <a:t>, </a:t>";
                  strXml += "              </a:r>";
                  strXml += '              <a:fld id="{' + getUuid("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx") + '}" type="YVALUE">';
                  strXml += '                  <a:rPr lang="' + (opts.lang || "en-US") + '" baseline="0"/>';
                  strXml += "                  <a:pPr>";
                  strXml += "                      <a:defRPr/>";
                  strXml += "                  </a:pPr>";
                  strXml += "                  <a:t>[" + encodeXmlEntities(obj.name) + "]</a:t>";
                  strXml += "              </a:fld>";
                  strXml += "              <a:r>";
                  strXml += '                  <a:rPr lang="' + (opts.lang || "en-US") + '" baseline="0" dirty="0"/>';
                  strXml += "                  <a:t>)</a:t>";
                  strXml += "              </a:r>";
                  strXml += '              <a:endParaRPr lang="' + (opts.lang || "en-US") + '" dirty="0"/>';
                }
                strXml += "            </a:p>";
                strXml += "      </c:rich>";
                strXml += "    </c:tx>";
                strXml += "    <c:spPr>";
                strXml += "        <a:noFill/>";
                strXml += "        <a:ln>";
                strXml += "            <a:noFill/>";
                strXml += "        </a:ln>";
                strXml += "        <a:effectLst/>";
                strXml += "    </c:spPr>";
                if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="0"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '       <c:showLeaderLines val="1"/>';
                strXml += "    <c:extLst>";
                strXml += '      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/>';
                strXml += '      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">';
                strXml += `            <c16:uniqueId val="{${"00000000".substring(0, 8 - (idx2 + 1).toString().length).toString()}${idx2 + 1}${chartUuid}}"/>`;
                strXml += "      </c:ext>";
                strXml += "        </c:extLst>";
                strXml += "</c:dLbl>";
              }
            });
            strXml += "</c:dLbls>";
          }
          if (opts.dataLabelFormatScatter === "XY") {
            strXml += "<c:dLbls>";
            strXml += "    <c:spPr>";
            strXml += "        <a:noFill/>";
            strXml += "        <a:ln>";
            strXml += "            <a:noFill/>";
            strXml += "        </a:ln>";
            strXml += "          <a:effectLst/>";
            strXml += "    </c:spPr>";
            strXml += "    <c:txPr>";
            strXml += "        <a:bodyPr>";
            strXml += "            <a:spAutoFit/>";
            strXml += "        </a:bodyPr>";
            strXml += "        <a:lstStyle/>";
            strXml += "        <a:p>";
            strXml += "            <a:pPr>";
            strXml += "                <a:defRPr/>";
            strXml += "            </a:pPr>";
            strXml += '            <a:endParaRPr lang="en-US"/>';
            strXml += "        </a:p>";
            strXml += "    </c:txPr>";
            if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
            strXml += '    <c:showLegendKey val="0"/>';
            strXml += ` <c:showVal val="${opts.showLabel ? "1" : "0"}"/>`;
            strXml += ` <c:showCatName val="${opts.showLabel ? "1" : "0"}"/>`;
            strXml += ` <c:showSerName val="${opts.showSerName ? "1" : "0"}"/>`;
            strXml += '    <c:showPercent val="0"/>';
            strXml += '    <c:showBubbleSize val="0"/>';
            strXml += "    <c:extLst>";
            strXml += '        <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
            strXml += '            <c15:showLeaderLines val="1"/>';
            strXml += "        </c:ext>";
            strXml += "    </c:extLst>";
            strXml += "</c:dLbls>";
          }
        }
        if (data.length === 1 && opts.chartColors !== BARCHART_COLORS) {
          obj.values.forEach((value, index) => {
            const arrColors = value < 0 ? opts.invertedColors || opts.chartColors || BARCHART_COLORS : opts.chartColors || [];
            strXml += "  <c:dPt>";
            strXml += `    <c:idx val="${index}"/>`;
            strXml += '      <c:invertIfNegative val="0"/>';
            strXml += '    <c:bubble3D val="0"/>';
            strXml += "    <c:spPr>";
            if (opts.lineSize === 0) {
              strXml += "<a:ln><a:noFill/></a:ln>";
            } else {
              strXml += "<a:solidFill>";
              strXml += ' <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
              strXml += "</a:solidFill>";
            }
            strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
            strXml += "    </c:spPr>";
            strXml += "  </c:dPt>";
          });
        }
        {
          strXml += "<c:xVal>";
          strXml += "  <c:numRef>";
          strXml += `    <c:f>Sheet1!$A$2:$A$${data[0].values.length + 1}</c:f>`;
          strXml += "    <c:numCache>";
          strXml += "      <c:formatCode>General</c:formatCode>";
          strXml += `      <c:ptCount val="${data[0].values.length}"/>`;
          data[0].values.forEach((value, idx2) => {
            strXml += `<c:pt idx="${idx2}"><c:v>${value || value === 0 ? value : ""}</c:v></c:pt>`;
          });
          strXml += "    </c:numCache>";
          strXml += "  </c:numRef>";
          strXml += "</c:xVal>";
          strXml += "<c:yVal>";
          strXml += "  <c:numRef>";
          strXml += `    <c:f>Sheet1!$${getExcelColName(idx + 2)}$2:$${getExcelColName(idx + 2)}$${data[0].values.length + 1}</c:f>`;
          strXml += "    <c:numCache>";
          strXml += "      <c:formatCode>General</c:formatCode>";
          strXml += `      <c:ptCount val="${data[0].values.length}"/>`;
          data[0].values.forEach((_value, idx2) => {
            strXml += `<c:pt idx="${idx2}"><c:v>${obj.values[idx2] || obj.values[idx2] === 0 ? obj.values[idx2] : ""}</c:v></c:pt>`;
          });
          strXml += "    </c:numCache>";
          strXml += "  </c:numRef>";
          strXml += "</c:yVal>";
        }
        strXml += '<c:smooth val="' + (opts.lineSmooth ? "1" : "0") + '"/>';
        strXml += "</c:ser>";
      });
      {
        strXml += "  <c:dLbls>";
        strXml += `    <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || "General"}" sourceLinked="0"/>`;
        strXml += "    <c:txPr>";
        strXml += "      <a:bodyPr/>";
        strXml += "      <a:lstStyle/>";
        strXml += "      <a:p><a:pPr>";
        strXml += `        <a:defRPr b="${opts.dataLabelFontBold ? "1" : "0"}" i="${opts.dataLabelFontItalic ? "1" : "0"}" strike="noStrike" sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" u="none">`;
        strXml += "          <a:solidFill>" + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + "</a:solidFill>";
        strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || "Arial") + '"/>';
        strXml += "        </a:defRPr>";
        strXml += "      </a:pPr></a:p>";
        strXml += "    </c:txPr>";
        if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>';
        strXml += '    <c:showLegendKey val="0"/>';
        strXml += '    <c:showVal val="' + (opts.showValue ? "1" : "0") + '"/>';
        strXml += '    <c:showCatName val="0"/>';
        strXml += '    <c:showSerName val="' + (opts.showSerName ? "1" : "0") + '"/>';
        strXml += '    <c:showPercent val="0"/>';
        strXml += '    <c:showBubbleSize val="0"/>';
        strXml += "  </c:dLbls>";
      }
      strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`;
      strXml += "</c:" + chartType + "Chart>";
      break;
    case "bubble" /* BUBBLE */:
    case "bubble3D" /* BUBBLE3D */:
      strXml += "<c:bubbleChart>";
      strXml += '<c:varyColors val="0"/>';
      colorIndex = -1;
      data.filter((_obj, idx) => idx > 0).forEach((obj, idx) => {
        colorIndex++;
        strXml += "<c:ser>";
        strXml += `  <c:idx val="${idx}"/>`;
        strXml += `  <c:order val="${idx}"/>`;
        strXml += "  <c:tx>";
        strXml += "    <c:strRef>";
        strXml += "      <c:f>Sheet1!$" + getExcelColName(idxColLtr + 1) + "$1</c:f>";
        strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + "</c:v></c:pt></c:strCache>";
        strXml += "    </c:strRef>";
        strXml += "  </c:tx>";
        {
          strXml += "<c:spPr>";
          const tmpSerColor = opts.chartColors[colorIndex % opts.chartColors.length];
          if (tmpSerColor === "transparent") {
            strXml += "<a:noFill/>";
          } else if (opts.chartColorsOpacity) {
            strXml += `<a:solidFill>${createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * 1e3).toString() + '"/>')}</a:solidFill>`;
          } else {
            strXml += "<a:solidFill>" + createColorElement(tmpSerColor) + "</a:solidFill>";
          }
          if (opts.lineSize === 0) {
            strXml += "<a:ln><a:noFill/></a:ln>";
          } else if (opts.dataBorder) {
            strXml += `<a:ln w="${valToPts(opts.dataBorder.pt)}" cap="flat"><a:solidFill>${createColorElement(opts.dataBorder.color)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`;
          } else {
            strXml += `<a:ln w="${valToPts(opts.lineSize)}" cap="flat"><a:solidFill>${createColorElement(tmpSerColor)}</a:solidFill>`;
            strXml += `<a:prstDash val="${opts.lineDash || "solid"}"/><a:round/></a:ln>`;
          }
          strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
          strXml += "</c:spPr>";
        }
        {
          strXml += "<c:xVal>";
          strXml += "  <c:numRef>";
          strXml += `    <c:f>Sheet1!$A$2:$A$${data[0].values.length + 1}</c:f>`;
          strXml += "    <c:numCache>";
          strXml += "      <c:formatCode>General</c:formatCode>";
          strXml += `      <c:ptCount val="${data[0].values.length}"/>`;
          data[0].values.forEach((value, idx2) => {
            strXml += `<c:pt idx="${idx2}"><c:v>${value || value === 0 ? value : ""}</c:v></c:pt>`;
          });
          strXml += "    </c:numCache>";
          strXml += "  </c:numRef>";
          strXml += "</c:xVal>";
          strXml += "<c:yVal>";
          strXml += "  <c:numRef>";
          strXml += `<c:f>Sheet1!$${getExcelColName(idxColLtr + 1)}$2:$${getExcelColName(idxColLtr + 1)}$${data[0].values.length + 1}</c:f>`;
          idxColLtr++;
          strXml += "    <c:numCache>";
          strXml += "      <c:formatCode>General</c:formatCode>";
          strXml += `      <c:ptCount val="${data[0].values.length}"/>`;
          data[0].values.forEach((_value, idx2) => {
            strXml += `<c:pt idx="${idx2}"><c:v>${obj.values[idx2] || obj.values[idx2] === 0 ? obj.values[idx2] : ""}</c:v></c:pt>`;
          });
          strXml += "    </c:numCache>";
          strXml += "  </c:numRef>";
          strXml += "</c:yVal>";
        }
        strXml += "  <c:bubbleSize>";
        strXml += "    <c:numRef>";
        strXml += `<c:f>Sheet1!$${getExcelColName(idxColLtr + 1)}$2:$${getExcelColName(idxColLtr + 1)}$${obj.sizes.length + 1}</c:f>`;
        idxColLtr++;
        strXml += "      <c:numCache>";
        strXml += "        <c:formatCode>General</c:formatCode>";
        strXml += `           <c:ptCount val="${obj.sizes.length}"/>`;
        obj.sizes.forEach((value, idx2) => {
          strXml += `<c:pt idx="${idx2}"><c:v>${value || ""}</c:v></c:pt>`;
        });
        strXml += "      </c:numCache>";
        strXml += "    </c:numRef>";
        strXml += "  </c:bubbleSize>";
        strXml += '  <c:bubble3D val="' + (chartType === "bubble3D" /* BUBBLE3D */ ? "1" : "0") + '"/>';
        strXml += "</c:ser>";
      });
      {
        strXml += "<c:dLbls>";
        strXml += `<c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || "General"}" sourceLinked="0"/>`;
        strXml += "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>";
        strXml += `<a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${Math.round(
          Math.round(opts.dataLabelFontSize || DEF_FONT_SIZE) * 100
        )}" u="none">`;
        strXml += `<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>`;
        strXml += `<a:latin typeface="${opts.dataLabelFontFace || "Arial"}"/>`;
        strXml += "</a:defRPr></a:pPr></a:p></c:txPr>";
        if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`;
        strXml += '<c:showLegendKey val="0"/>';
        strXml += `<c:showVal val="${opts.showValue ? "1" : "0"}"/>`;
        strXml += `<c:showCatName val="0"/><c:showSerName val="${opts.showSerName ? "1" : "0"}"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`;
        strXml += "<c:extLst>";
        strXml += '  <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
        strXml += '    <c15:showLeaderLines val="' + (opts.showLeaderLines ? "1" : "0") + '"/>';
        strXml += "  </c:ext>";
        strXml += "</c:extLst>";
        strXml += "</c:dLbls>";
      }
      strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`;
      strXml += "</c:bubbleChart>";
      break;
    case "doughnut" /* DOUGHNUT */:
    case "pie" /* PIE */:
      optsChartData = data[0];
      strXml += "<c:" + chartType + "Chart>";
      strXml += '  <c:varyColors val="1"/>';
      strXml += "<c:ser>";
      strXml += '  <c:idx val="0"/>';
      strXml += '  <c:order val="0"/>';
      strXml += "  <c:tx>";
      strXml += "    <c:strRef>";
      strXml += "      <c:f>Sheet1!$B$1</c:f>";
      strXml += "      <c:strCache>";
      strXml += '        <c:ptCount val="1"/>';
      strXml += '        <c:pt idx="0"><c:v>' + encodeXmlEntities(optsChartData.name) + "</c:v></c:pt>";
      strXml += "      </c:strCache>";
      strXml += "    </c:strRef>";
      strXml += "  </c:tx>";
      strXml += "  <c:spPr>";
      strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
      strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
      if (opts.dataNoEffects) {
        strXml += "<a:effectLst/>";
      } else {
        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
      }
      strXml += "  </c:spPr>";
      optsChartData.labels[0].forEach((_label, idx) => {
        strXml += "<c:dPt>";
        strXml += ` <c:idx val="${idx}"/>`;
        strXml += ' <c:bubble3D val="0"/>';
        strXml += " <c:spPr>";
        strXml += `<a:solidFill>${createColorElement(
          opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]
        )}</a:solidFill>`;
        if (opts.dataBorder) {
          strXml += `<a:ln w="${valToPts(opts.dataBorder.pt)}" cap="flat"><a:solidFill>${createColorElement(
            opts.dataBorder.color
          )}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`;
        }
        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
        strXml += "  </c:spPr>";
        strXml += "</c:dPt>";
      });
      strXml += "<c:dLbls>";
      optsChartData.labels[0].forEach((_label, idx) => {
        strXml += "<c:dLbl>";
        strXml += ` <c:idx val="${idx}"/>`;
        strXml += `  <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || "General"}" sourceLinked="0"/>`;
        strXml += "  <c:spPr/><c:txPr>";
        strXml += "   <a:bodyPr/><a:lstStyle/>";
        strXml += "   <a:p><a:pPr>";
        strXml += `   <a:defRPr sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" u="none" strike="noStrike">`;
        strXml += "    <a:solidFill>" + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + "</a:solidFill>";
        strXml += `    <a:latin typeface="${opts.dataLabelFontFace || "Arial"}"/>`;
        strXml += "   </a:defRPr>";
        strXml += "      </a:pPr></a:p>";
        strXml += "    </c:txPr>";
        if (chartType === "pie" /* PIE */ && opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`;
        strXml += '    <c:showLegendKey val="0"/>';
        strXml += '    <c:showVal val="' + (opts.showValue ? "1" : "0") + '"/>';
        strXml += '    <c:showCatName val="' + (opts.showLabel ? "1" : "0") + '"/>';
        strXml += '    <c:showSerName val="' + (opts.showSerName ? "1" : "0") + '"/>';
        strXml += '    <c:showPercent val="' + (opts.showPercent ? "1" : "0") + '"/>';
        strXml += '    <c:showBubbleSize val="0"/>';
        strXml += "  </c:dLbl>";
      });
      strXml += ` <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || "General"}" sourceLinked="0"/>`;
      strXml += "    <c:txPr>";
      strXml += "      <a:bodyPr/>";
      strXml += "      <a:lstStyle/>";
      strXml += "      <a:p>";
      strXml += "        <a:pPr>";
      strXml += `          <a:defRPr sz="1800" b="${opts.dataLabelFontBold ? "1" : "0"}" i="${opts.dataLabelFontItalic ? "1" : "0"}" u="none" strike="noStrike">`;
      strXml += '            <a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>';
      strXml += "          </a:defRPr>";
      strXml += "        </a:pPr>";
      strXml += "      </a:p>";
      strXml += "    </c:txPr>";
      strXml += chartType === "pie" /* PIE */ ? '<c:dLblPos val="ctr"/>' : "";
      strXml += '    <c:showLegendKey val="0"/>';
      strXml += '    <c:showVal val="0"/>';
      strXml += '    <c:showCatName val="1"/>';
      strXml += '    <c:showSerName val="0"/>';
      strXml += '    <c:showPercent val="1"/>';
      strXml += '    <c:showBubbleSize val="0"/>';
      strXml += ` <c:showLeaderLines val="${opts.showLeaderLines ? "1" : "0"}"/>`;
      strXml += "</c:dLbls>";
      strXml += "<c:cat>";
      strXml += "  <c:strRef>";
      strXml += `    <c:f>Sheet1!$A$2:$A$${optsChartData.labels[0].length + 1}</c:f>`;
      strXml += "    <c:strCache>";
      strXml += `         <c:ptCount val="${optsChartData.labels[0].length}"/>`;
      optsChartData.labels[0].forEach((label, idx) => {
        strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`;
      });
      strXml += "    </c:strCache>";
      strXml += "  </c:strRef>";
      strXml += "</c:cat>";
      strXml += "  <c:val>";
      strXml += "    <c:numRef>";
      strXml += `      <c:f>Sheet1!$B$2:$B$${optsChartData.labels[0].length + 1}</c:f>`;
      strXml += "      <c:numCache>";
      strXml += `           <c:ptCount val="${optsChartData.labels[0].length}"/>`;
      optsChartData.values.forEach((value, idx) => {
        strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ""}</c:v></c:pt>`;
      });
      strXml += "      </c:numCache>";
      strXml += "    </c:numRef>";
      strXml += "  </c:val>";
      strXml += "  </c:ser>";
      strXml += `  <c:firstSliceAng val="${opts.firstSliceAng ? Math.round(opts.firstSliceAng) : 0}"/>`;
      if (chartType === "doughnut" /* DOUGHNUT */) strXml += `<c:holeSize val="${typeof opts.holeSize === "number" ? opts.holeSize : "50"}"/>`;
      strXml += "</c:" + chartType + "Chart>";
      break;
    default:
      strXml += "";
      break;
  }
  return strXml;
}
function makeCatAxis(opts, axisId, valAxisId) {
  let strXml = "";
  if (opts._type === "scatter" /* SCATTER */ || opts._type === "bubble" /* BUBBLE */ || opts._type === "bubble3D" /* BUBBLE3D */) {
    strXml += "<c:valAx>";
  } else {
    strXml += "<c:" + (opts.catLabelFormatCode ? "dateAx" : "catAx") + ">";
  }
  strXml += '  <c:axId val="' + axisId + '"/>';
  strXml += "  <c:scaling>";
  strXml += '<c:orientation val="' + (opts.catAxisOrientation || (opts.barDir === "col" ? "minMax" : "minMax")) + '"/>';
  if (opts.catAxisMaxVal || opts.catAxisMaxVal === 0) strXml += `<c:max val="${opts.catAxisMaxVal}"/>`;
  if (opts.catAxisMinVal || opts.catAxisMinVal === 0) strXml += `<c:min val="${opts.catAxisMinVal}"/>`;
  strXml += "</c:scaling>";
  strXml += '  <c:delete val="' + (opts.catAxisHidden ? "1" : "0") + '"/>';
  strXml += '  <c:axPos val="' + (opts.barDir === "col" ? "b" : "l") + '"/>';
  strXml += opts.catGridLine.style !== "none" ? createGridLineElement(opts.catGridLine) : "";
  if (opts.showCatAxisTitle) {
    strXml += genXmlTitle({
      color: opts.catAxisTitleColor,
      fontFace: opts.catAxisTitleFontFace,
      fontSize: opts.catAxisTitleFontSize,
      titleRotate: opts.catAxisTitleRotate,
      title: opts.catAxisTitle || "Axis Title"
    });
  }
  if (opts._type === "scatter" /* SCATTER */ || opts._type === "bubble" /* BUBBLE */ || opts._type === "bubble3D" /* BUBBLE3D */) {
    strXml += '  <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? encodeXmlEntities(opts.valAxisLabelFormatCode) : "General") + '" sourceLinked="1"/>';
  } else {
    strXml += '  <c:numFmt formatCode="' + (encodeXmlEntities(opts.catLabelFormatCode) || "General") + '" sourceLinked="1"/>';
  }
  if (opts._type === "scatter" /* SCATTER */) {
    strXml += '  <c:majorTickMark val="none"/>';
    strXml += '  <c:minorTickMark val="none"/>';
    strXml += '  <c:tickLblPos val="nextTo"/>';
  } else {
    strXml += '  <c:majorTickMark val="' + (opts.catAxisMajorTickMark || "out") + '"/>';
    strXml += '  <c:minorTickMark val="' + (opts.catAxisMinorTickMark || "none") + '"/>';
    strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || (opts.barDir === "col" ? "low" : "nextTo")) + '"/>';
  }
  strXml += "  <c:spPr>";
  strXml += `    <a:ln w="${opts.catAxisLineSize ? valToPts(opts.catAxisLineSize) : ONEPT}" cap="flat">`;
  strXml += !opts.catAxisLineShow ? "<a:noFill/>" : "<a:solidFill>" + createColorElement(opts.catAxisLineColor || DEF_CHART_GRIDLINE.color) + "</a:solidFill>";
  strXml += '      <a:prstDash val="' + (opts.catAxisLineStyle || "solid") + '"/>';
  strXml += "      <a:round/>";
  strXml += "    </a:ln>";
  strXml += "  </c:spPr>";
  strXml += "  <c:txPr>";
  if (opts.catAxisLabelRotate) {
    strXml += `<a:bodyPr rot="${convertRotationDegrees(opts.catAxisLabelRotate)}"/>`;
  } else {
    strXml += "<a:bodyPr/>";
  }
  strXml += "    <a:lstStyle/>";
  strXml += "    <a:p>";
  strXml += "    <a:pPr>";
  strXml += `      <a:defRPr sz="${Math.round((opts.catAxisLabelFontSize || DEF_FONT_SIZE) * 100)}" b="${opts.catAxisLabelFontBold ? 1 : 0}" i="${opts.catAxisLabelFontItalic ? 1 : 0}" u="none" strike="noStrike">`;
  strXml += "      <a:solidFill>" + createColorElement(opts.catAxisLabelColor || DEF_FONT_COLOR) + "</a:solidFill>";
  strXml += '      <a:latin typeface="' + (opts.catAxisLabelFontFace || "Arial") + '"/>';
  strXml += "   </a:defRPr>";
  strXml += "  </a:pPr>";
  strXml += '  <a:endParaRPr lang="' + (opts.lang || "en-US") + '"/>';
  strXml += "  </a:p>";
  strXml += " </c:txPr>";
  strXml += ' <c:crossAx val="' + valAxisId + '"/>';
  strXml += ` <c:${typeof opts.valAxisCrossesAt === "number" ? "crossesAt" : "crosses"} val="${opts.valAxisCrossesAt || "autoZero"}"/>`;
  strXml += ' <c:auto val="1"/>';
  strXml += ' <c:lblAlgn val="ctr"/>';
  strXml += ` <c:noMultiLvlLbl val="${opts.catAxisMultiLevelLabels ? 0 : 1}"/>`;
  if (opts.catAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.catAxisLabelFrequency + '"/>';
  if (opts.catLabelFormatCode || opts._type === "scatter" /* SCATTER */ || opts._type === "bubble" /* BUBBLE */ || opts._type === "bubble3D" /* BUBBLE3D */) {
    if (opts.catLabelFormatCode) {
      ["catAxisBaseTimeUnit", "catAxisMajorTimeUnit", "catAxisMinorTimeUnit"].forEach((opt) => {
        if (opts[opt] && (typeof opts[opt] !== "string" || !["days", "months", "years"].includes(opts[opt].toLowerCase()))) {
          console.warn(`"${opt}" must be one of: 'days','months','years' !`);
          opts[opt] = null;
        }
      });
      if (opts.catAxisBaseTimeUnit) strXml += '<c:baseTimeUnit val="' + opts.catAxisBaseTimeUnit.toLowerCase() + '"/>';
      if (opts.catAxisMajorTimeUnit) strXml += '<c:majorTimeUnit val="' + opts.catAxisMajorTimeUnit.toLowerCase() + '"/>';
      if (opts.catAxisMinorTimeUnit) strXml += '<c:minorTimeUnit val="' + opts.catAxisMinorTimeUnit.toLowerCase() + '"/>';
    }
    if (opts.catAxisMajorUnit) strXml += `<c:majorUnit val="${opts.catAxisMajorUnit}"/>`;
    if (opts.catAxisMinorUnit) strXml += `<c:minorUnit val="${opts.catAxisMinorUnit}"/>`;
  }
  if (opts._type === "scatter" /* SCATTER */ || opts._type === "bubble" /* BUBBLE */ || opts._type === "bubble3D" /* BUBBLE3D */) {
    strXml += "</c:valAx>";
  } else {
    strXml += "</c:" + (opts.catLabelFormatCode ? "dateAx" : "catAx") + ">";
  }
  return strXml;
}
function makeValAxis(opts, valAxisId) {
  let axisPos = valAxisId === AXIS_ID_VALUE_PRIMARY ? opts.barDir === "col" ? "l" : "b" : opts.barDir !== "col" ? "r" : "t";
  if (valAxisId === AXIS_ID_VALUE_SECONDARY) axisPos = "r";
  const crossAxId = valAxisId === AXIS_ID_VALUE_PRIMARY ? AXIS_ID_CATEGORY_PRIMARY : AXIS_ID_CATEGORY_SECONDARY;
  let strXml = "";
  strXml += "<c:valAx>";
  strXml += '  <c:axId val="' + valAxisId + '"/>';
  strXml += "  <c:scaling>";
  if (opts.valAxisLogScaleBase) strXml += `<c:logBase val="${opts.valAxisLogScaleBase}"/>`;
  strXml += '<c:orientation val="' + (opts.valAxisOrientation || (opts.barDir === "col" ? "minMax" : "minMax")) + '"/>';
  if (opts.valAxisMaxVal || opts.valAxisMaxVal === 0) strXml += `<c:max val="${opts.valAxisMaxVal}"/>`;
  if (opts.valAxisMinVal || opts.valAxisMinVal === 0) strXml += `<c:min val="${opts.valAxisMinVal}"/>`;
  strXml += "  </c:scaling>";
  strXml += `  <c:delete val="${opts.valAxisHidden ? 1 : 0}"/>`;
  strXml += '  <c:axPos val="' + axisPos + '"/>';
  if (opts.valGridLine.style !== "none") strXml += createGridLineElement(opts.valGridLine);
  if (opts.showValAxisTitle) {
    strXml += genXmlTitle({
      color: opts.valAxisTitleColor,
      fontFace: opts.valAxisTitleFontFace,
      fontSize: opts.valAxisTitleFontSize,
      titleRotate: opts.valAxisTitleRotate,
      title: opts.valAxisTitle || "Axis Title"
    });
  }
  strXml += `<c:numFmt formatCode="${opts.valAxisLabelFormatCode ? encodeXmlEntities(opts.valAxisLabelFormatCode) : "General"}" sourceLinked="0"/>`;
  if (opts._type === "scatter" /* SCATTER */) {
    strXml += '  <c:majorTickMark val="none"/>';
    strXml += '  <c:minorTickMark val="none"/>';
    strXml += '  <c:tickLblPos val="nextTo"/>';
  } else {
    strXml += ' <c:majorTickMark val="' + (opts.valAxisMajorTickMark || "out") + '"/>';
    strXml += ' <c:minorTickMark val="' + (opts.valAxisMinorTickMark || "none") + '"/>';
    strXml += ' <c:tickLblPos val="' + (opts.valAxisLabelPos || (opts.barDir === "col" ? "nextTo" : "low")) + '"/>';
  }
  strXml += " <c:spPr>";
  strXml += `   <a:ln w="${opts.valAxisLineSize ? valToPts(opts.valAxisLineSize) : ONEPT}" cap="flat">`;
  strXml += !opts.valAxisLineShow ? "<a:noFill/>" : "<a:solidFill>" + createColorElement(opts.valAxisLineColor || DEF_CHART_GRIDLINE.color) + "</a:solidFill>";
  strXml += '     <a:prstDash val="' + (opts.valAxisLineStyle || "solid") + '"/>';
  strXml += "     <a:round/>";
  strXml += "   </a:ln>";
  strXml += " </c:spPr>";
  strXml += " <c:txPr>";
  strXml += `  <a:bodyPr${opts.valAxisLabelRotate ? ' rot="' + convertRotationDegrees(opts.valAxisLabelRotate).toString() + '"' : ""}/>`;
  strXml += "  <a:lstStyle/>";
  strXml += "  <a:p>";
  strXml += "    <a:pPr>";
  strXml += `      <a:defRPr sz="${Math.round((opts.valAxisLabelFontSize || DEF_FONT_SIZE) * 100)}" b="${opts.valAxisLabelFontBold ? 1 : 0}" i="${opts.valAxisLabelFontItalic ? 1 : 0}" u="none" strike="noStrike">`;
  strXml += "        <a:solidFill>" + createColorElement(opts.valAxisLabelColor || DEF_FONT_COLOR) + "</a:solidFill>";
  strXml += '        <a:latin typeface="' + (opts.valAxisLabelFontFace || "Arial") + '"/>';
  strXml += "      </a:defRPr>";
  strXml += "    </a:pPr>";
  strXml += '  <a:endParaRPr lang="' + (opts.lang || "en-US") + '"/>';
  strXml += "  </a:p>";
  strXml += " </c:txPr>";
  strXml += ' <c:crossAx val="' + crossAxId + '"/>';
  if (typeof opts.catAxisCrossesAt === "number") {
    strXml += ` <c:crossesAt val="${opts.catAxisCrossesAt}"/>`;
  } else if (typeof opts.catAxisCrossesAt === "string") {
    strXml += ' <c:crosses val="' + opts.catAxisCrossesAt + '"/>';
  } else {
    const isRight = axisPos === "r" || axisPos === "t";
    const crosses = isRight ? "max" : "autoZero";
    strXml += ' <c:crosses val="' + crosses + '"/>';
  }
  strXml += ' <c:crossBetween val="' + (opts._type === "scatter" /* SCATTER */ || !!(Array.isArray(opts._type) && opts._type.filter((type) => type.type === "area" /* AREA */).length > 0) ? "midCat" : "between") + '"/>';
  if (opts.valAxisMajorUnit) strXml += ` <c:majorUnit val="${opts.valAxisMajorUnit}"/>`;
  if (opts.valAxisDisplayUnit) {
    strXml += `<c:dispUnits><c:builtInUnit val="${opts.valAxisDisplayUnit}"/>${opts.valAxisDisplayUnitLabel ? "<c:dispUnitsLbl/>" : ""}</c:dispUnits>`;
  }
  strXml += "</c:valAx>";
  return strXml;
}
function makeSerAxis(opts, axisId, valAxisId) {
  let strXml = "";
  strXml += "<c:serAx>";
  strXml += '  <c:axId val="' + axisId + '"/>';
  strXml += '  <c:scaling><c:orientation val="' + (opts.serAxisOrientation || (opts.barDir === "col" ? "minMax" : "minMax")) + '"/></c:scaling>';
  strXml += '  <c:delete val="' + (opts.serAxisHidden ? "1" : "0") + '"/>';
  strXml += '  <c:axPos val="' + (opts.barDir === "col" ? "b" : "l") + '"/>';
  strXml += opts.serGridLine.style !== "none" ? createGridLineElement(opts.serGridLine) : "";
  if (opts.showSerAxisTitle) {
    strXml += genXmlTitle({
      color: opts.serAxisTitleColor,
      fontFace: opts.serAxisTitleFontFace,
      fontSize: opts.serAxisTitleFontSize,
      titleRotate: opts.serAxisTitleRotate,
      title: opts.serAxisTitle || "Axis Title"
    });
  }
  strXml += `  <c:numFmt formatCode="${encodeXmlEntities(opts.serLabelFormatCode) || "General"}" sourceLinked="0"/>`;
  strXml += '  <c:majorTickMark val="out"/>';
  strXml += '  <c:minorTickMark val="none"/>';
  strXml += `  <c:tickLblPos val="${opts.serAxisLabelPos || opts.barDir === "col" ? "low" : "nextTo"}"/>`;
  strXml += "  <c:spPr>";
  strXml += '    <a:ln w="12700" cap="flat">';
  strXml += !opts.serAxisLineShow ? "<a:noFill/>" : `<a:solidFill>${createColorElement(opts.serAxisLineColor || DEF_CHART_GRIDLINE.color)}</a:solidFill>`;
  strXml += '      <a:prstDash val="solid"/>';
  strXml += "      <a:round/>";
  strXml += "    </a:ln>";
  strXml += "  </c:spPr>";
  strXml += "  <c:txPr>";
  strXml += "    <a:bodyPr/>";
  strXml += "    <a:lstStyle/>";
  strXml += "    <a:p>";
  strXml += "    <a:pPr>";
  strXml += `    <a:defRPr sz="${Math.round((opts.serAxisLabelFontSize || DEF_FONT_SIZE) * 100)}" b="${opts.serAxisLabelFontBold ? "1" : "0"}" i="${opts.serAxisLabelFontItalic ? "1" : "0"}" u="none" strike="noStrike">`;
  strXml += `      <a:solidFill>${createColorElement(opts.serAxisLabelColor || DEF_FONT_COLOR)}</a:solidFill>`;
  strXml += `      <a:latin typeface="${opts.serAxisLabelFontFace || "Arial"}"/>`;
  strXml += "   </a:defRPr>";
  strXml += "  </a:pPr>";
  strXml += '  <a:endParaRPr lang="' + (opts.lang || "en-US") + '"/>';
  strXml += "  </a:p>";
  strXml += " </c:txPr>";
  strXml += ' <c:crossAx val="' + valAxisId + '"/>';
  strXml += ' <c:crosses val="autoZero"/>';
  if (opts.serAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.serAxisLabelFrequency + '"/>';
  if (opts.serLabelFormatCode) {
    ["serAxisBaseTimeUnit", "serAxisMajorTimeUnit", "serAxisMinorTimeUnit"].forEach((opt) => {
      if (opts[opt] && (typeof opts[opt] !== "string" || !["days", "months", "years"].includes(opt.toLowerCase()))) {
        console.warn(`"${opt}" must be one of: 'days','months','years' !`);
        opts[opt] = null;
      }
    });
    if (opts.serAxisBaseTimeUnit) strXml += ` <c:baseTimeUnit  val="${opts.serAxisBaseTimeUnit.toLowerCase()}"/>`;
    if (opts.serAxisMajorTimeUnit) strXml += ` <c:majorTimeUnit val="${opts.serAxisMajorTimeUnit.toLowerCase()}"/>`;
    if (opts.serAxisMinorTimeUnit) strXml += ` <c:minorTimeUnit val="${opts.serAxisMinorTimeUnit.toLowerCase()}"/>`;
    if (opts.serAxisMajorUnit) strXml += ` <c:majorUnit val="${opts.serAxisMajorUnit}"/>`;
    if (opts.serAxisMinorUnit) strXml += ` <c:minorUnit val="${opts.serAxisMinorUnit}"/>`;
  }
  strXml += "</c:serAx>";
  return strXml;
}
function genXmlTitle(opts, chartX, chartY) {
  const align = opts.titleAlign === "left" || opts.titleAlign === "right" ? `<a:pPr algn="${opts.titleAlign.substring(0, 1)}">` : "<a:pPr>";
  const rotate = opts.titleRotate ? `<a:bodyPr rot="${convertRotationDegrees(opts.titleRotate)}"/>` : "<a:bodyPr/>";
  const sizeAttr = opts.fontSize ? `sz="${Math.round(opts.fontSize * 100)}"` : "";
  const titleBold = opts.titleBold ? 1 : 0;
  let layout = "<c:layout/>";
  if (opts.titlePos && typeof opts.titlePos.x === "number" && typeof opts.titlePos.y === "number") {
    const totalX = opts.titlePos.x + chartX;
    const totalY = opts.titlePos.y + chartY;
    let valX = totalX === 0 ? 0 : totalX * (totalX / 5) / 10;
    if (valX >= 1) valX = valX / 10;
    if (valX >= 0.1) valX = valX / 10;
    let valY = totalY === 0 ? 0 : totalY * (totalY / 5) / 10;
    if (valY >= 1) valY = valY / 10;
    if (valY >= 0.1) valY = valY / 10;
    layout = `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="${valX}"/><c:y val="${valY}"/></c:manualLayout></c:layout>`;
  }
  return `<c:title>
      <c:tx>
        <c:rich>
          ${rotate}
          <a:lstStyle/>
          <a:p>
            ${align}
            <a:defRPr ${sizeAttr} b="${titleBold}" i="0" u="none" strike="noStrike">
              <a:solidFill>${createColorElement(opts.color || DEF_FONT_COLOR)}</a:solidFill>
              <a:latin typeface="${opts.fontFace || "Arial"}"/>
            </a:defRPr>
          </a:pPr>
          <a:r>
            <a:rPr ${sizeAttr} b="${titleBold}" i="0" u="none" strike="noStrike">
              <a:solidFill>${createColorElement(opts.color || DEF_FONT_COLOR)}</a:solidFill>
              <a:latin typeface="${opts.fontFace || "Arial"}"/>
            </a:rPr>
            <a:t>${encodeXmlEntities(opts.title) || ""}</a:t>
          </a:r>
        </a:p>
        </c:rich>
      </c:tx>
      ${layout}
      <c:overlay val="0"/>
    </c:title>`;
}
function getExcelColName(colIndex) {
  let colStr = "";
  const colIdx = colIndex - 1;
  if (colIdx <= 25) {
    colStr = LETTERS[colIdx];
  } else {
    colStr = `${LETTERS[Math.floor(colIdx / LETTERS.length - 1)]}${LETTERS[colIdx % LETTERS.length]}`;
  }
  return colStr;
}
function createShadowElement(options, defaults) {
  if (!options) {
    return "<a:effectLst/>";
  } else if (typeof options !== "object") {
    console.warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`");
    return "<a:effectLst/>";
  }
  let strXml = "<a:effectLst>";
  const opts = { ...defaults, ...options };
  const type = opts.type || "outer";
  const blur = valToPts(opts.blur);
  const offset = valToPts(opts.offset);
  const angle = Math.round(opts.angle * 6e4);
  const color = opts.color;
  const opacity = Math.round(opts.opacity * 1e5);
  const rotShape = opts.rotateWithShape ? 1 : 0;
  strXml += `<a:${type}Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="${blur}" rotWithShape="${rotShape}" dist="${offset}" dir="${angle}">`;
  strXml += `<a:srgbClr val="${color}">`;
  strXml += `<a:alpha val="${opacity}"/></a:srgbClr>`;
  strXml += `</a:${type}Shdw>`;
  strXml += "</a:effectLst>";
  return strXml;
}
function createGridLineElement(glOpts) {
  let strXml = "<c:majorGridlines>";
  strXml += " <c:spPr>";
  strXml += `  <a:ln w="${valToPts(glOpts.size || DEF_CHART_GRIDLINE.size)}" cap="${createLineCap(glOpts.cap || DEF_CHART_GRIDLINE.cap)}">`;
  strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || DEF_CHART_GRIDLINE.color) + '"/></a:solidFill>';
  strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>';
  strXml += "  </a:ln>";
  strXml += " </c:spPr>";
  strXml += "</c:majorGridlines>";
  return strXml;
}
function createLineCap(lineCap) {
  if (!lineCap || lineCap === "flat") {
    return "flat";
  } else if (lineCap === "square") {
    return "sq";
  } else if (lineCap === "round") {
    return "rnd";
  } else {
    const neverLineCap = lineCap;
    throw new Error(`Invalid chart line cap: ${neverLineCap}`);
  }
}

// vendor/pptxgenjs/src/gen-media.ts
function encodeSlideMediaRels(layout) {
  const isNode = typeof process !== "undefined" && !!process.versions?.node && process.release?.name === "node";
  let fs;
  let https;
  const loadNodeDeps = isNode ? async () => {
    ;
    ({ default: fs } = await import("node:fs"));
    ({ default: https } = await import("node:https"));
  } : async () => {
  };
  if (isNode) loadNodeDeps();
  const imageProms = [];
  const candidateRels = layout._relsMedia.filter(
    (rel) => rel.type !== "online" && !rel.data && (!rel.path || rel.path && !rel.path.includes("preencoded"))
  );
  const unqPaths = [];
  candidateRels.forEach((rel) => {
    if (!unqPaths.includes(rel.path)) {
      rel.isDuplicate = false;
      unqPaths.push(rel.path);
    } else {
      rel.isDuplicate = true;
    }
  });
  candidateRels.filter((rel) => !rel.isDuplicate).forEach((rel) => {
    imageProms.push(
      (async () => {
        if (!https) await loadNodeDeps();
        if (isNode && fs && rel.path.indexOf("http") !== 0) {
          try {
            const bitmap = fs.readFileSync(rel.path);
            rel.data = Buffer.from(bitmap).toString("base64");
            candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path).forEach((dupe) => dupe.data = rel.data);
            return "done";
          } catch (ex) {
            rel.data = IMG_BROKEN;
            candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path).forEach((dupe) => dupe.data = rel.data);
            throw new Error(`ERROR: Unable to read media: "${rel.path}"
${String(ex)}`);
          }
        }
        if (isNode && https && rel.path.startsWith("http")) {
          return await new Promise((resolve3, reject) => {
            https.get(rel.path, (res) => {
              let raw = "";
              res.setEncoding("binary");
              res.on("data", (chunk) => raw += chunk);
              res.on("end", () => {
                rel.data = Buffer.from(raw, "binary").toString("base64");
                candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path).forEach((dupe) => dupe.data = rel.data);
                resolve3("done");
              });
              res.on("error", () => {
                rel.data = IMG_BROKEN;
                candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path).forEach((dupe) => dupe.data = rel.data);
                reject(new Error(`ERROR! Unable to load image (https.get): ${rel.path}`));
              });
            });
          });
        }
        return await new Promise((resolve3, reject) => {
          const xhr = new XMLHttpRequest();
          xhr.onload = () => {
            const reader = new FileReader();
            reader.onloadend = () => {
              rel.data = reader.result;
              candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path).forEach((dupe) => dupe.data = rel.data);
              if (!rel.isSvgPng) {
                resolve3("done");
              } else {
                createSvgPngPreview(rel).then(() => resolve3("done")).catch(reject);
              }
            };
            reader.readAsDataURL(xhr.response);
          };
          xhr.onerror = () => {
            rel.data = IMG_BROKEN;
            candidateRels.filter((dupe) => dupe.isDuplicate && dupe.path === rel.path).forEach((dupe) => dupe.data = rel.data);
            reject(new Error(`ERROR! Unable to load image (xhr.onerror): ${rel.path}`));
          };
          xhr.open("GET", rel.path);
          xhr.responseType = "blob";
          xhr.send();
        });
      })()
    );
  });
  layout._relsMedia.filter((rel) => rel.isSvgPng && rel.data).forEach((rel) => {
    (async () => {
      if (isNode && !fs) await loadNodeDeps();
      if (isNode && fs) {
        rel.data = IMG_BROKEN;
        imageProms.push(Promise.resolve("done"));
      } else {
        imageProms.push(createSvgPngPreview(rel));
      }
    })();
  });
  return imageProms;
}
async function createSvgPngPreview(rel) {
  return await new Promise((resolve3, reject) => {
    const image = new Image();
    image.onload = () => {
      if (image.width + image.height === 0) {
        image.onerror("h/w=0");
      }
      let canvas = document.createElement("CANVAS");
      const ctx = canvas.getContext("2d");
      canvas.width = image.width;
      canvas.height = image.height;
      ctx.drawImage(image, 0, 0);
      try {
        rel.data = canvas.toDataURL(rel.type);
        resolve3("done");
      } catch (ex) {
        image.onerror(ex.toString());
      }
      canvas = null;
    };
    image.onerror = () => {
      rel.data = IMG_BROKEN;
      reject(new Error(`ERROR! Unable to load image (image.onerror): ${rel.path}`));
    };
    image.src = typeof rel.data === "string" ? rel.data : IMG_BROKEN;
  });
}

// vendor/pptxgenjs/src/gen-xml.ts
var ImageSizingXml = {
  cover: function(imgSize, boxDim) {
    const imgRatio = imgSize.h / imgSize.w;
    const boxRatio = boxDim.h / boxDim.w;
    const isBoxBased = boxRatio > imgRatio;
    const width = isBoxBased ? boxDim.h / imgRatio : boxDim.w;
    const height = isBoxBased ? boxDim.h : boxDim.w * imgRatio;
    const hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width));
    const vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
    return `<a:srcRect l="${hzPerc}" r="${hzPerc}" t="${vzPerc}" b="${vzPerc}"/><a:stretch/>`;
  },
  contain: function(imgSize, boxDim) {
    const imgRatio = imgSize.h / imgSize.w;
    const boxRatio = boxDim.h / boxDim.w;
    const widthBased = boxRatio > imgRatio;
    const width = widthBased ? boxDim.w : boxDim.h / imgRatio;
    const height = widthBased ? boxDim.w * imgRatio : boxDim.h;
    const hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width));
    const vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
    return `<a:srcRect l="${hzPerc}" r="${hzPerc}" t="${vzPerc}" b="${vzPerc}"/><a:stretch/>`;
  },
  crop: function(imgSize, boxDim) {
    const l = boxDim.x;
    const r = imgSize.w - (boxDim.x + boxDim.w);
    const t = boxDim.y;
    const b = imgSize.h - (boxDim.y + boxDim.h);
    const lPerc = Math.round(1e5 * (l / imgSize.w));
    const rPerc = Math.round(1e5 * (r / imgSize.w));
    const tPerc = Math.round(1e5 * (t / imgSize.h));
    const bPerc = Math.round(1e5 * (b / imgSize.h));
    return `<a:srcRect l="${lPerc}" r="${rPerc}" t="${tPerc}" b="${bPerc}"/><a:stretch/>`;
  }
};
function slideObjectToXml(slide) {
  let strSlideXml = slide._name ? '<p:cSld name="' + slide._name + '">' : "<p:cSld>";
  let intTableNum = 1;
  if (slide._bkgdImgRid) {
    strSlideXml += `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId${slide._bkgdImgRid}"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
  } else if (slide.background?.color) {
    strSlideXml += `<p:bg><p:bgPr>${genXmlColorSelection(slide.background)}</p:bgPr></p:bg>`;
  } else if (!slide.bkgd && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
    strSlideXml += '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
  }
  strSlideXml += "<p:spTree>";
  strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
  strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>';
  strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  slide._slideObjects.forEach((slideItemObj, idx) => {
    let x = 0;
    let y = 0;
    let cx = getSmartParseNumber("75%", "X", slide._presLayout);
    let cy = 0;
    let placeholderObj;
    let locationAttr = "";
    let arrTabRows = null;
    let objTabOpts = null;
    let intColCnt = 0;
    let intColW = 0;
    let cellOpts = null;
    let strXml = null;
    const sizing = slideItemObj.options?.sizing;
    const rounding = slideItemObj.options?.rounding;
    if (slide._slideLayout !== void 0 && slide._slideLayout._slideObjects !== void 0 && slideItemObj.options && slideItemObj.options.placeholder) {
      placeholderObj = slide._slideLayout._slideObjects.filter(
        (object) => object.options.placeholder === slideItemObj.options.placeholder
      )[0];
    }
    slideItemObj.options = slideItemObj.options || {};
    if (typeof slideItemObj.options.x !== "undefined") x = getSmartParseNumber(slideItemObj.options.x, "X", slide._presLayout);
    if (typeof slideItemObj.options.y !== "undefined") y = getSmartParseNumber(slideItemObj.options.y, "Y", slide._presLayout);
    if (typeof slideItemObj.options.w !== "undefined") cx = getSmartParseNumber(slideItemObj.options.w, "X", slide._presLayout);
    if (typeof slideItemObj.options.h !== "undefined") cy = getSmartParseNumber(slideItemObj.options.h, "Y", slide._presLayout);
    let imgWidth = cx;
    let imgHeight = cy;
    if (placeholderObj) {
      if (placeholderObj.options.x || placeholderObj.options.x === 0) x = getSmartParseNumber(placeholderObj.options.x, "X", slide._presLayout);
      if (placeholderObj.options.y || placeholderObj.options.y === 0) y = getSmartParseNumber(placeholderObj.options.y, "Y", slide._presLayout);
      if (placeholderObj.options.w || placeholderObj.options.w === 0) cx = getSmartParseNumber(placeholderObj.options.w, "X", slide._presLayout);
      if (placeholderObj.options.h || placeholderObj.options.h === 0) cy = getSmartParseNumber(placeholderObj.options.h, "Y", slide._presLayout);
    }
    if (slideItemObj.options.flipH) locationAttr += ' flipH="1"';
    if (slideItemObj.options.flipV) locationAttr += ' flipV="1"';
    if (slideItemObj.options.rotate) locationAttr += ` rot="${convertRotationDegrees(slideItemObj.options.rotate)}"`;
    switch (slideItemObj._type) {
      case "table" /* table */:
        arrTabRows = slideItemObj.arrTabRows;
        objTabOpts = slideItemObj.options;
        intColCnt = 0;
        intColW = 0;
        arrTabRows[0].forEach((cell) => {
          cellOpts = cell.options || null;
          intColCnt += cellOpts?.colspan ? Number(cellOpts.colspan) : 1;
        });
        strXml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${intTableNum * slide._slideNum + 1}" name="${slideItemObj.options.objectName}"/>`;
        strXml += '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>  <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr></p:nvGraphicFramePr>';
        strXml += `<p:xfrm><a:off x="${x || (x === 0 ? 0 : EMU)}" y="${y || (y === 0 ? 0 : EMU)}"/><a:ext cx="${cx || (cx === 0 ? 0 : EMU)}" cy="${cy || EMU}"/></p:xfrm>`;
        strXml += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>';
        if (Array.isArray(objTabOpts.colW)) {
          strXml += "<a:tblGrid>";
          for (let col = 0; col < intColCnt; col++) {
            let w = inch2Emu(objTabOpts.colW[col]);
            if (w == null || isNaN(w)) {
              w = (typeof slideItemObj.options.w === "number" ? slideItemObj.options.w : 1) / intColCnt;
            }
            strXml += `<a:gridCol w="${Math.round(w)}"/>`;
          }
          strXml += "</a:tblGrid>";
        } else {
          intColW = objTabOpts.colW ? objTabOpts.colW : EMU;
          if (slideItemObj.options.w && !objTabOpts.colW) intColW = Math.round((typeof slideItemObj.options.w === "number" ? slideItemObj.options.w : 1) / intColCnt);
          strXml += "<a:tblGrid>";
          for (let colw = 0; colw < intColCnt; colw++) {
            strXml += `<a:gridCol w="${intColW}"/>`;
          }
          strXml += "</a:tblGrid>";
        }
        arrTabRows.forEach((cells) => {
          for (let cIdx = 0; cIdx < cells.length; ) {
            const cell = cells[cIdx];
            const colspan = cell.options?.colspan;
            const rowspan = cell.options?.rowspan;
            if (colspan && colspan > 1) {
              const vMergeCells = new Array(colspan - 1).fill(void 0).map(() => {
                return { _type: "tablecell" /* tablecell */, options: { rowspan }, _hmerge: true };
              });
              cells.splice(cIdx + 1, 0, ...vMergeCells);
              cIdx += colspan;
            } else {
              cIdx += 1;
            }
          }
        });
        arrTabRows.forEach((cells, rIdx) => {
          const nextRow = arrTabRows[rIdx + 1];
          if (!nextRow) return;
          cells.forEach((cell, cIdx) => {
            const rowspan = cell._rowContinue || cell.options?.rowspan;
            const colspan = cell.options?.colspan;
            const _hmerge = cell._hmerge;
            if (rowspan && rowspan > 1) {
              const hMergeCell = { _type: "tablecell" /* tablecell */, options: { colspan }, _rowContinue: rowspan - 1, _vmerge: true, _hmerge };
              nextRow.splice(cIdx, 0, hMergeCell);
            }
          });
        });
        arrTabRows.forEach((cells, rIdx) => {
          let intRowH = 0;
          if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]));
          else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH));
          else if (slideItemObj.options.cy || slideItemObj.options.h) {
            intRowH = Math.round(
              (slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : typeof slideItemObj.options.cy === "number" ? slideItemObj.options.cy : 1) / arrTabRows.length
            );
          }
          strXml += `<a:tr h="${intRowH}">`;
          cells.forEach((cellObj) => {
            const cell = cellObj;
            const cellSpanAttrs = {
              rowSpan: cell.options?.rowspan > 1 ? cell.options.rowspan : void 0,
              gridSpan: cell.options?.colspan > 1 ? cell.options.colspan : void 0,
              vMerge: cell._vmerge ? 1 : void 0,
              hMerge: cell._hmerge ? 1 : void 0
            };
            let cellSpanAttrStr = Object.keys(cellSpanAttrs).map((k) => [k, cellSpanAttrs[k]]).filter(([, v]) => !!v).map(([k, v]) => `${String(k)}="${String(v)}"`).join(" ");
            if (cellSpanAttrStr) cellSpanAttrStr = " " + cellSpanAttrStr;
            if (cell._hmerge || cell._vmerge) {
              strXml += `<a:tc${cellSpanAttrStr}><a:tcPr/></a:tc>`;
              return;
            }
            const cellOpts2 = cell.options || {};
            cell.options = cellOpts2;
            ["align", "bold", "border", "color", "fill", "fontFace", "fontSize", "margin", "textDirection", "underline", "valign"].forEach((name) => {
              if (objTabOpts[name] && !cellOpts2[name] && cellOpts2[name] !== 0) cellOpts2[name] = objTabOpts[name];
            });
            const cellValign = cellOpts2.valign ? ` anchor="${cellOpts2.valign.replace(/^c$/i, "ctr").replace(/^m$/i, "ctr").replace("center", "ctr").replace("middle", "ctr").replace("top", "t").replace("btm", "b").replace("bottom", "b")}"` : "";
            const cellTextDir = cellOpts2.textDirection && cellOpts2.textDirection !== "horz" ? ` vert="${cellOpts2.textDirection}"` : "";
            let fillColor = cell._optImp?.fill?.color ? cell._optImp.fill.color : cell._optImp?.fill && typeof cell._optImp.fill === "string" ? cell._optImp.fill : "";
            fillColor = fillColor || cellOpts2.fill ? cellOpts2.fill : "";
            const cellFill = fillColor ? genXmlColorSelection(fillColor) : "";
            let cellMargin = cellOpts2.margin === 0 || cellOpts2.margin ? cellOpts2.margin : DEF_CELL_MARGIN_IN;
            if (!Array.isArray(cellMargin) && typeof cellMargin === "number") cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin];
            let cellMarginXml = "";
            if (cellMargin[0] >= 1) {
              cellMarginXml = ` marL="${valToPts(cellMargin[3])}" marR="${valToPts(cellMargin[1])}" marT="${valToPts(cellMargin[0])}" marB="${valToPts(
                cellMargin[2]
              )}"`;
            } else {
              cellMarginXml = ` marL="${inch2Emu(cellMargin[3])}" marR="${inch2Emu(cellMargin[1])}" marT="${inch2Emu(cellMargin[0])}" marB="${inch2Emu(
                cellMargin[2]
              )}"`;
            }
            strXml += `<a:tc${cellSpanAttrStr}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`;
            if (cellOpts2.border && Array.isArray(cellOpts2.border)) {
              [
                { idx: 3, name: "lnL" },
                { idx: 1, name: "lnR" },
                { idx: 0, name: "lnT" },
                { idx: 2, name: "lnB" }
              ].forEach((obj) => {
                if (cellOpts2.border[obj.idx].type !== "none") {
                  strXml += `<a:${obj.name} w="${valToPts(cellOpts2.border[obj.idx].pt)}" cap="flat" cmpd="sng" algn="ctr">`;
                  strXml += `<a:solidFill>${createColorElement(cellOpts2.border[obj.idx].color)}</a:solidFill>`;
                  strXml += `<a:prstDash val="${cellOpts2.border[obj.idx].type === "dash" ? "sysDash" : "solid"}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>`;
                  strXml += `</a:${obj.name}>`;
                } else {
                  strXml += `<a:${obj.name} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${obj.name}>`;
                }
              });
            }
            strXml += cellFill;
            strXml += "  </a:tcPr>";
            strXml += " </a:tc>";
          });
          strXml += "</a:tr>";
        });
        strXml += "      </a:tbl>";
        strXml += "    </a:graphicData>";
        strXml += "  </a:graphic>";
        strXml += "</p:graphicFrame>";
        strSlideXml += strXml;
        intTableNum++;
        break;
      case "text" /* text */:
      case "placeholder" /* placeholder */:
        if (!slideItemObj.options.line && cy === 0) cy = EMU * 0.3;
        if (!slideItemObj.options._bodyProp) slideItemObj.options._bodyProp = {};
        if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
          slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin[0] || 0);
          slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin[1] || 0);
          slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin[2] || 0);
          slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin[3] || 0);
        } else if (typeof slideItemObj.options.margin === "number") {
          slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin);
          slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin);
          slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin);
          slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin);
        }
        strSlideXml += "<p:sp>";
        strSlideXml += `<p:nvSpPr><p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}">`;
        if (slideItemObj.options.hyperlink?.url) {
          strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.options.hyperlink._rId}" tooltip="${slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : ""}"/>`;
        }
        if (slideItemObj.options.hyperlink?.slide) {
          strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.options.hyperlink._rId}" tooltip="${slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : ""}" action="ppaction://hlinksldjump"/>`;
        }
        strSlideXml += "</p:cNvPr>";
        strSlideXml += "<p:cNvSpPr" + (slideItemObj.options?.isTextBox ? ' txBox="1"/>' : "/>");
        strSlideXml += `<p:nvPr>${slideItemObj._type === "placeholder" ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj)}</p:nvPr>`;
        strSlideXml += "</p:nvSpPr><p:spPr>";
        strSlideXml += `<a:xfrm${locationAttr}>`;
        strSlideXml += `<a:off x="${x}" y="${y}"/>`;
        strSlideXml += `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
        if (slideItemObj.shape === "custGeom") {
          strSlideXml += "<a:custGeom><a:avLst />";
          strSlideXml += "<a:gdLst>";
          strSlideXml += "</a:gdLst>";
          strSlideXml += "<a:ahLst />";
          strSlideXml += "<a:cxnLst>";
          strSlideXml += "</a:cxnLst>";
          strSlideXml += '<a:rect l="l" t="t" r="r" b="b" />';
          strSlideXml += "<a:pathLst>";
          strSlideXml += `<a:path w="${cx}" h="${cy}">`;
          slideItemObj.options.points?.forEach((point, i) => {
            if ("curve" in point) {
              switch (point.curve.type) {
                case "arc":
                  strSlideXml += `<a:arcTo hR="${getSmartParseNumber(point.curve.hR, "Y", slide._presLayout)}" wR="${getSmartParseNumber(
                    point.curve.wR,
                    "X",
                    slide._presLayout
                  )}" stAng="${convertRotationDegrees(point.curve.stAng)}" swAng="${convertRotationDegrees(point.curve.swAng)}" />`;
                  break;
                case "cubic":
                  strSlideXml += `<a:cubicBezTo>
									<a:pt x="${getSmartParseNumber(point.curve.x1, "X", slide._presLayout)}" y="${getSmartParseNumber(point.curve.y1, "Y", slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.curve.x2, "X", slide._presLayout)}" y="${getSmartParseNumber(point.curve.y2, "Y", slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.x, "X", slide._presLayout)}" y="${getSmartParseNumber(point.y, "Y", slide._presLayout)}" />
									</a:cubicBezTo>`;
                  break;
                case "quadratic":
                  strSlideXml += `<a:quadBezTo>
									<a:pt x="${getSmartParseNumber(point.curve.x1, "X", slide._presLayout)}" y="${getSmartParseNumber(point.curve.y1, "Y", slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.x, "X", slide._presLayout)}" y="${getSmartParseNumber(point.y, "Y", slide._presLayout)}" />
									</a:quadBezTo>`;
                  break;
                default:
                  break;
              }
            } else if ("close" in point) {
              strSlideXml += "<a:close />";
            } else if (point.moveTo || i === 0) {
              strSlideXml += `<a:moveTo><a:pt x="${getSmartParseNumber(point.x, "X", slide._presLayout)}" y="${getSmartParseNumber(
                point.y,
                "Y",
                slide._presLayout
              )}" /></a:moveTo>`;
            } else {
              strSlideXml += `<a:lnTo><a:pt x="${getSmartParseNumber(point.x, "X", slide._presLayout)}" y="${getSmartParseNumber(
                point.y,
                "Y",
                slide._presLayout
              )}" /></a:lnTo>`;
            }
          });
          strSlideXml += "</a:path>";
          strSlideXml += "</a:pathLst>";
          strSlideXml += "</a:custGeom>";
        } else {
          strSlideXml += '<a:prstGeom prst="' + slideItemObj.shape + '"><a:avLst>';
          if (slideItemObj.options.rectRadius) {
            strSlideXml += `<a:gd name="adj" fmla="val ${Math.round(slideItemObj.options.rectRadius * EMU * 1e5 / Math.min(cx, cy))}"/>`;
          } else if (slideItemObj.options.angleRange) {
            for (let i = 0; i < 2; i++) {
              const angle = slideItemObj.options.angleRange[i];
              strSlideXml += `<a:gd name="adj${i + 1}" fmla="val ${convertRotationDegrees(angle)}" />`;
            }
            if (slideItemObj.options.arcThicknessRatio) {
              strSlideXml += `<a:gd name="adj3" fmla="val ${Math.round(slideItemObj.options.arcThicknessRatio * 5e4)}" />`;
            }
          }
          strSlideXml += "</a:avLst></a:prstGeom>";
        }
        strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : "<a:noFill/>";
        if (slideItemObj.options.line) {
          strSlideXml += slideItemObj.options.line.width ? `<a:ln w="${valToPts(slideItemObj.options.line.width)}">` : "<a:ln>";
          if (slideItemObj.options.line.color) strSlideXml += genXmlColorSelection(slideItemObj.options.line);
          if (slideItemObj.options.line.dashType) strSlideXml += `<a:prstDash val="${slideItemObj.options.line.dashType}"/>`;
          if (slideItemObj.options.line.beginArrowType) strSlideXml += `<a:headEnd type="${slideItemObj.options.line.beginArrowType}"/>`;
          if (slideItemObj.options.line.endArrowType) strSlideXml += `<a:tailEnd type="${slideItemObj.options.line.endArrowType}"/>`;
          strSlideXml += "</a:ln>";
        }
        if (slideItemObj.options.shadow && slideItemObj.options.shadow.type !== "none") {
          slideItemObj.options.shadow.type = slideItemObj.options.shadow.type || "outer";
          slideItemObj.options.shadow.blur = valToPts(slideItemObj.options.shadow.blur || 8);
          slideItemObj.options.shadow.offset = valToPts(slideItemObj.options.shadow.offset || 4);
          slideItemObj.options.shadow.angle = Math.round((slideItemObj.options.shadow.angle || 270) * 6e4);
          slideItemObj.options.shadow.opacity = Math.round((slideItemObj.options.shadow.opacity || 0.75) * 1e5);
          slideItemObj.options.shadow.color = slideItemObj.options.shadow.color || DEF_TEXT_SHADOW.color;
          strSlideXml += "<a:effectLst>";
          strSlideXml += ` <a:${slideItemObj.options.shadow.type}Shdw ${slideItemObj.options.shadow.type === "outer" ? 'sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="0"' : ""} blurRad="${slideItemObj.options.shadow.blur}" dist="${slideItemObj.options.shadow.offset}" dir="${slideItemObj.options.shadow.angle}">`;
          strSlideXml += ` <a:srgbClr val="${slideItemObj.options.shadow.color}">`;
          strSlideXml += ` <a:alpha val="${slideItemObj.options.shadow.opacity}"/></a:srgbClr>`;
          strSlideXml += " </a:outerShdw>";
          strSlideXml += "</a:effectLst>";
        }
        strSlideXml += "</p:spPr>";
        strSlideXml += genXmlTextBody(slideItemObj);
        strSlideXml += "</p:sp>";
        break;
      case "image" /* image */:
        strSlideXml += "<p:pic>";
        strSlideXml += "  <p:nvPicPr>";
        strSlideXml += `<p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(
          slideItemObj.options.altText || slideItemObj.image
        )}">`;
        if (slideItemObj.hyperlink?.url) {
          strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.hyperlink._rId}" tooltip="${slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : ""}"/>`;
        }
        if (slideItemObj.hyperlink?.slide) {
          strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.hyperlink._rId}" tooltip="${slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : ""}" action="ppaction://hlinksldjump"/>`;
        }
        strSlideXml += "    </p:cNvPr>";
        strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
        strSlideXml += "    <p:nvPr>" + genXmlPlaceholder(placeholderObj) + "</p:nvPr>";
        strSlideXml += "  </p:nvPicPr>";
        strSlideXml += "<p:blipFill>";
        if ((slide._relsMedia || []).filter((rel) => rel.rId === slideItemObj.imageRid)[0] && (slide._relsMedia || []).filter((rel) => rel.rId === slideItemObj.imageRid)[0].extn === "svg") {
          strSlideXml += `<a:blip r:embed="rId${slideItemObj.imageRid - 1}">`;
          strSlideXml += slideItemObj.options.transparency ? ` <a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * 1e3)}"/>` : "";
          strSlideXml += " <a:extLst>";
          strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">';
          strSlideXml += `   <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId${slideItemObj.imageRid}"/>`;
          strSlideXml += "  </a:ext>";
          strSlideXml += " </a:extLst>";
          strSlideXml += "</a:blip>";
        } else {
          strSlideXml += `<a:blip r:embed="rId${slideItemObj.imageRid}">`;
          strSlideXml += slideItemObj.options.transparency ? `<a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * 1e3)}"/>` : "";
          strSlideXml += "</a:blip>";
        }
        if (sizing?.type) {
          const boxW = sizing.w ? getSmartParseNumber(sizing.w, "X", slide._presLayout) : cx;
          const boxH = sizing.h ? getSmartParseNumber(sizing.h, "Y", slide._presLayout) : cy;
          const boxX = getSmartParseNumber(sizing.x || 0, "X", slide._presLayout);
          const boxY = getSmartParseNumber(sizing.y || 0, "Y", slide._presLayout);
          strSlideXml += ImageSizingXml[sizing.type]({ w: imgWidth, h: imgHeight }, { w: boxW, h: boxH, x: boxX, y: boxY });
          imgWidth = boxW;
          imgHeight = boxH;
        } else {
          strSlideXml += "  <a:stretch><a:fillRect/></a:stretch>";
        }
        strSlideXml += "</p:blipFill>";
        strSlideXml += "<p:spPr>";
        strSlideXml += " <a:xfrm" + locationAttr + ">";
        strSlideXml += `  <a:off x="${x}" y="${y}"/>`;
        strSlideXml += `  <a:ext cx="${imgWidth}" cy="${imgHeight}"/>`;
        strSlideXml += " </a:xfrm>";
        strSlideXml += ` <a:prstGeom prst="${rounding ? "ellipse" : "rect"}"><a:avLst/></a:prstGeom>`;
        if (slideItemObj.options.shadow && slideItemObj.options.shadow.type !== "none") {
          slideItemObj.options.shadow.type = slideItemObj.options.shadow.type || "outer";
          slideItemObj.options.shadow.blur = valToPts(slideItemObj.options.shadow.blur || 8);
          slideItemObj.options.shadow.offset = valToPts(slideItemObj.options.shadow.offset || 4);
          slideItemObj.options.shadow.angle = Math.round((slideItemObj.options.shadow.angle || 270) * 6e4);
          slideItemObj.options.shadow.opacity = Math.round((slideItemObj.options.shadow.opacity || 0.75) * 1e5);
          slideItemObj.options.shadow.color = slideItemObj.options.shadow.color || DEF_TEXT_SHADOW.color;
          strSlideXml += "<a:effectLst>";
          strSlideXml += `<a:${slideItemObj.options.shadow.type}Shdw ${slideItemObj.options.shadow.type === "outer" ? 'sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="0"' : ""} blurRad="${slideItemObj.options.shadow.blur}" dist="${slideItemObj.options.shadow.offset}" dir="${slideItemObj.options.shadow.angle}">`;
          strSlideXml += `<a:srgbClr val="${slideItemObj.options.shadow.color}">`;
          strSlideXml += `<a:alpha val="${slideItemObj.options.shadow.opacity}"/></a:srgbClr>`;
          strSlideXml += `</a:${slideItemObj.options.shadow.type}Shdw>`;
          strSlideXml += "</a:effectLst>";
        }
        strSlideXml += "</p:spPr>";
        strSlideXml += "</p:pic>";
        break;
      case "media" /* media */:
        if (slideItemObj.mtype === "online") {
          strSlideXml += "<p:pic>";
          strSlideXml += " <p:nvPicPr>";
          strSlideXml += `<p:cNvPr id="${slideItemObj.mediaRid + 2}" name="${slideItemObj.options.objectName}"/>`;
          strSlideXml += " <p:cNvPicPr/>";
          strSlideXml += " <p:nvPr>";
          strSlideXml += `  <a:videoFile r:link="rId${slideItemObj.mediaRid}"/>`;
          strSlideXml += " </p:nvPr>";
          strSlideXml += " </p:nvPicPr>";
          strSlideXml += ` <p:blipFill><a:blip r:embed="rId${slideItemObj.mediaRid + 1}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
          strSlideXml += " <p:spPr>";
          strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
          strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
          strSlideXml += " </p:spPr>";
          strSlideXml += "</p:pic>";
        } else {
          strSlideXml += "<p:pic>";
          strSlideXml += " <p:nvPicPr>";
          strSlideXml += `<p:cNvPr id="${slideItemObj.mediaRid + 2}" name="${slideItemObj.options.objectName}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>`;
          strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
          strSlideXml += " <p:nvPr>";
          strSlideXml += `  <a:videoFile r:link="rId${slideItemObj.mediaRid}"/>`;
          strSlideXml += "  <p:extLst>";
          strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">';
          strSlideXml += `    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId${slideItemObj.mediaRid + 1}"/>`;
          strSlideXml += "   </p:ext>";
          strSlideXml += "  </p:extLst>";
          strSlideXml += " </p:nvPr>";
          strSlideXml += " </p:nvPicPr>";
          strSlideXml += ` <p:blipFill><a:blip r:embed="rId${slideItemObj.mediaRid + 2}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
          strSlideXml += " <p:spPr>";
          strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
          strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
          strSlideXml += " </p:spPr>";
          strSlideXml += "</p:pic>";
        }
        break;
      case "chart" /* chart */:
        strSlideXml += "<p:graphicFrame>";
        strSlideXml += " <p:nvGraphicFramePr>";
        strSlideXml += `   <p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || "")}"/>`;
        strSlideXml += "   <p:cNvGraphicFramePr/>";
        strSlideXml += `   <p:nvPr>${genXmlPlaceholder(placeholderObj)}</p:nvPr>`;
        strSlideXml += " </p:nvGraphicFramePr>";
        strSlideXml += ` <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`;
        strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
        strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">';
        strSlideXml += `   <c:chart r:id="rId${slideItemObj.chartRid}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`;
        strSlideXml += "  </a:graphicData>";
        strSlideXml += " </a:graphic>";
        strSlideXml += "</p:graphicFrame>";
        break;
      default:
        strSlideXml += "";
        break;
    }
  });
  if (slide._slideNumberProps) {
    if (!slide._slideNumberProps.align) slide._slideNumberProps.align = "left";
    strSlideXml += "<p:sp>";
    strSlideXml += " <p:nvSpPr>";
    strSlideXml += '  <p:cNvPr id="25" name="Slide Number Placeholder 0"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>';
    strSlideXml += '  <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>';
    strSlideXml += " </p:nvSpPr>";
    strSlideXml += " <p:spPr>";
    strSlideXml += `<a:xfrm><a:off x="${getSmartParseNumber(slide._slideNumberProps.x, "X", slide._presLayout)}" y="${getSmartParseNumber(slide._slideNumberProps.y, "Y", slide._presLayout)}"/><a:ext cx="${slide._slideNumberProps.w ? getSmartParseNumber(slide._slideNumberProps.w, "X", slide._presLayout) : "800000"}" cy="${slide._slideNumberProps.h ? getSmartParseNumber(slide._slideNumberProps.h, "Y", slide._presLayout) : "300000"}"/></a:xfrm> <a:prstGeom prst="rect"><a:avLst/></a:prstGeom> <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst></p:spPr>`;
    strSlideXml += "<p:txBody>";
    strSlideXml += "<a:bodyPr";
    if (slide._slideNumberProps.margin && Array.isArray(slide._slideNumberProps.margin)) {
      strSlideXml += ` lIns="${valToPts(slide._slideNumberProps.margin[3] || 0)}"`;
      strSlideXml += ` tIns="${valToPts(slide._slideNumberProps.margin[0] || 0)}"`;
      strSlideXml += ` rIns="${valToPts(slide._slideNumberProps.margin[1] || 0)}"`;
      strSlideXml += ` bIns="${valToPts(slide._slideNumberProps.margin[2] || 0)}"`;
    } else if (typeof slide._slideNumberProps.margin === "number") {
      strSlideXml += ` lIns="${valToPts(slide._slideNumberProps.margin || 0)}"`;
      strSlideXml += ` tIns="${valToPts(slide._slideNumberProps.margin || 0)}"`;
      strSlideXml += ` rIns="${valToPts(slide._slideNumberProps.margin || 0)}"`;
      strSlideXml += ` bIns="${valToPts(slide._slideNumberProps.margin || 0)}"`;
    }
    if (slide._slideNumberProps.valign) {
      strSlideXml += ` anchor="${slide._slideNumberProps.valign.replace("top", "t").replace("middle", "ctr").replace("bottom", "b")}"`;
    }
    strSlideXml += "/>";
    strSlideXml += "  <a:lstStyle><a:lvl1pPr>";
    if (slide._slideNumberProps.fontFace || slide._slideNumberProps.fontSize || slide._slideNumberProps.color) {
      strSlideXml += `<a:defRPr sz="${Math.round((slide._slideNumberProps.fontSize || 12) * 100)}">`;
      if (slide._slideNumberProps.color) strSlideXml += genXmlColorSelection(slide._slideNumberProps.color);
      if (slide._slideNumberProps.fontFace) {
        strSlideXml += `<a:latin typeface="${slide._slideNumberProps.fontFace}"/><a:ea typeface="${slide._slideNumberProps.fontFace}"/><a:cs typeface="${slide._slideNumberProps.fontFace}"/>`;
      }
      strSlideXml += "</a:defRPr>";
    }
    strSlideXml += "</a:lvl1pPr></a:lstStyle>";
    strSlideXml += "<a:p>";
    if (slide._slideNumberProps.align.startsWith("l")) strSlideXml += '<a:pPr algn="l"/>';
    else if (slide._slideNumberProps.align.startsWith("c")) strSlideXml += '<a:pPr algn="ctr"/>';
    else if (slide._slideNumberProps.align.startsWith("r")) strSlideXml += '<a:pPr algn="r"/>';
    else strSlideXml += '<a:pPr algn="l"/>';
    strSlideXml += `<a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr b="${slide._slideNumberProps.bold ? 1 : 0}" lang="en-US"/>`;
    strSlideXml += `<a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`;
    strSlideXml += "</p:txBody></p:sp>";
  }
  strSlideXml += "</p:spTree>";
  strSlideXml += "</p:cSld>";
  return strSlideXml;
}
function slideObjectRelationsToXml(slide, defaultRels) {
  let lastRid = 0;
  let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  slide._rels.forEach((rel) => {
    lastRid = Math.max(lastRid, rel.rId);
    if (rel.type.toLowerCase().includes("hyperlink")) {
      if (rel.data === "slide") {
        strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${rel.Target}.xml"/>`;
      } else {
        strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${rel.Target}" TargetMode="External"/>`;
      }
    } else if (rel.type.toLowerCase().includes("notesSlide")) {
      strXml += `<Relationship Id="rId${rel.rId}" Target="${rel.Target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>`;
    }
  });
  (slide._relsChart || []).forEach((rel) => {
    lastRid = Math.max(lastRid, rel.rId);
    strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${rel.Target}"/>`;
  });
  (slide._relsMedia || []).forEach((rel) => {
    const relRid = rel.rId.toString();
    lastRid = Math.max(lastRid, rel.rId);
    if (rel.type.toLowerCase().includes("image")) {
      strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel.Target + '"/>';
    } else if (rel.type.toLowerCase().includes("audio")) {
      if (strXml.includes(' Target="' + rel.Target + '"')) {
        strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>';
      } else {
        strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="' + rel.Target + '"/>';
      }
    } else if (rel.type.toLowerCase().includes("video")) {
      if (strXml.includes(' Target="' + rel.Target + '"')) {
        strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>';
      } else {
        strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="' + rel.Target + '"/>';
      }
    } else if (rel.type.toLowerCase().includes("online")) {
      if (strXml.includes(' Target="' + rel.Target + '"')) {
        strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/image" Target="' + rel.Target + '"/>';
      } else {
        strXml += '<Relationship Id="rId' + relRid + '" Target="' + rel.Target + '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
      }
    }
  });
  defaultRels.forEach((rel, idx) => {
    strXml += `<Relationship Id="rId${lastRid + idx + 1}" Type="${rel.type}" Target="${rel.target}"/>`;
  });
  strXml += "</Relationships>";
  return strXml;
}
function genXmlParagraphProperties(textObj, isDefault) {
  let strXmlBullet = "";
  let strXmlLnSpc = "";
  let strXmlParaSpc = "";
  let strXmlTabStops = "";
  const tag = isDefault ? "a:lvl1pPr" : "a:pPr";
  let bulletMarL = valToPts(DEF_BULLET_MARGIN);
  let paragraphPropXml = `<${tag}${textObj.options.rtlMode ? ' rtl="1" ' : ""}`;
  {
    if (textObj.options.align) {
      switch (textObj.options.align) {
        case "left":
          paragraphPropXml += ' algn="l"';
          break;
        case "right":
          paragraphPropXml += ' algn="r"';
          break;
        case "center":
          paragraphPropXml += ' algn="ctr"';
          break;
        case "justify":
          paragraphPropXml += ' algn="just"';
          break;
        default:
          paragraphPropXml += "";
          break;
      }
    }
    if (textObj.options.lineSpacing) {
      strXmlLnSpc = `<a:lnSpc><a:spcPts val="${Math.round(textObj.options.lineSpacing * 100)}"/></a:lnSpc>`;
    } else if (textObj.options.lineSpacingMultiple) {
      strXmlLnSpc = `<a:lnSpc><a:spcPct val="${Math.round(textObj.options.lineSpacingMultiple * 1e5)}"/></a:lnSpc>`;
    }
    if (textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0) {
      paragraphPropXml += ` lvl="${textObj.options.indentLevel}"`;
    }
    if (textObj.options.paraSpaceBefore && !isNaN(Number(textObj.options.paraSpaceBefore)) && textObj.options.paraSpaceBefore > 0) {
      strXmlParaSpc += `<a:spcBef><a:spcPts val="${Math.round(textObj.options.paraSpaceBefore * 100)}"/></a:spcBef>`;
    }
    if (textObj.options.paraSpaceAfter && !isNaN(Number(textObj.options.paraSpaceAfter)) && textObj.options.paraSpaceAfter > 0) {
      strXmlParaSpc += `<a:spcAft><a:spcPts val="${Math.round(textObj.options.paraSpaceAfter * 100)}"/></a:spcAft>`;
    }
    if (typeof textObj.options.bullet === "object") {
      if (textObj?.options?.bullet?.indent) bulletMarL = valToPts(textObj.options.bullet.indent);
      if (textObj.options.bullet.type) {
        if (textObj.options.bullet.type.toString().toLowerCase() === "number") {
          paragraphPropXml += ` marL="${textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL}" indent="-${bulletMarL}"`;
          strXmlBullet = `<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${textObj.options.bullet.style || "arabicPeriod"}" startAt="${textObj.options.bullet.numberStartAt || textObj.options.bullet.startAt || "1"}"/>`;
        }
      } else if (textObj.options.bullet.characterCode) {
        let bulletCode = `&#x${textObj.options.bullet.characterCode};`;
        if (!/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.characterCode)) {
          console.warn("Warning: `bullet.characterCode should be a 4-digit unicode charatcer (ex: 22AB)`!");
          bulletCode = "&#x2022;" /* DEFAULT */;
        }
        paragraphPropXml += ` marL="${textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL}" indent="-${bulletMarL}"`;
        strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>';
      } else if (textObj.options.bullet.code) {
        let bulletCode = `&#x${textObj.options.bullet.code};`;
        if (!/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code)) {
          console.warn("Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!");
          bulletCode = "&#x2022;" /* DEFAULT */;
        }
        paragraphPropXml += ` marL="${textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL}" indent="-${bulletMarL}"`;
        strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>';
      } else {
        paragraphPropXml += ` marL="${textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL}" indent="-${bulletMarL}"`;
        strXmlBullet = `<a:buSzPct val="100000"/><a:buChar char="${"&#x2022;" /* DEFAULT */}"/>`;
      }
    } else if (textObj.options.bullet) {
      paragraphPropXml += ` marL="${textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletMarL + bulletMarL * textObj.options.indentLevel : bulletMarL}" indent="-${bulletMarL}"`;
      strXmlBullet = `<a:buSzPct val="100000"/><a:buChar char="${"&#x2022;" /* DEFAULT */}"/>`;
    } else if (!textObj.options.bullet) {
      paragraphPropXml += ' indent="0" marL="0"';
      strXmlBullet = "<a:buNone/>";
    }
    if (textObj.options.tabStops && Array.isArray(textObj.options.tabStops)) {
      const tabStopsXml = textObj.options.tabStops.map((stop) => `<a:tab pos="${inch2Emu(stop.position || 1)}" algn="${stop.alignment || "l"}"/>`).join("");
      strXmlTabStops = `<a:tabLst>${tabStopsXml}</a:tabLst>`;
    }
    paragraphPropXml += ">" + strXmlLnSpc + strXmlParaSpc + strXmlBullet + strXmlTabStops;
    if (isDefault) paragraphPropXml += genXmlTextRunProperties(textObj.options, true);
    paragraphPropXml += "</" + tag + ">";
  }
  return paragraphPropXml;
}
function genXmlTextRunProperties(opts, isDefault) {
  let runProps = "";
  const runPropsTag = isDefault ? "a:defRPr" : "a:rPr";
  runProps += "<" + runPropsTag + ' lang="' + (opts.lang ? opts.lang : "en-US") + '"' + (opts.lang ? ' altLang="en-US"' : "");
  runProps += opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : "";
  runProps += opts?.bold ? ` b="${opts.bold ? "1" : "0"}"` : "";
  runProps += opts?.italic ? ` i="${opts.italic ? "1" : "0"}"` : "";
  runProps += opts?.strike ? ` strike="${typeof opts.strike === "string" ? opts.strike : "sngStrike"}"` : "";
  if (typeof opts.underline === "object" && opts.underline?.style) {
    runProps += ` u="${opts.underline.style}"`;
  } else if (typeof opts.underline === "string") {
    runProps += ` u="${String(opts.underline)}"`;
  } else if (opts.hyperlink) {
    runProps += ' u="sng"';
  }
  if (opts.baseline) {
    runProps += ` baseline="${Math.round(opts.baseline * 50)}"`;
  } else if (opts.subscript) {
    runProps += ' baseline="-40000"';
  } else if (opts.superscript) {
    runProps += ' baseline="30000"';
  }
  runProps += opts.charSpacing ? ` spc="${Math.round(opts.charSpacing * 100)}" kern="0"` : "";
  runProps += ' dirty="0">';
  if (opts.color || opts.fontFace || opts.outline || typeof opts.underline === "object" && opts.underline.color) {
    if (opts.outline && typeof opts.outline === "object") {
      runProps += `<a:ln w="${valToPts(opts.outline.size || 0.75)}">${genXmlColorSelection(opts.outline.color || "FFFFFF")}</a:ln>`;
    }
    if (opts.color) runProps += genXmlColorSelection({ color: opts.color, transparency: opts.transparency });
    if (opts.highlight) runProps += `<a:highlight>${createColorElement(opts.highlight)}</a:highlight>`;
    if (typeof opts.underline === "object" && opts.underline.color) runProps += `<a:uFill>${genXmlColorSelection(opts.underline.color)}</a:uFill>`;
    if (opts.glow) runProps += `<a:effectLst>${createGlowElement(opts.glow, DEF_TEXT_GLOW)}</a:effectLst>`;
    if (opts.fontFace) {
      runProps += `<a:latin typeface="${opts.fontFace}" pitchFamily="34" charset="0"/><a:ea typeface="${opts.fontFace}" pitchFamily="34" charset="-122"/><a:cs typeface="${opts.fontFace}" pitchFamily="34" charset="-120"/>`;
    }
  }
  if (opts.hyperlink) {
    if (typeof opts.hyperlink !== "object") throw new Error("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ");
    else if (!opts.hyperlink.url && !opts.hyperlink.slide) throw new Error("ERROR: 'hyperlink requires either `url` or `slide`'");
    else if (opts.hyperlink.url) {
      runProps += `<a:hlinkClick r:id="rId${opts.hyperlink._rId}" invalidUrl="" action="" tgtFrame="" tooltip="${opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : ""}" history="1" highlightClick="0" endSnd="0"${opts.color ? ">" : "/>"}`;
    } else if (opts.hyperlink.slide) {
      runProps += `<a:hlinkClick r:id="rId${opts.hyperlink._rId}" action="ppaction://hlinksldjump" tooltip="${opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : ""}"${opts.color ? ">" : "/>"}`;
    }
    if (opts.color) {
      runProps += " <a:extLst>";
      runProps += '  <a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}">';
      runProps += '   <ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/>';
      runProps += "  </a:ext>";
      runProps += " </a:extLst>";
      runProps += "</a:hlinkClick>";
    }
  }
  runProps += `</${runPropsTag}>`;
  return runProps;
}
function genXmlTextRun(textObj) {
  return textObj.text ? `<a:r>${genXmlTextRunProperties(textObj.options, false)}<a:t>${encodeXmlEntities(textObj.text)}</a:t></a:r>` : "";
}
function genXmlBodyProperties(slideObject) {
  let bodyProperties = "<a:bodyPr";
  if (slideObject && slideObject._type === "text" /* text */ && slideObject.options._bodyProp) {
    bodyProperties += slideObject.options._bodyProp.wrap ? ' wrap="square"' : ' wrap="none"';
    if (slideObject.options._bodyProp.lIns || slideObject.options._bodyProp.lIns === 0) bodyProperties += ` lIns="${slideObject.options._bodyProp.lIns}"`;
    if (slideObject.options._bodyProp.tIns || slideObject.options._bodyProp.tIns === 0) bodyProperties += ` tIns="${slideObject.options._bodyProp.tIns}"`;
    if (slideObject.options._bodyProp.rIns || slideObject.options._bodyProp.rIns === 0) bodyProperties += ` rIns="${slideObject.options._bodyProp.rIns}"`;
    if (slideObject.options._bodyProp.bIns || slideObject.options._bodyProp.bIns === 0) bodyProperties += ` bIns="${slideObject.options._bodyProp.bIns}"`;
    bodyProperties += ' rtlCol="0"';
    if (slideObject.options._bodyProp.anchor) bodyProperties += ' anchor="' + slideObject.options._bodyProp.anchor + '"';
    if (slideObject.options._bodyProp.vert) bodyProperties += ' vert="' + slideObject.options._bodyProp.vert + '"';
    bodyProperties += ">";
    if (slideObject.options.fit) {
      if (slideObject.options.fit === "none") bodyProperties += "";
      else if (slideObject.options.fit === "shrink") bodyProperties += "<a:normAutofit/>";
      else if (slideObject.options.fit === "resize") bodyProperties += "<a:spAutoFit/>";
    }
    if (slideObject.options.shrinkText) bodyProperties += "<a:normAutofit/>";
    bodyProperties += slideObject.options._bodyProp.autoFit ? "<a:spAutoFit/>" : "";
    bodyProperties += "</a:bodyPr>";
  } else {
    bodyProperties += ' wrap="square" rtlCol="0">';
    bodyProperties += "</a:bodyPr>";
  }
  return slideObject._type === "tablecell" /* tablecell */ ? "<a:bodyPr/>" : bodyProperties;
}
function genXmlTextBody(slideObj) {
  const opts = slideObj.options || {};
  let tmpTextObjects = [];
  const arrTextObjects = [];
  if (opts && slideObj._type !== "tablecell" /* tablecell */ && (typeof slideObj.text === "undefined" || slideObj.text === null)) return "";
  let strSlideXml = slideObj._type === "tablecell" /* tablecell */ ? "<a:txBody>" : "<p:txBody>";
  {
    strSlideXml += genXmlBodyProperties(slideObj);
    if (opts.h === 0 && opts.line && opts.align) strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>';
    else if (slideObj._type === "placeholder") strSlideXml += `<a:lstStyle>${genXmlParagraphProperties(slideObj, true)}</a:lstStyle>`;
    else strSlideXml += "<a:lstStyle/>";
  }
  if (typeof slideObj.text === "string" || typeof slideObj.text === "number") {
    tmpTextObjects.push({ text: slideObj.text.toString(), options: opts || {} });
  } else if (slideObj.text && !Array.isArray(slideObj.text) && typeof slideObj.text === "object" && Object.keys(slideObj.text).includes("text")) {
    tmpTextObjects.push({ text: slideObj.text || "", options: slideObj.options || {} });
  } else if (Array.isArray(slideObj.text)) {
    tmpTextObjects = slideObj.text.map((item) => ({ text: item.text, options: item.options }));
  }
  tmpTextObjects.forEach((itext, idx) => {
    if (!itext.text) itext.text = "";
    itext.options = itext.options || opts || {};
    if (idx === 0 && itext.options && !itext.options.bullet && opts.bullet) itext.options.bullet = opts.bullet;
    if (typeof itext.text === "string" || typeof itext.text === "number") {
      itext.text = itext.text.toString().replace(/\r*\n/g, CRLF);
    }
    if (itext.text.includes(CRLF) && itext.text.match(/\n$/g) === null) {
      itext.text.split(CRLF).forEach((line) => {
        itext.options.breakLine = true;
        arrTextObjects.push({ text: line, options: itext.options });
      });
    } else {
      arrTextObjects.push(itext);
    }
  });
  const arrLines = [];
  let arrTexts = [];
  arrTextObjects.forEach((textObj, idx) => {
    if (arrTexts.length > 0 && (textObj.options.align || opts.align)) {
      if (textObj.options.align !== arrTextObjects[idx - 1].options.align) {
        arrLines.push(arrTexts);
        arrTexts = [];
      }
    } else if (arrTexts.length > 0 && textObj.options.bullet && arrTexts.length > 0) {
      arrLines.push(arrTexts);
      arrTexts = [];
      textObj.options.breakLine = false;
    }
    arrTexts.push(textObj);
    if (arrTexts.length > 0 && textObj.options.breakLine) {
      if (idx + 1 < arrTextObjects.length) {
        arrLines.push(arrTexts);
        arrTexts = [];
      }
    }
    if (idx + 1 === arrTextObjects.length) arrLines.push(arrTexts);
  });
  arrLines.forEach((line) => {
    let reqsClosingFontSize = false;
    strSlideXml += "<a:p>";
    let paragraphPropXml = `<a:pPr ${line[0].options?.rtlMode ? ' rtl="1" ' : ""}`;
    line.forEach((textObj, idx) => {
      textObj.options._lineIdx = idx;
      if (idx > 0 && textObj.options.softBreakBefore) {
        strSlideXml += "<a:br/>";
      }
      textObj.options.align = textObj.options.align || opts.align;
      textObj.options.lineSpacing = textObj.options.lineSpacing || opts.lineSpacing;
      textObj.options.lineSpacingMultiple = textObj.options.lineSpacingMultiple || opts.lineSpacingMultiple;
      textObj.options.indentLevel = textObj.options.indentLevel || opts.indentLevel;
      textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || opts.paraSpaceBefore;
      textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || opts.paraSpaceAfter;
      paragraphPropXml = genXmlParagraphProperties(textObj, false);
      strSlideXml += paragraphPropXml.replace("<a:pPr></a:pPr>", "");
      Object.entries(opts).filter(([key]) => !(textObj.options.hyperlink && key === "color")).forEach(([key, val]) => {
        if (key !== "bullet" && !textObj.options[key]) textObj.options[key] = val;
      });
      strSlideXml += genXmlTextRun(textObj);
      if (!textObj.text && opts.fontSize || textObj.options.fontSize) {
        reqsClosingFontSize = true;
        opts.fontSize = opts.fontSize || textObj.options.fontSize;
      }
    });
    if (slideObj._type === "tablecell" /* tablecell */ && (opts.fontSize || opts.fontFace)) {
      if (opts.fontFace) {
        strSlideXml += `<a:endParaRPr lang="${opts.lang || "en-US"}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : "") + ' dirty="0">';
        strSlideXml += `<a:latin typeface="${opts.fontFace}" charset="0"/>`;
        strSlideXml += `<a:ea typeface="${opts.fontFace}" charset="0"/>`;
        strSlideXml += `<a:cs typeface="${opts.fontFace}" charset="0"/>`;
        strSlideXml += "</a:endParaRPr>";
      } else {
        strSlideXml += `<a:endParaRPr lang="${opts.lang || "en-US"}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : "") + ' dirty="0"/>';
      }
    } else if (reqsClosingFontSize) {
      strSlideXml += `<a:endParaRPr lang="${opts.lang || "en-US"}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : "") + ' dirty="0"/>';
    } else {
      strSlideXml += `<a:endParaRPr lang="${opts.lang || "en-US"}" dirty="0"/>`;
    }
    strSlideXml += "</a:p>";
  });
  if (strSlideXml.indexOf("<a:p>") === -1) {
    strSlideXml += "<a:p><a:endParaRPr/></a:p>";
  }
  strSlideXml += slideObj._type === "tablecell" /* tablecell */ ? "</a:txBody>" : "</p:txBody>";
  return strSlideXml;
}
function genXmlPlaceholder(placeholderObj) {
  if (!placeholderObj) return "";
  const placeholderIdx = placeholderObj.options?._placeholderIdx ? placeholderObj.options._placeholderIdx : "";
  const placeholderTyp = placeholderObj.options?._placeholderType ? placeholderObj.options._placeholderType : "";
  const placeholderType = placeholderTyp && PLACEHOLDER_TYPES[placeholderTyp] ? PLACEHOLDER_TYPES[placeholderTyp].toString() : "";
  return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx.toString() + '"' : ""}
		${placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ` type="${placeholderType}"` : ""}
		${placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ""}
		/>`;
}
function makeXmlContTypes(slides, slideLayouts, masterSlide) {
  let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
  strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
  strXml += '<Default Extension="xml" ContentType="application/xml"/>';
  strXml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
  strXml += '<Default Extension="jpeg" ContentType="image/jpeg"/>';
  strXml += '<Default Extension="jpg" ContentType="image/jpg"/>';
  strXml += '<Default Extension="svg" ContentType="image/svg+xml"/>';
  strXml += '<Default Extension="png" ContentType="image/png"/>';
  strXml += '<Default Extension="gif" ContentType="image/gif"/>';
  strXml += '<Default Extension="m4v" ContentType="video/mp4"/>';
  strXml += '<Default Extension="mp4" ContentType="video/mp4"/>';
  slides.forEach((slide) => {
    (slide._relsMedia || []).forEach((rel) => {
      if (rel.type !== "image" && rel.type !== "online" && rel.type !== "chart" && rel.extn !== "m4v" && !strXml.includes(rel.type)) {
        strXml += '<Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
      }
    });
  });
  strXml += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
  strXml += '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';
  strXml += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
  strXml += '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>';
  slides.forEach((slide, idx) => {
    strXml += `<Override PartName="/ppt/slideMasters/slideMaster${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`;
    strXml += `<Override PartName="/ppt/slides/slide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    slide._relsChart.forEach((rel) => {
      strXml += `<Override PartName="${rel.Target}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
    });
  });
  strXml += '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>';
  strXml += '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
  strXml += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
  strXml += '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>';
  slideLayouts.forEach((layout, idx) => {
    strXml += `<Override PartName="/ppt/slideLayouts/slideLayout${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;
    (layout._relsChart || []).forEach((rel) => {
      strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
    });
  });
  slides.forEach((_slide, idx) => {
    strXml += `<Override PartName="/ppt/notesSlides/notesSlide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
  });
  masterSlide._relsChart.forEach((rel) => {
    strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
  });
  masterSlide._relsMedia.forEach((rel) => {
    if (rel.type !== "image" && rel.type !== "online" && rel.type !== "chart" && rel.extn !== "m4v" && !strXml.includes(rel.type)) {
      strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
    }
  });
  strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
  strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
  strXml += "</Types>";
  return strXml;
}
function makeXmlRootRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
		</Relationships>`;
}
function makeXmlApp(slides, company) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
	<TotalTime>0</TotalTime>
	<Words>0</Words>
	<Application>Microsoft Office PowerPoint</Application>
	<PresentationFormat>On-screen Show (16:9)</PresentationFormat>
	<Paragraphs>0</Paragraphs>
	<Slides>${slides.length}</Slides>
	<Notes>${slides.length}</Notes>
	<HiddenSlides>0</HiddenSlides>
	<MMClips>0</MMClips>
	<ScaleCrop>false</ScaleCrop>
	<HeadingPairs>
		<vt:vector size="6" baseType="variant">
			<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>2</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>1</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>${slides.length}</vt:i4></vt:variant>
		</vt:vector>
	</HeadingPairs>
	<TitlesOfParts>
		<vt:vector size="${slides.length + 1 + 2}" baseType="lpstr">
			<vt:lpstr>Arial</vt:lpstr>
			<vt:lpstr>Calibri</vt:lpstr>
			<vt:lpstr>Office Theme</vt:lpstr>
			${slides.map((_slideObj, idx) => `<vt:lpstr>Slide ${idx + 1}</vt:lpstr>`).join("")}
		</vt:vector>
	</TitlesOfParts>
	<Company>${company}</Company>
	<LinksUpToDate>false</LinksUpToDate>
	<SharedDoc>false</SharedDoc>
	<HyperlinksChanged>false</HyperlinksChanged>
	<AppVersion>16.0000</AppVersion>
	</Properties>`;
}
function makeXmlCore(title, subject, author, revision) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
		<dc:title>${encodeXmlEntities(title)}</dc:title>
		<dc:subject>${encodeXmlEntities(subject)}</dc:subject>
		<dc:creator>${encodeXmlEntities(author)}</dc:creator>
		<cp:lastModifiedBy>${encodeXmlEntities(author)}</cp:lastModifiedBy>
		<cp:revision>${revision}</cp:revision>
		<dcterms:created xsi:type="dcterms:W3CDTF">${(/* @__PURE__ */ new Date()).toISOString().replace(/\.\d\d\dZ/, "Z")}</dcterms:created>
		<dcterms:modified xsi:type="dcterms:W3CDTF">${(/* @__PURE__ */ new Date()).toISOString().replace(/\.\d\d\dZ/, "Z")}</dcterms:modified>
	</cp:coreProperties>`;
}
function makeXmlPresentationRels(slides) {
  let intRelNum = 1;
  let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
  strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  strXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
  for (let idx = 1; idx <= slides.length; idx++) {
    strXml += `<Relationship Id="rId${++intRelNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${idx}.xml"/>`;
  }
  intRelNum++;
  strXml += `<Relationship Id="rId${intRelNum + 0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/><Relationship Id="rId${intRelNum + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/><Relationship Id="rId${intRelNum + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/><Relationship Id="rId${intRelNum + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId${intRelNum + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/></Relationships>`;
  return strXml;
}
function makeXmlSlide(slide) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${slide?.hidden ? ' show="0"' : ""}>${slideObjectToXml(slide)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` + (slide._timingXml || "") + "</p:sld>";
}
function getNotesFromSlide(slide) {
  let notesText = "";
  slide._slideObjects.forEach((data) => {
    if (data._type === "notes" /* notes */) notesText += data?.text && data.text[0] ? data.text[0].text : "";
  });
  return notesText.replace(/\r*\n/g, CRLF);
}
function makeXmlNotesMaster() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{5282F153-3F37-0F45-9E97-73ACFA13230C}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>7/23/19</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US"/><a:t>Fifth level</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="458787"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213"/><a:ext cx="2971800" cy="458787"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}" type="slidenum"><a:rPr lang="en-US"/><a:t>\u2039#\u203A</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/></p:ext></p:extLst></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>`;
}
function makeXmlNotesSlide(slide) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${encodeXmlEntities(getNotesFromSlide(slide))}</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr lang="en-US"/><a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/></p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}
function makeXmlLayout(layout) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">
		${slideObjectToXml(layout)}
		<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}
function makeXmlMaster(slide, layouts) {
  const layoutDefs = layouts.map((_layoutDef, idx) => `<p:sldLayoutId id="${LAYOUT_IDX_SERIES_BASE + idx}" r:id="rId${slide._rels.length + idx + 1}"/>`);
  let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
  strXml += '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">';
  strXml += slideObjectToXml(slide);
  strXml += '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  strXml += "<p:sldLayoutIdLst>" + layoutDefs.join("") + "</p:sldLayoutIdLst>";
  strXml += '<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>';
  strXml += '<p:txStyles> <p:titleStyle>  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr> </p:titleStyle> <p:bodyStyle>  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2022"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2013"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2022"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2013"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\xBB"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2022"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2022"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2022"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="\u2022"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr> </p:bodyStyle> <p:otherStyle>  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr> </p:otherStyle></p:txStyles>';
  strXml += "</p:sldMaster>";
  return strXml;
}
function makeXmlSlideLayoutRel(layoutNumber, slideLayouts) {
  return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
    {
      target: "../slideMasters/slideMaster1.xml",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    }
  ]);
}
function makeXmlSlideRel(slides, slideLayouts, slideNumber) {
  return slideObjectRelationsToXml(slides[slideNumber - 1], [
    {
      target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    },
    {
      target: `../notesSlides/notesSlide${slideNumber}.xml`,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    }
  ]);
}
function makeXmlNotesSlideRel(slideNumber) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNumber}.xml"/>
		</Relationships>`;
}
function makeXmlMasterRel(masterSlide, slideLayouts) {
  const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
    target: `../slideLayouts/slideLayout${idx + 1}.xml`,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
  }));
  defaultRels.push({ target: "../theme/theme1.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" });
  return slideObjectRelationsToXml(masterSlide, defaultRels);
}
function makeXmlNotesMasterRel() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
		</Relationships>`;
}
function getLayoutIdxForSlide(slides, slideLayouts, slideNumber) {
  for (let i = 0; i < slideLayouts.length; i++) {
    if (slideLayouts[i]._name === slides[slideNumber - 1]._slideLayout._name) {
      return i + 1;
    }
  }
  return 1;
}
function makeXmlTheme(pres) {
  const majorFont = pres.theme?.headFontFace ? `<a:latin typeface="${pres.theme?.headFontFace}"/>` : '<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>';
  const minorFont = pres.theme?.bodyFontFace ? `<a:latin typeface="${pres.theme?.bodyFontFace}"/>` : '<a:latin typeface="Calibri" panose="020F0502020204030204"/>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont>${majorFont}<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="\u6E38\u30B4\u30B7\u30C3\u30AF Light"/><a:font script="Hang" typeface="\uB9D1\uC740 \uACE0\uB515"/><a:font script="Hans" typeface="\u7B49\u7EBF Light"/><a:font script="Hant" typeface="\u65B0\u7D30\u660E\u9AD4"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont>${minorFont}<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="\u6E38\u30B4\u30B7\u30C3\u30AF"/><a:font script="Hang" typeface="\uB9D1\uC740 \uACE0\uB515"/><a:font script="Hans" typeface="\u7B49\u7EBF"/><a:font script="Hant" typeface="\u65B0\u7D30\u660E\u9AD4"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>`;
}
function makeXmlPresentation(pres) {
  let strXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${pres.rtlMode ? 'rtl="1"' : ""} saveSubsetFonts="1" autoCompressPictures="0">`;
  strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
  strXml += "<p:sldIdLst>";
  pres.slides.forEach((slide) => strXml += `<p:sldId id="${slide._slideId}" r:id="rId${slide._rId}"/>`);
  strXml += "</p:sldIdLst>";
  strXml += `<p:notesMasterIdLst><p:notesMasterId r:id="rId${pres.slides.length + 2}"/></p:notesMasterIdLst>`;
  strXml += `<p:sldSz cx="${pres.presLayout.width}" cy="${pres.presLayout.height}"/>`;
  strXml += `<p:notesSz cx="${pres.presLayout.height}" cy="${pres.presLayout.width}"/>`;
  strXml += "<p:defaultTextStyle>";
  for (let idy = 1; idy < 10; idy++) {
    strXml += `<a:lvl${idy}pPr marL="${(idy - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${idy}pPr>`;
  }
  strXml += "</p:defaultTextStyle>";
  if (pres.sections && pres.sections.length > 0) {
    strXml += '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">';
    strXml += '<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">';
    pres.sections.forEach((sect) => {
      strXml += `<p14:section name="${encodeXmlEntities(sect.title)}" id="{${getUuid("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")}}"><p14:sldIdLst>`;
      sect._slides.forEach((slide) => strXml += `<p14:sldId id="${slide._slideId}"/>`);
      strXml += "</p14:sldIdLst></p14:section>";
    });
    strXml += "</p14:sectionLst></p:ext>";
    strXml += '<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext>';
    strXml += "</p:extLst>";
  }
  strXml += "</p:presentation>";
  return strXml;
}
function makeXmlPresProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
}
function makeXmlTableStyles() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}
function makeXmlViewProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`;
}

// vendor/pptxgenjs/src/pptxgen.ts
var VERSION = "3.13.0-beta.2-20250426-1520";
var PptxGenJS2 = class {
  // Property getters/setters
  /**
   * Presentation layout name
   * Standard layouts:
   * - 'LAYOUT_4x3'   (10"    x 7.5")
   * - 'LAYOUT_16x9'  (10"    x 5.625")
   * - 'LAYOUT_16x10' (10"    x 6.25")
   * - 'LAYOUT_WIDE'  (13.33" x 7.5")
   * Custom layouts:
   * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
   * @type {string}
   * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
   */
  _layout;
  set layout(value) {
    const newLayout = this.LAYOUTS[value];
    if (newLayout) {
      this._layout = value;
      this._presLayout = newLayout;
    } else {
      throw new Error("UNKNOWN-LAYOUT");
    }
  }
  get layout() {
    return this._layout;
  }
  /**
   * PptxGenJS Library Version
   */
  _version = VERSION;
  get version() {
    return this._version;
  }
  /**
   * @type {string}
   */
  _author;
  set author(value) {
    this._author = value;
  }
  get author() {
    return this._author;
  }
  /**
   * @type {string}
   */
  _company;
  set company(value) {
    this._company = value;
  }
  get company() {
    return this._company;
  }
  /**
   * @type {string}
   * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
   */
  _revision;
  set revision(value) {
    this._revision = value;
  }
  get revision() {
    return this._revision;
  }
  /**
   * @type {string}
   */
  _subject;
  set subject(value) {
    this._subject = value;
  }
  get subject() {
    return this._subject;
  }
  /**
   * @type {ThemeProps}
   */
  _theme;
  set theme(value) {
    this._theme = value;
  }
  get theme() {
    return this._theme;
  }
  /**
   * @type {string}
   */
  _title;
  set title(value) {
    this._title = value;
  }
  get title() {
    return this._title;
  }
  /**
   * Whether Right-to-Left (RTL) mode is enabled
   * @type {boolean}
   */
  _rtlMode;
  set rtlMode(value) {
    this._rtlMode = value;
  }
  get rtlMode() {
    return this._rtlMode;
  }
  /** master slide layout object */
  _masterSlide;
  get masterSlide() {
    return this._masterSlide;
  }
  /** this Presentation's Slide objects */
  _slides;
  get slides() {
    return this._slides;
  }
  /** this Presentation's sections */
  _sections;
  get sections() {
    return this._sections;
  }
  /** slide layout definition objects, used for generating slide layout files */
  _slideLayouts;
  get slideLayouts() {
    return this._slideLayouts;
  }
  LAYOUTS;
  // Exposed class props
  _alignH = AlignH;
  get AlignH() {
    return this._alignH;
  }
  _alignV = AlignV;
  get AlignV() {
    return this._alignV;
  }
  _chartType = ChartType;
  get ChartType() {
    return this._chartType;
  }
  _outputType = OutputType;
  get OutputType() {
    return this._outputType;
  }
  _presLayout;
  get presLayout() {
    return this._presLayout;
  }
  _schemeColor = SchemeColor;
  get SchemeColor() {
    return this._schemeColor;
  }
  _shapeType = ShapeType;
  get ShapeType() {
    return this._shapeType;
  }
  /**
   * @depricated use `ChartType`
   */
  _charts = CHART_TYPE;
  get charts() {
    return this._charts;
  }
  /**
   * @depricated use `SchemeColor`
   */
  _colors = SCHEME_COLOR_NAMES;
  get colors() {
    return this._colors;
  }
  /**
   * @depricated use `ShapeType`
   */
  _shapes = SHAPE_TYPE;
  get shapes() {
    return this._shapes;
  }
  constructor() {
    const layout4x3 = { name: "screen4x3", width: 9144e3, height: 6858e3 };
    const layout16x9 = { name: "screen16x9", width: 9144e3, height: 5143500 };
    const layout16x10 = { name: "screen16x10", width: 9144e3, height: 5715e3 };
    const layoutWide = { name: "custom", width: 12192e3, height: 6858e3 };
    this.LAYOUTS = {
      LAYOUT_4x3: layout4x3,
      LAYOUT_16x9: layout16x9,
      LAYOUT_16x10: layout16x10,
      LAYOUT_WIDE: layoutWide
    };
    this._author = "PptxGenJS";
    this._company = "PptxGenJS";
    this._revision = "1";
    this._subject = "PptxGenJS Presentation";
    this._title = "PptxGenJS Presentation";
    this._presLayout = {
      name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
      _sizeW: this.LAYOUTS[DEF_PRES_LAYOUT].width,
      _sizeH: this.LAYOUTS[DEF_PRES_LAYOUT].height,
      width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
      height: this.LAYOUTS[DEF_PRES_LAYOUT].height
    };
    this._rtlMode = false;
    this._slideLayouts = [
      {
        _margin: DEF_SLIDE_MARGIN_IN,
        _name: DEF_PRES_LAYOUT_NAME,
        _presLayout: this._presLayout,
        _rels: [],
        _relsChart: [],
        _relsMedia: [],
        _slide: null,
        _slideNum: 1e3,
        _slideNumberProps: null,
        _slideObjects: []
      }
    ];
    this._slides = [];
    this._sections = [];
    this._masterSlide = {
      addChart: null,
      addImage: null,
      addMedia: null,
      addNotes: null,
      addShape: null,
      addTable: null,
      addText: null,
      //
      _name: null,
      _presLayout: this._presLayout,
      _rId: null,
      _rels: [],
      _relsChart: [],
      _relsMedia: [],
      _slideId: null,
      _slideLayout: null,
      _slideNum: null,
      _slideNumberProps: null,
      _slideObjects: []
    };
  }
  /**
   * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
   * @param {AddSlideProps} options - slide masterName and/or sectionTitle
   * @return {PresSlide} new Slide
   */
  addNewSlide = (options) => {
    const sectAlreadyInUse = this.sections.length > 0 && this.sections[this.sections.length - 1]._slides.filter((slide) => slide._slideNum === this.slides[this.slides.length - 1]._slideNum).length > 0;
    options.sectionTitle = sectAlreadyInUse ? this.sections[this.sections.length - 1].title : null;
    return this.addSlide(options);
  };
  /**
   * Provides an API for `addTableDefinition` to get slide reference by number
   * @param {number} slideNum - slide number
   * @return {PresSlide} Slide
   * @since 3.0.0
   */
  getSlide = (slideNum) => this.slides.filter((slide) => slide._slideNum === slideNum)[0];
  /**
   * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
   * @param {SlideNumberProps} slideNum - slide number config
   */
  setSlideNumber = (slideNum) => {
    this.masterSlide._slideNumberProps = slideNum;
    this.slideLayouts.filter((layout) => layout._name === DEF_PRES_LAYOUT_NAME)[0]._slideNumberProps = slideNum;
  };
  /**
   * Create all chart and media rels for this Presentation
   * @param {PresSlide | SlideLayout} slide - slide with rels
   * @param {JSZip} zip - JSZip instance
   * @param {Promise<string>[]} chartPromises - promise array
   */
  createChartMediaRels = (slide, zip, chartPromises) => {
    slide._relsChart.forEach((rel) => chartPromises.push(createExcelWorksheet(rel, zip)));
    slide._relsMedia.forEach((rel) => {
      if (rel.type !== "online" && rel.type !== "hyperlink") {
        let data = rel.data && typeof rel.data === "string" ? rel.data : "";
        if (!data.includes(",") && !data.includes(";")) data = "image/png;base64," + data;
        else if (!data.includes(",")) data = "image/png;base64," + data;
        else if (!data.includes(";")) data = "image/png;" + data;
        zip.file(rel.Target.replace("..", "ppt"), data.split(",").pop(), { base64: true });
      }
    });
  };
  /**
   * Create and export the .pptx file
   * @param {string} exportName - output file type
   * @param {Blob} blobContent - Blob content
   * @return {Promise<string>} Promise with file name
   */
  writeFileToBrowser = async (exportName, blobContent) => {
    const eleLink = document.createElement("a");
    eleLink.setAttribute("style", "display:none;");
    eleLink.dataset.interception = "off";
    document.body.appendChild(eleLink);
    if (window.URL.createObjectURL) {
      const url = window.URL.createObjectURL(new Blob([blobContent], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
      eleLink.href = url;
      eleLink.download = exportName;
      eleLink.click();
      setTimeout(() => {
        window.URL.revokeObjectURL(url);
        document.body.removeChild(eleLink);
      }, 100);
      return await Promise.resolve(exportName);
    }
  };
  /**
   * Create and export the .pptx file
   * @param {WRITE_OUTPUT_TYPE} outputType - output file type
   * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
   */
  exportPresentation = async (props) => {
    const arrChartPromises = [];
    let arrMediaPromises = [];
    const zip = new JSZip2();
    this.slides.forEach((slide) => {
      arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(slide));
    });
    this.slideLayouts.forEach((layout) => {
      arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(layout));
    });
    arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(this.masterSlide));
    return await Promise.all(arrMediaPromises).then(async () => {
      this.slides.forEach((slide) => {
        if (slide._slideLayout) addPlaceholdersToSlideLayouts(slide);
      });
      zip.folder("_rels");
      zip.folder("docProps");
      zip.folder("ppt").folder("_rels");
      zip.folder("ppt/charts").folder("_rels");
      zip.folder("ppt/embeddings");
      zip.folder("ppt/media");
      zip.folder("ppt/slideLayouts").folder("_rels");
      zip.folder("ppt/slideMasters").folder("_rels");
      zip.folder("ppt/slides").folder("_rels");
      zip.folder("ppt/theme");
      zip.folder("ppt/notesMasters").folder("_rels");
      zip.folder("ppt/notesSlides").folder("_rels");
      zip.file("[Content_Types].xml", makeXmlContTypes(this.slides, this.slideLayouts, this.masterSlide));
      zip.file("_rels/.rels", makeXmlRootRels());
      zip.file("docProps/app.xml", makeXmlApp(this.slides, this.company));
      zip.file("docProps/core.xml", makeXmlCore(this.title, this.subject, this.author, this.revision));
      zip.file("ppt/_rels/presentation.xml.rels", makeXmlPresentationRels(this.slides));
      zip.file("ppt/theme/theme1.xml", makeXmlTheme(this));
      zip.file("ppt/presentation.xml", makeXmlPresentation(this));
      zip.file("ppt/presProps.xml", makeXmlPresProps());
      zip.file("ppt/tableStyles.xml", makeXmlTableStyles());
      zip.file("ppt/viewProps.xml", makeXmlViewProps());
      this.slideLayouts.forEach((layout, idx) => {
        zip.file(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, makeXmlLayout(layout));
        zip.file(`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`, makeXmlSlideLayoutRel(idx + 1, this.slideLayouts));
      });
      this.slides.forEach((slide, idx) => {
        zip.file(`ppt/slides/slide${idx + 1}.xml`, makeXmlSlide(slide));
        zip.file(`ppt/slides/_rels/slide${idx + 1}.xml.rels`, makeXmlSlideRel(this.slides, this.slideLayouts, idx + 1));
        zip.file(`ppt/notesSlides/notesSlide${idx + 1}.xml`, makeXmlNotesSlide(slide));
        zip.file(`ppt/notesSlides/_rels/notesSlide${idx + 1}.xml.rels`, makeXmlNotesSlideRel(idx + 1));
      });
      zip.file("ppt/slideMasters/slideMaster1.xml", makeXmlMaster(this.masterSlide, this.slideLayouts));
      zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", makeXmlMasterRel(this.masterSlide, this.slideLayouts));
      zip.file("ppt/notesMasters/notesMaster1.xml", makeXmlNotesMaster());
      zip.file("ppt/notesMasters/_rels/notesMaster1.xml.rels", makeXmlNotesMasterRel());
      this.slideLayouts.forEach((layout) => {
        this.createChartMediaRels(layout, zip, arrChartPromises);
      });
      this.slides.forEach((slide) => {
        this.createChartMediaRels(slide, zip, arrChartPromises);
      });
      this.createChartMediaRels(this.masterSlide, zip, arrChartPromises);
      return await Promise.all(arrChartPromises).then(async () => {
        if (props.outputType === "STREAM") {
          return await zip.generateAsync({ type: "nodebuffer", compression: props.compression ? "DEFLATE" : "STORE" });
        } else if (props.outputType) {
          return await zip.generateAsync({ type: props.outputType });
        } else {
          return await zip.generateAsync({ type: "blob", compression: props.compression ? "DEFLATE" : "STORE" });
        }
      });
    });
  };
  // EXPORT METHODS
  /**
   * Export the current Presentation to stream
   * @param {WriteBaseProps} props - output properties
   * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
   */
  async stream(props) {
    return await this.exportPresentation({
      compression: props?.compression,
      outputType: "STREAM"
    });
  }
  /**
   * Export the current Presentation as JSZip content with the selected type
   * @param {WriteProps} props output properties
   * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
   */
  async write(props) {
    const propsOutpType = typeof props === "object" && props?.outputType ? props.outputType : props ? props : null;
    const propsCompress = typeof props === "object" && props?.compression ? props.compression : false;
    return await this.exportPresentation({
      compression: propsCompress,
      outputType: propsOutpType
    });
  }
  /**
   * Export the current Presentation.
   * Write the generated presentation to disk (Node) or trigger a download (browser).
   * @param {WriteFileProps} props - output file properties
   * @returns {Promise<string>} the presentation name
   */
  async writeFile(props) {
    const isNode = typeof process !== "undefined" && !!process.versions?.node && process.release?.name === "node";
    if (typeof props === "string") {
      console.warn("[WARNING] writeFile(string) is deprecated - pass { fileName } instead.");
      props = { fileName: props };
    }
    const { fileName: rawName = "Presentation.pptx", compression = false } = props;
    const fileName = rawName.toLowerCase().endsWith(".pptx") ? rawName : `${rawName}.pptx`;
    const outputType = isNode ? "nodebuffer" : null;
    const data = await this.exportPresentation({ compression, outputType });
    if (isNode) {
      const { writeFile: writeFile2 } = await import("node:fs/promises");
      await writeFile2(fileName, data);
      return fileName;
    }
    await this.writeFileToBrowser(fileName, data);
    return fileName;
  }
  // PRESENTATION METHODS
  /**
   * Add a new Section to Presentation
   * @param {ISectionProps} section - section properties
   * @example pptx.addSection({ title:'Charts' });
   */
  addSection(section) {
    if (!section) console.warn("addSection requires an argument");
    else if (!section.title) console.warn("addSection requires a title");
    const newSection = {
      _type: "user",
      _slides: [],
      title: section.title
    };
    if (section.order) this.sections.splice(section.order, 0, newSection);
    else this._sections.push(newSection);
  }
  /**
   * Add a new Slide to Presentation
   * @param {AddSlideProps} options - slide options
   * @returns {PresSlide} the new Slide
   */
  addSlide(options) {
    const masterSlideName = typeof options === "string" ? options : options?.masterName ? options.masterName : "";
    let slideLayout = {
      _name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
      _presLayout: this.presLayout,
      _rels: [],
      _relsChart: [],
      _relsMedia: [],
      _slideNum: this.slides.length + 1
    };
    if (masterSlideName) {
      const tmpLayout = this.slideLayouts.filter((layout) => layout._name === masterSlideName)[0];
      if (tmpLayout) slideLayout = tmpLayout;
    }
    const newSlide = new Slide({
      addSlide: this.addNewSlide,
      getSlide: this.getSlide,
      presLayout: this.presLayout,
      setSlideNum: this.setSlideNumber,
      slideId: this.slides.length + 256,
      slideRId: this.slides.length + 2,
      slideNumber: this.slides.length + 1,
      slideLayout
    });
    this._slides.push(newSlide);
    if (options?.sectionTitle) {
      const sect = this.sections.filter((section) => section.title === options.sectionTitle)[0];
      if (!sect) console.warn(`addSlide: unable to find section with title: "${options.sectionTitle}"`);
      else sect._slides.push(newSlide);
    } else if (this.sections && this.sections.length > 0 && !options?.sectionTitle) {
      const lastSect = this._sections[this.sections.length - 1];
      if (lastSect._type === "default") lastSect._slides.push(newSlide);
      else {
        this._sections.push({
          title: `Default-${this.sections.filter((sect) => sect._type === "default").length + 1}`,
          _type: "default",
          _slides: [newSlide]
        });
      }
    }
    return newSlide;
  }
  /**
   * Create a custom Slide Layout in any size
   * @param {PresLayout} layout - layout properties
   * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
   */
  defineLayout(layout) {
    if (!layout) console.warn("defineLayout requires `{name, width, height}`");
    else if (!layout.name) console.warn("defineLayout requires `name`");
    else if (!layout.width) console.warn("defineLayout requires `width`");
    else if (!layout.height) console.warn("defineLayout requires `height`");
    else if (typeof layout.height !== "number") console.warn("defineLayout `height` should be a number (inches)");
    else if (typeof layout.width !== "number") console.warn("defineLayout `width` should be a number (inches)");
    this.LAYOUTS[layout.name] = {
      name: layout.name,
      _sizeW: Math.round(Number(layout.width) * EMU),
      _sizeH: Math.round(Number(layout.height) * EMU),
      width: Math.round(Number(layout.width) * EMU),
      height: Math.round(Number(layout.height) * EMU)
    };
  }
  /**
   * Create a new slide master [layout] for the Presentation
   * @param {SlideMasterProps} props - layout properties
   */
  defineSlideMaster(props) {
    const propsClone = JSON.parse(JSON.stringify(props));
    if (!propsClone.title) throw new Error("defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)");
    const newLayout = {
      _margin: propsClone.margin || DEF_SLIDE_MARGIN_IN,
      _name: propsClone.title,
      _presLayout: this.presLayout,
      _rels: [],
      _relsChart: [],
      _relsMedia: [],
      _slide: null,
      _slideNum: 1e3 + this.slideLayouts.length + 1,
      _slideNumberProps: propsClone.slideNumber || null,
      _slideObjects: [],
      background: propsClone.background || null,
      bkgd: propsClone.bkgd || null
    };
    createSlideMaster(propsClone, newLayout);
    this.slideLayouts.push(newLayout);
    if (propsClone.background || propsClone.bkgd) addBackgroundDefinition(propsClone.background, newLayout);
    if (newLayout._slideNumberProps && !this.masterSlide._slideNumberProps) this.masterSlide._slideNumberProps = newLayout._slideNumberProps;
  }
  // HTML-TO-SLIDES METHODS
  /**
   * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
   * @param {string} eleId - table HTML element ID
   * @param {TableToSlidesProps} options - generation options
   */
  tableToSlides(eleId, options = {}) {
    genTableToSlides(
      this,
      eleId,
      options,
      options?.masterSlideName ? this.slideLayouts.filter((layout) => layout._name === options.masterSlideName)[0] : null
    );
  }
};

// src/core/color.ts
function parseColor(input) {
  if (!input) return null;
  const r = input.trim().toLowerCase();
  if (r === "transparent" || r === "none") return null;
  if (r[0] === "#") {
    let body = r.slice(1);
    if (body.length === 3 || body.length === 4) {
      body = body.split("").map((c) => c + c).join("");
    }
    let alpha = 1;
    if (body.length === 8) {
      alpha = parseInt(body.slice(6, 8), 16) / 255;
      body = body.slice(0, 6);
    }
    if (alpha === 0 || body.length !== 6 || /[^0-9a-f]/.test(body)) return null;
    return { hex: body.toUpperCase(), alpha };
  }
  const m = r.match(
    /rgba?\(\s*(-?[\d.]+)[\s,]+(-?[\d.]+)[\s,]+(-?[\d.]+)(?:[\s,/]+([\d.]+%?))?\s*\)/
  );
  if (m) {
    const red = clamp(Math.round(parseFloat(m[1])), 0, 255);
    const green = clamp(Math.round(parseFloat(m[2])), 0, 255);
    const blue = clamp(Math.round(parseFloat(m[3])), 0, 255);
    let alpha = 1;
    if (m[4] !== void 0) {
      alpha = m[4].endsWith("%") ? parseFloat(m[4]) / 100 : parseFloat(m[4]);
      alpha = clamp(alpha, 0, 1);
    }
    if (alpha === 0) return null;
    return {
      hex: (red << 16 | green << 8 | blue).toString(16).padStart(6, "0").toUpperCase(),
      alpha
    };
  }
  return null;
}
function parseGradient(input) {
  if (!input || !input.includes("gradient(")) return null;
  const m = input.match(/(rgba?\([^)]+\)|#[0-9a-fA-F]{3,8})/);
  return m ? parseColor(m[0]) : null;
}
function opacityToTransparency(alpha) {
  return clamp(Math.round((1 - alpha) * 100), 0, 100);
}

// src/core/css.ts
function extractPx(value) {
  if (!value) return 0;
  const m = value.match(/(-?[\d.]+)px/);
  return m ? parseFloat(m[1]) : 0;
}
function isBold(weight) {
  if (!weight) return false;
  if (weight === "bold" || weight === "bolder") return true;
  const n = parseInt(weight, 10);
  return !isNaN(n) && n >= 600;
}
function textAlign(value) {
  switch (value) {
    case "center":
      return "center";
    case "right":
    case "end":
      return "right";
    case "justify":
      return "justify";
    default:
      return "left";
  }
}
function borderStyleToDashType(value) {
  switch (value) {
    case "dashed":
      return "dash";
    case "dotted":
      return "dot";
    case "double":
      return "solid";
    default:
      return void 0;
  }
}
function parseBorderRadius(value, minSide) {
  if (!value || minSide <= 0) return 0;
  const pct = value.match(/^([\d.]+)%/);
  const px = pct ? parseFloat(pct[1]) / 100 * minSide : extractPx(value);
  return px <= 0 ? 0 : clamp(px, 0, minSide / 2);
}
function extractRotation(transform) {
  if (!transform || transform === "none") return void 0;
  const rot = transform.match(/rotate\((-?[\d.]+)deg\)/);
  if (rot) {
    const deg = parseFloat(rot[1]);
    return deg === 0 ? void 0 : deg;
  }
  const mat = transform.match(/matrix\(\s*([^,\s]+),\s*([^,\s]+),/);
  if (mat) {
    const a = parseFloat(mat[1]);
    const b = parseFloat(mat[2]);
    const deg = Math.atan2(b, a) * (180 / Math.PI);
    return Math.abs(deg) < 0.01 ? void 0 : deg;
  }
  return void 0;
}
function parseShadow(value) {
  if (!value || value === "none") return null;
  const parts = [];
  let depth = 0;
  let buf = "";
  for (const ch of value) {
    if (ch === "(") depth++;
    else if (ch === ")") depth--;
    if (ch === "," && depth === 0) {
      parts.push(buf);
      buf = "";
    } else {
      buf += ch;
    }
  }
  if (buf) parts.push(buf);
  for (const part of parts) {
    const s = part.trim();
    if (/\binset\b/.test(s)) continue;
    const colorMatch = s.match(/(rgba?\([^)]+\)|#[0-9a-fA-F]{3,8})/);
    const color = colorMatch ? parseColor(colorMatch[0]) : { hex: "000000", alpha: 0.3 };
    if (!color) continue;
    const lengths = s.replace(colorMatch?.[0] ?? "", "").match(/-?[\d.]+px/g);
    if (!lengths || lengths.length < 2) continue;
    const offsetX = parseFloat(lengths[0]);
    const offsetY = parseFloat(lengths[1]);
    const blur = lengths[2] ? parseFloat(lengths[2]) : 0;
    const dist = Math.sqrt(offsetX * offsetX + offsetY * offsetY);
    let angle = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
    if (angle < 0) angle += 360;
    return {
      type: "outer",
      color: color.hex,
      opacity: clamp(color.alpha, 0, 1),
      blur: clamp(pxToPoints(blur), 0, 100),
      offset: clamp(pxToPoints(dist), 0, 200),
      angle: Math.round(angle)
    };
  }
  return null;
}
function lineSpacingMultiple(lineHeight, fontSizePx) {
  if (!lineHeight || lineHeight === "normal") return void 0;
  const px = extractPx(lineHeight);
  if (px > 0) {
    const baseline = fontSizePx * 1.3333333333333333;
    return baseline <= 0 ? void 0 : clamp(px / baseline, 0.5, 5);
  }
  const mult = parseFloat(lineHeight);
  if (!isNaN(mult) && mult > 0 && mult < 10) return clamp(mult, 0.5, 5);
  return void 0;
}
function letterSpacingPoints(value) {
  if (!value || value === "normal") return void 0;
  const px = extractPx(value);
  if (px !== 0) return clamp(pxToPoints(px), -20, 100);
  return void 0;
}
function underlineStyle(value) {
  switch (value) {
    case "double":
      return "dbl";
    case "dotted":
      return "dotted";
    case "dashed":
      return "dash";
    case "wavy":
      return "wavy";
    default:
      return "sng";
  }
}
function isPreserveWhitespace(whiteSpace) {
  return whiteSpace === "pre" || whiteSpace === "pre-wrap" || whiteSpace === "break-spaces";
}
function trimBlockNewlines(text) {
  return text.replace(/^\n+|\n+$/g, "");
}
function normalizeText(text, whiteSpace) {
  if (isPreserveWhitespace(whiteSpace)) {
    return text;
  }
  if (whiteSpace === "pre-line") {
    return text.split("\n").map((line) => line.replace(/[ \t]+/g, " ").trim()).join("\n");
  }
  return text.replace(/\s+/g, " ").trim();
}
function noWrap(style) {
  if (style.whiteSpace === "pre") return true;
  if (style.whiteSpace !== "nowrap") return false;
  const overflow = (style.overflow ?? "").trim().split(/\s+/)[0] ?? "";
  const scrollable = overflow === "auto" || overflow === "scroll" || overflow === "overlay";
  const ellipsis = (overflow === "hidden" || overflow === "clip") && (style.textOverflow ?? "").includes("ellipsis");
  return !scrollable && !ellipsis;
}

// src/render/context.ts
function mintAnimName(ctx, entry) {
  return `anim-s${ctx.slideNo}-n${entry.def.index}-k${entry.minted++}`;
}
function recordAnimShape(ctx, entry, name) {
  const objs = ctx.slide._slideObjects;
  if (!Array.isArray(objs) || objs.length === 0) return;
  entry.shapeNames.push(name);
  entry.spids.push(objs.length + 1);
}
function rectToPptx(rect, ctx) {
  const x = rect.x - ctx.originX;
  const y = rect.y - ctx.originY;
  return {
    x: pxToInches(clamp(x, -ctx.slideW, ctx.slideW * 2)),
    y: pxToInches(clamp(y, -ctx.slideH, ctx.slideH * 2)),
    w: pxToInches(Math.max(clamp(rect.w, 0, ctx.slideW * 2), 1)),
    h: pxToInches(Math.max(clamp(rect.h, 0, ctx.slideH * 2), 1))
  };
}

// src/render/media-cache.ts
function hashSvg(s) {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619) >>> 0;
  }
  return h.toString(36);
}
function imageKey(node) {
  if (node.imageUrl) return node.imageUrl;
  if (node.svg) return `svg:${hashSvg(node.svg)}:${node.rect.w}x${node.rect.h}`;
  if (node.gradient) {
    return `grad:${hashSvg(node.gradient)}:${Math.round(node.rect.w)}x${Math.round(
      node.rect.h
    )}:r${Math.round(node.gradientRadius ?? 0)}`;
  }
  return null;
}
function collectImageRefs(node, urls, svgs, grads) {
  if (node.imageUrl) {
    urls.add(node.imageUrl);
  } else if (node.svg) {
    const key = imageKey(node);
    if (!svgs.has(key)) svgs.set(key, { svg: node.svg, w: node.rect.w, h: node.rect.h });
  } else if (node.gradient) {
    const key = imageKey(node);
    if (!grads.has(key)) {
      grads.set(key, {
        css: node.gradient,
        w: node.rect.w,
        h: node.rect.h,
        radius: node.gradientRadius ?? 0
      });
    }
  }
  for (const child of node.children) collectImageRefs(child, urls, svgs, grads);
}
async function buildMediaCache(slides, warnings, resolveMedia) {
  const urls = /* @__PURE__ */ new Set();
  const svgs = /* @__PURE__ */ new Map();
  const grads = /* @__PURE__ */ new Map();
  for (const slide of slides) collectImageRefs(slide.root, urls, svgs, grads);
  const refs = [];
  for (const url of urls) refs.push({ kind: "url", key: url, url });
  for (const [key, { svg, w, h }] of svgs) refs.push({ kind: "svg", key, svg, w, h });
  for (const [key, { css, w, h, radius }] of grads)
    refs.push({ kind: "gradient", key, css, w, h, radius });
  const cache = /* @__PURE__ */ new Map();
  if (refs.length === 0) return cache;
  const resolved = await resolveMedia(refs);
  for (const r of resolved) {
    for (const w of r.warnings) warnings.push(w);
    cache.set(r.key, r.value);
  }
  return cache;
}

// src/core/fonts.ts
var GENERIC_FONT_MAP = {
  serif: "Georgia",
  "sans-serif": "Arial",
  monospace: "Courier New",
  "system-ui": "Arial",
  "-apple-system": "Arial",
  blinkmacsystemfont: "Arial",
  "ui-serif": "Georgia",
  "ui-sans-serif": "Arial",
  "ui-monospace": "Courier New",
  "ui-rounded": "Arial",
  cursive: "Comic Sans MS",
  fantasy: "Impact",
  math: "Cambria Math",
  emoji: "Segoe UI Emoji"
};
function resolveFontFamily(family, swapMap) {
  if (!family) return "Arial";
  const first = family.split(",")[0].trim().replace(/^['"]|['"]$/g, "");
  if (!first) return "Arial";
  const lc = first.toLowerCase();
  if (swapMap) {
    for (const key in swapMap) if (key.toLowerCase() === lc) return swapMap[key];
  }
  return GENERIC_FONT_MAP[lc] ?? first;
}

// src/render/text-runs.ts
function textFormat(style, swapMap) {
  const color = parseColor(style.color) || { hex: "000000", alpha: 1 };
  const deco = `${style.textDecorationLine || ""} ${style.textDecoration || ""}`;
  const decoColor = parseColor(style.textDecorationColor);
  const underlineColor = decoColor && decoColor.hex !== color.hex ? decoColor.hex : void 0;
  const fontSizePx = extractPx(style.fontSize);
  return {
    color: color.hex,
    transparency: opacityToTransparency(color.alpha),
    bold: isBold(style.fontWeight),
    italic: style.fontStyle === "italic" || (style.fontStyle || "").startsWith("oblique"),
    underline: deco.includes("underline") ? { style: underlineStyle(style.textDecorationStyle), color: underlineColor } : void 0,
    strike: deco.includes("line-through") ? true : void 0,
    fontFace: resolveFontFamily(style.fontFamily, swapMap),
    fontSize: fontSizePx > 0 ? clamp(pxToPoints(fontSizePx), 1, 400) : void 0
  };
}
function textTransformFn(value) {
  switch (value) {
    case "uppercase":
      return (s) => s.toUpperCase();
    case "lowercase":
      return (s) => s.toLowerCase();
    case "capitalize":
      return (s) => s.replace(
        new RegExp("(^|[^\\p{L}\\p{N}'\\u2019])(\\p{L})", "gu"),
        (_m, pre, ch) => pre + ch.toUpperCase()
      );
    default:
      return (s) => s;
  }
}
var INLINE_MERGE_KEYS = ["textTransform", "letterSpacing"];
function subtreeHasAnim(node) {
  return node.anim !== void 0 || node.children.some(subtreeHasAnim);
}
function isRunMergeable(node, parentStyle) {
  if (node.tag === "#text") return true;
  const s = node.style;
  const display = s.display || "";
  if (display !== "inline" && display !== "inline-block" && display !== "inline-flex" && display !== "contents" || s.visibility === "hidden" || s.opacity && s.opacity !== "1") {
    return false;
  }
  const valign = s.verticalAlign || "baseline";
  if (valign !== "baseline" && valign !== "sub" && valign !== "super" && valign !== "0px" && valign !== "0" || imageKey(node) || parseColor(s.backgroundColor) && extractPx(s.paddingTop) + extractPx(s.paddingRight) + extractPx(s.paddingBottom) + extractPx(s.paddingLeft) > 0 || extractPx(s.borderTopWidth || s.borderWidth) > 0 || extractPx(s.borderBottomWidth) > 0 || extractPx(s.borderLeftWidth) > 0 || extractPx(s.borderRightWidth) > 0 || parseShadow(s.boxShadow) || parseShadow(s.textShadow)) {
    return false;
  }
  for (const k of INLINE_MERGE_KEYS) if ((s[k] || "") !== (parentStyle[k] || "")) return false;
  return true;
}
function valignFromBox(style) {
  if (style.display === "flex" || style.display === "inline-flex") {
    const v = (style.flexDirection || "").startsWith("column") ? style.justifyContent : style.alignItems;
    if (v === "center") return "middle";
    if (v === "flex-end" || v === "end") return "bottom";
  }
  if (style.display === "table-cell") {
    if (style.verticalAlign === "middle") return "middle";
    if (style.verticalAlign === "bottom") return "bottom";
  }
  return "top";
}
function alignFromFlex(style) {
  if (style.display !== "flex" && style.display !== "inline-flex") return void 0;
  const a = (style.flexDirection || "").startsWith("column") ? style.alignItems : style.justifyContent;
  if (a === "center") return "center";
  if (a === "flex-end" || a === "end") return "right";
  return void 0;
}
function httpHref(href) {
  return href && /^https?:\/\//i.test(href) ? href : void 0;
}
function formatsEqual(a, b) {
  return a.color === b.color && a.transparency === b.transparency && a.bold === b.bold && a.italic === b.italic && a.fontFace === b.fontFace && a.fontSize === b.fontSize && !!a.underline === !!b.underline && a.underline?.style === b.underline?.style && a.underline?.color === b.underline?.color && !!a.strike === !!b.strike && !!a.subscript === !!b.subscript && !!a.superscript === !!b.superscript && a.highlight === b.highlight;
}
function runsAdjacent(a, b) {
  if (!a.rect || !b.rect) return false;
  const gap = b.rect.x - (a.rect.x + a.rect.w);
  return gap >= -1 && gap < 2;
}
function extractTextRuns(node, swapMap) {
  const consumed = /* @__PURE__ */ new Set();
  const collected = [];
  const ownText = node.text ? normalizeText(node.text, node.style.whiteSpace) : "";
  if (ownText) collected.push({ text: ownText, fmt: textFormat(node.style, swapMap), rect: node.rect });
  const allMergeable = (n) => !subtreeHasAnim(n) && isRunMergeable(n, node.style) ? n.children.every(allMergeable) : false;
  const absorb = (n, inherited) => {
    consumed.add(n);
    if (n.tag !== "#text" && (n.style.visibility === "hidden" || n.style.opacity === "0")) return;
    if (n.tag === "br") {
      collected.push({ text: "\n", fmt: textFormat(n.style, swapMap), rect: n.rect });
      return;
    }
    const href = httpHref(n.href) || inherited?.href;
    const fmt = textFormat(n.style, swapMap);
    const valign = n.style.verticalAlign;
    const baseline = valign === "sub" ? "sub" : valign === "super" ? "super" : inherited?.baseline;
    if (baseline === "sub") fmt.subscript = true;
    else if (baseline === "super") fmt.superscript = true;
    if (!fmt.underline && inherited?.underline) fmt.underline = inherited.underline;
    if (!fmt.strike && inherited?.strike) fmt.strike = true;
    const bg = parseColor(n.style.backgroundColor);
    const highlight = bg ? bg.hex : inherited?.highlight;
    if (highlight) fmt.highlight = highlight;
    const text = n.text ? normalizeText(n.text, n.style.whiteSpace) : "";
    if (text) collected.push({ text, fmt, rect: n.rect, href });
    for (const child of n.children) {
      absorb(child, { href, baseline, underline: fmt.underline, strike: fmt.strike, highlight });
    }
  };
  if (node.children.length > 0) {
    const hasDirectText = node.children.some((c) => c.tag === "#text");
    const hasMedia = (n) => !!imageKey(n) || n.children.some(hasMedia);
    const allInlineNoMedia = node.children.every((c) => {
      if (c.tag === "#text") return true;
      if (hasMedia(c) || subtreeHasAnim(c)) return false;
      const display = c.style.display || "";
      return display === "inline" || display === "inline-block" || display === "inline-flex" || display === "contents";
    });
    if (node.children.every(allMergeable) || hasDirectText && allInlineNoMedia) {
      for (const child of node.children) absorb(child);
    }
  }
  const preserveWs = isPreserveWhitespace(node.style.whiteSpace);
  const merged = [];
  for (const run of collected) {
    const last = merged.at(-1);
    if (last && last.href === run.href && formatsEqual(last.fmt, run.fmt)) {
      const sep = preserveWs || runsAdjacent(last, run) || last.text.endsWith("\n") || run.text.startsWith("\n") ? "" : " ";
      last.text += sep + run.text;
      last.rect = run.rect;
    } else {
      merged.push({ ...run });
    }
  }
  if (merged.length > 0) {
    merged[0].text = merged[0].text.replace(/^\n+/, "");
    merged[merged.length - 1].text = merged[merged.length - 1].text.replace(/\n+$/, "");
  }
  return { runs: merged.filter((r) => r.text !== ""), consumed };
}

// src/render/list.ts
function listStyleConfig(value) {
  switch (value) {
    case "decimal":
    case "decimal-leading-zero":
      return { isNum: true };
    case "lower-alpha":
    case "lower-latin":
      return { isNum: true, numberType: "alphaLcPeriod" };
    case "upper-alpha":
    case "upper-latin":
      return { isNum: true, numberType: "alphaUcPeriod" };
    case "lower-roman":
      return { isNum: true, numberType: "romanLcPeriod" };
    case "upper-roman":
      return { isNum: true, numberType: "romanUcPeriod" };
    case "circle":
      return { isNum: false, characterCode: "25CB" };
    case "square":
      return { isNum: false, characterCode: "25A0" };
    default:
      return { isNum: false };
  }
}
var LIST_UNIFORM_KEYS = [
  "color",
  "fontWeight",
  "fontStyle",
  "fontFamily",
  "fontSize",
  "textDecorationLine",
  "textDecoration",
  "textDecorationStyle",
  "textDecorationColor",
  "textTransform",
  "letterSpacing",
  "textShadow"
];
function detectList(node) {
  if (node.tag !== "ul" && node.tag !== "ol") return null;
  const items = [];
  const first = node.children[0];
  for (const li of node.children) {
    if (li.anim) return null;
    if (li.tag !== "li" || li.style.visibility === "hidden" || li.style.opacity === "0" || li.children.length > 0 || imageKey(li) || parseColor(li.style.backgroundColor)) {
      return null;
    }
    for (const k of LIST_UNIFORM_KEYS) {
      if ((li.style[k] || "") !== (first.style[k] || "")) return null;
    }
    const text = li.text ? normalizeText(li.text, li.style.whiteSpace) : "";
    if (!text) return null;
    items.push(text);
  }
  return items.length > 0 ? items : null;
}

// src/render/render-node.ts
var PT_PER_INCH = 72;
function renderNodeToPptx(node, ctx, inheritedOpacity = 1, activeAnim) {
  let anim = activeAnim;
  if (node.anim) {
    anim = { def: node.anim, spids: [], shapeNames: [], minted: 0 };
    ctx.animEntries.push(anim);
  }
  if (node.tag === "#text") {
    const style2 = node.style;
    const text = node.text ? trimBlockNewlines(normalizeText(node.text, style2.whiteSpace)) : "";
    if (!text || node.rect.w < 0.5 || node.rect.h < 0.5 || style2.visibility === "hidden") return;
    const coords2 = rectToPptx(node.rect, ctx);
    const fmt = textFormat(style2, ctx.fontMap);
    const fontSize = fmt.fontSize ?? clamp(pxToPoints(extractPx(style2.fontSize) || 16), 1, 400);
    const opacity2 = inheritedOpacity * clamp(parseFloat(style2.opacity ?? "1") || 1, 0, 1);
    const widthPad = pxToInches(Math.max(8, node.rect.w * 0.03));
    const opts = {
      x: coords2.x,
      y: coords2.y,
      w: coords2.w + widthPad,
      h: coords2.h + pxToInches(4),
      margin: 0,
      fontFace: fmt.fontFace,
      fontSize,
      bold: fmt.bold,
      italic: fmt.italic,
      underline: fmt.underline,
      strike: fmt.strike,
      color: fmt.color,
      transparency: opacityToTransparency((1 - fmt.transparency / 100) * opacity2),
      align: textAlign(style2.textAlign),
      valign: "top",
      fit: "shrink",
      wrap: !noWrap(style2),
      lineSpacingMultiple: lineSpacingMultiple(style2.lineHeight, fontSize),
      charSpacing: letterSpacingPoints(style2.letterSpacing)
    };
    if (anim) opts.objectName = mintAnimName(ctx, anim);
    try {
      ctx.slide.addText(textTransformFn(style2.textTransform)(text), opts);
      if (anim) recordAnimShape(ctx, anim, opts.objectName);
    } catch (err) {
      ctx.warnings.push(`addText(#text) failed: ${errMsg(err)}`);
    }
    return;
  }
  const style = node.style;
  const opacity = inheritedOpacity * clamp(parseFloat(style.opacity ?? "1") || 1, 0, 1);
  if (node.rect.w < 0.5 || node.rect.h < 0.5) {
    for (const child of node.children) renderNodeToPptx(child, ctx, opacity, anim);
    return;
  }
  if (style.display === "none" || style.opacity === "0") return;
  if (style.visibility === "hidden") {
    for (const child of node.children) renderNodeToPptx(child, ctx, opacity, anim);
    return;
  }
  const rotation = extractRotation(style.transform);
  const box = rotation !== void 0 && node.untransformedRect ? node.untransformedRect : node.rect;
  const coords = rectToPptx(box, ctx);
  const minSide = Math.min(box.w, box.h);
  const radiusPx = parseBorderRadius(style.borderTopLeftRadius || style.borderRadius, minSide);
  const radiusRatio = minSide > 0 ? radiusPx / minSide : 0;
  const hasRasterGradient = !!node.gradient && !!ctx.mediaCache.get(imageKey(node) ?? "");
  const bgColor = parseColor(style.backgroundColor) || (node.imageUrl || hasRasterGradient ? null : parseGradient(style.backgroundImage));
  const bgFill = bgColor ? { hex: bgColor.hex, alpha: bgColor.alpha * opacity } : null;
  const topW = extractPx(style.borderTopWidth || style.borderWidth);
  const bottomW = extractPx(style.borderBottomWidth);
  const leftW = extractPx(style.borderLeftWidth);
  const rightW = extractPx(style.borderRightWidth);
  const anyBorder = topW || bottomW || leftW || rightW;
  const topColor = style.borderTopColor || style.borderColor;
  const bottomColor = style.borderBottomColor || style.borderColor;
  const leftColor = style.borderLeftColor || style.borderColor;
  const rightColor = style.borderRightColor || style.borderColor;
  const topStyle = style.borderTopStyle || style.borderStyle;
  const bottomStyle = style.borderBottomStyle || style.borderStyle;
  const leftStyle = style.borderLeftStyle || style.borderStyle;
  const rightStyle = style.borderRightStyle || style.borderStyle;
  const uniformBorder = topW > 0 && topW === bottomW && topW === leftW && topW === rightW && topColor === bottomColor && topColor === leftColor && topColor === rightColor && topStyle === bottomStyle && topStyle === leftStyle && topStyle === rightStyle;
  const uniformBorderColor = uniformBorder ? parseColor(topColor) : null;
  const shadow = parseShadow(style.boxShadow);
  const dashType = borderStyleToDashType(style.borderTopStyle || style.borderStyle);
  if (bgFill || uniformBorderColor || shadow) {
    const opts = {
      x: coords.x,
      y: coords.y,
      w: coords.w,
      h: coords.h,
      fill: bgFill ? { color: bgFill.hex, transparency: opacityToTransparency(bgFill.alpha) } : { type: "none" },
      line: uniformBorderColor ? {
        color: uniformBorderColor.hex,
        width: clamp(pxToPoints(topW), 0.25, 20),
        transparency: opacityToTransparency(uniformBorderColor.alpha * opacity),
        dashType
      } : { type: "none" }
    };
    if (shadow) opts.shadow = { ...shadow, opacity: shadow.opacity * opacity };
    if (rotation !== void 0) opts.rotate = rotation;
    let shapeName = "rect";
    if (radiusRatio >= 0.49) {
      const aspect = box.w / box.h;
      if (aspect > 0.75 && aspect < 1.33) {
        shapeName = "ellipse";
      } else {
        shapeName = "roundRect";
        opts.rectRadius = pxToInches(radiusPx);
      }
    } else if (radiusPx > 0.5) {
      shapeName = "roundRect";
      opts.rectRadius = pxToInches(radiusPx);
    }
    if (anim) opts.objectName = mintAnimName(ctx, anim);
    try {
      ctx.slide.addShape(shapeName, opts);
      if (anim) recordAnimShape(ctx, anim, opts.objectName);
    } catch (err) {
      ctx.warnings.push(`addShape failed for <${node.tag}>: ${errMsg(err)}`);
    }
  }
  if (!uniformBorder && anyBorder > 0) {
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    const rad = rotation !== void 0 ? rotation * Math.PI / 180 : 0;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const side = (width, colorStr, px, py, pw, ph) => {
      if (width <= 0) return;
      const color = parseColor(colorStr);
      if (!color) return;
      let ox = px;
      let oy = py;
      if (rotation !== void 0) {
        const dx = px + pw / 2 - cx;
        const dy = py + ph / 2 - cy;
        ox = cx + dx * cos - dy * sin - pw / 2;
        oy = cy + dx * sin + dy * cos - ph / 2;
      }
      try {
        const opts = {
          x: pxToInches(ox - ctx.originX),
          y: pxToInches(oy - ctx.originY),
          w: pxToInches(Math.max(pw, 1)),
          h: pxToInches(Math.max(ph, 1)),
          fill: { color: color.hex, transparency: opacityToTransparency(color.alpha * opacity) },
          line: { type: "none" }
        };
        if (rotation !== void 0) opts.rotate = rotation;
        if (anim) opts.objectName = mintAnimName(ctx, anim);
        ctx.slide.addShape("rect", opts);
        if (anim) recordAnimShape(ctx, anim, opts.objectName);
      } catch {
      }
    };
    if (bottomW > 0) side(bottomW, bottomColor, box.x, box.y + box.h - bottomW, box.w, bottomW);
    if (topW > 0) side(topW, topColor, box.x, box.y, box.w, topW);
    if (leftW > 0) side(leftW, leftColor, box.x, box.y, leftW, box.h);
    if (rightW > 0) side(rightW, rightColor, box.x + box.w - rightW, box.y, rightW, box.h);
  }
  const key = imageKey(node);
  if (key) {
    const entry = ctx.mediaCache.get(key);
    if (entry) {
      const opts = { x: coords.x, y: coords.y, w: coords.w, h: coords.h, data: entry.dataUrl };
      const isMediaTag = node.tag === "img" || node.tag === "canvas" || node.tag === "svg" || node.tag === "object";
      const hasBgImage = !!style.backgroundImage && style.backgroundImage !== "none";
      const fit = (style.objectFit && (isMediaTag || style.objectFit !== "fill" && !hasBgImage) ? style.objectFit : void 0) || (style.backgroundSize === "cover" ? "cover" : style.backgroundSize === "contain" ? "contain" : void 0);
      if (fit === "cover" || fit === "contain") {
        opts.sizing = { type: fit, w: coords.w, h: coords.h };
        const centered = (pos) => !pos || /^50%\s+50%$/.test(pos.trim());
        const objFitCentered = !!style.objectFit && (isMediaTag || style.objectFit !== "fill") && (style.objectFit === "cover" || style.objectFit === "contain") && (isMediaTag ? centered(style.objectPosition) : !hasBgImage);
        const bgCentered = !isMediaTag && hasBgImage && /^50%\s+50%$/.test((style.backgroundPosition ?? "").trim());
        if (node.imageUrl && entry.w && entry.h && (objFitCentered || bgCentered)) {
          opts.w = pxToInches(entry.w);
          opts.h = pxToInches(entry.h);
        }
      }
      if (radiusRatio >= 0.4 && !node.gradient) opts.rounding = true;
      if (opacity < 1) opts.transparency = clamp(Math.round((1 - opacity) * 100), 0, 100);
      if (rotation !== void 0) opts.rotate = rotation;
      if (anim) opts.objectName = mintAnimName(ctx, anim);
      try {
        ctx.slide.addImage(opts);
        if (anim) recordAnimShape(ctx, anim, opts.objectName);
      } catch (err) {
        ctx.warnings.push(`addImage failed for <${node.tag}>: ${errMsg(err)}`);
      }
    }
    if (node.tag === "svg" || node.tag === "img" || node.tag === "canvas") return;
  }
  const listItems = detectList(node);
  const consumed = /* @__PURE__ */ new Set();
  if (listItems) {
    const firstLi = node.children[0];
    const liStyle = firstLi.style;
    const listType = liStyle.listStyleType || "";
    const fontSize = clamp(pxToPoints(extractPx(liStyle.fontSize) || 16), 1, 400);
    const fmt = textFormat(liStyle, ctx.fontMap);
    const transform = textTransformFn(liStyle.textTransform);
    const cfg = listStyleConfig(listType);
    const indentPt = cfg.isNum ? Math.max(fontSize * 1.5, 14) : Math.max(fontSize * 0.7, 8);
    const lastLi = node.children[node.children.length - 1];
    const top = firstLi.rect.y - ctx.originY;
    const bottom = lastLi.rect.y + lastLi.rect.h - ctx.originY;
    const left = firstLi.rect.x - ctx.originX;
    const opts = {
      x: pxToInches(left) - indentPt / PT_PER_INCH,
      y: pxToInches(top),
      w: pxToInches(Math.max(firstLi.rect.w, 4)) + indentPt / PT_PER_INCH + pxToInches(8),
      h: pxToInches(Math.max(bottom - top, 4)) + pxToInches(4),
      margin: 2,
      fontFace: fmt.fontFace,
      fontSize,
      bold: fmt.bold,
      italic: fmt.italic,
      underline: fmt.underline,
      strike: fmt.strike,
      color: fmt.color,
      transparency: clamp(Math.round(100 - (100 - fmt.transparency) * opacity), 0, 100),
      align: textAlign(liStyle.textAlign),
      valign: "top",
      fit: "shrink",
      wrap: !noWrap(liStyle),
      lineSpacingMultiple: lineSpacingMultiple(liStyle.lineHeight, fontSize),
      charSpacing: letterSpacingPoints(liStyle.letterSpacing)
    };
    const listShadow = parseShadow(liStyle.textShadow);
    if (listShadow) opts.shadow = { ...listShadow, opacity: listShadow.opacity * opacity };
    const tabPos = indentPt / PT_PER_INCH;
    const bulletFor = (i) => listType === "none" ? false : cfg.isNum ? { type: "number", indent: indentPt, numberStartAt: i + 1, numberType: cfg.numberType } : { indent: indentPt, characterCode: cfg.characterCode };
    const textObjs = listItems.map((item, i) => ({
      text: transform(item),
      options: {
        bullet: bulletFor(i),
        tabStops: listType === "none" ? void 0 : [{ position: tabPos }],
        breakLine: i < listItems.length - 1
      }
    }));
    if (anim) opts.objectName = mintAnimName(ctx, anim);
    try {
      ctx.slide.addText(textObjs, opts);
      if (anim) recordAnimShape(ctx, anim, opts.objectName);
    } catch (err) {
      ctx.warnings.push(`addText(list) failed: ${errMsg(err)}`);
    }
    for (const child of node.children) consumed.add(child);
  }
  const { runs, consumed: runConsumed } = extractTextRuns(node, ctx.fontMap);
  for (const c of runConsumed) consumed.add(c);
  if (runs.length > 0) {
    const fontSize = clamp(pxToPoints(extractPx(style.fontSize) || 16), 1, 400);
    const valign = valignFromBox(style);
    const align = alignFromFlex(style) || textAlign(style.textAlign);
    const padTop = extractPx(style.paddingTop);
    const padBottom = extractPx(style.paddingBottom);
    const padLeft = extractPx(style.paddingLeft);
    const padRight = extractPx(style.paddingRight);
    const insetLeft = leftW + padLeft;
    const insetRight = rightW + padRight;
    const insetTop = topW + padTop;
    const insetBottom = bottomW + padBottom;
    const boxX = coords.x + pxToInches(insetLeft);
    const boxY = coords.y + pxToInches(insetTop);
    const boxW = Math.max(coords.w - pxToInches(insetLeft + insetRight), pxToInches(4));
    const boxH = Math.max(coords.h - pxToInches(insetTop + insetBottom), pxToInches(4));
    const extraW = pxToInches(Math.max(8, box.w * 0.03));
    const shiftX = align === "right" ? extraW : align === "center" ? extraW / 2 : 0;
    const opts = {
      x: boxX - shiftX,
      y: boxY,
      w: boxW + extraW,
      h: boxH + pxToInches(4),
      margin: 2,
      fontSize,
      align,
      valign,
      fit: "shrink",
      wrap: !noWrap(style),
      lineSpacingMultiple: lineSpacingMultiple(style.lineHeight, fontSize),
      charSpacing: letterSpacingPoints(style.letterSpacing)
    };
    if (rotation !== void 0) opts.rotate = rotation;
    const textShadow = parseShadow(style.textShadow);
    if (textShadow) opts.shadow = { ...textShadow, opacity: textShadow.opacity * opacity };
    if (node.tag === "li") {
      const listType = style.listStyleType || "";
      if (listType !== "none") {
        const cfg = listStyleConfig(listType);
        const indentPt = cfg.isNum ? Math.max(fontSize * 1.5, 14) : Math.max(fontSize * 0.7, 8);
        const indentIn = indentPt / PT_PER_INCH;
        opts.x = opts.x - indentIn;
        opts.w = opts.w + indentIn;
        opts.tabStops = [{ position: indentIn }];
        opts.bullet = cfg.isNum ? {
          type: "number",
          indent: indentPt,
          numberType: cfg.numberType,
          ...node.liIndex ? { numberStartAt: node.liIndex } : {}
        } : { indent: indentPt, characterCode: cfg.characterCode };
      }
    }
    const nodeHref = httpHref(node.href);
    if (nodeHref) opts.hyperlink = { url: nodeHref };
    const transform = textTransformFn(style.textTransform);
    if (runs.length === 1) {
      const run = runs[0];
      const alpha = (1 - run.fmt.transparency / 100) * opacity;
      if (!nodeHref && run.href) opts.hyperlink = { url: run.href };
      Object.assign(opts, {
        fontFace: run.fmt.fontFace,
        fontSize: run.fmt.fontSize ?? opts.fontSize,
        bold: run.fmt.bold,
        italic: run.fmt.italic,
        underline: run.fmt.underline,
        strike: run.fmt.strike,
        subscript: run.fmt.subscript,
        superscript: run.fmt.superscript,
        highlight: run.fmt.highlight,
        color: run.fmt.color,
        transparency: opacityToTransparency(alpha)
      });
      if (anim) opts.objectName = mintAnimName(ctx, anim);
      try {
        ctx.slide.addText(transform(run.text), opts);
        if (anim) recordAnimShape(ctx, anim, opts.objectName);
      } catch (err) {
        ctx.warnings.push(`addText failed for <${node.tag}>: ${errMsg(err)}`);
      }
    } else {
      const preserveWs = isPreserveWhitespace(style.whiteSpace);
      const runOpts = (run) => {
        const alpha = (1 - run.fmt.transparency / 100) * opacity;
        return {
          fontFace: run.fmt.fontFace,
          fontSize: run.fmt.fontSize,
          bold: run.fmt.bold,
          italic: run.fmt.italic,
          underline: run.fmt.underline,
          strike: run.fmt.strike,
          subscript: run.fmt.subscript,
          superscript: run.fmt.superscript,
          highlight: run.fmt.highlight,
          color: run.fmt.color,
          transparency: opacityToTransparency(alpha),
          hyperlink: !nodeHref && run.href ? { url: run.href } : void 0
        };
      };
      const lines = [[]];
      runs.forEach((run, i) => {
        const segs = transform(run.text).split("\n");
        segs.forEach((seg, j) => {
          if (j > 0) lines.push([]);
          if (seg !== "") lines[lines.length - 1].push({ text: seg, options: runOpts(run) });
        });
        const next = runs[i + 1];
        const sep = !preserveWs && next && !runsAdjacent(run, next) && !/\n$/.test(run.text) && !/^\n/.test(next.text) ? " " : "";
        if (sep) {
          const cur = lines[lines.length - 1];
          if (cur.length) cur[cur.length - 1].text += sep;
        }
      });
      const lastLine = lines.length - 1;
      const textObjs = [];
      lines.forEach((line, li) => {
        const breakAfter = li < lastLine;
        if (line.length === 0) {
          textObjs.push({ text: "", options: { breakLine: breakAfter } });
          return;
        }
        line.forEach((piece, pi) => {
          textObjs.push({
            text: piece.text,
            options: { ...piece.options, breakLine: pi === line.length - 1 ? breakAfter : false }
          });
        });
      });
      if (anim) opts.objectName = mintAnimName(ctx, anim);
      try {
        ctx.slide.addText(textObjs, opts);
        if (anim) recordAnimShape(ctx, anim, opts.objectName);
      } catch (err) {
        ctx.warnings.push(`addText(runs) failed for <${node.tag}>: ${errMsg(err)}`);
      }
    }
  }
  for (const child of orderByZIndex(node.children))
    if (!consumed.has(child)) renderNodeToPptx(child, ctx, opacity, anim);
}
function orderByZIndex(children) {
  const zOf = (n) => {
    const pos = n.style.position;
    if (pos !== "relative" && pos !== "absolute" && pos !== "fixed" && pos !== "sticky") return 0;
    const v = parseInt(n.style.zIndex ?? "", 10);
    return Number.isNaN(v) ? 0 : v;
  };
  if (children.every((c) => zOf(c) === 0)) return children;
  return children.map((c, i) => ({ c, z: zOf(c), i })).sort((a, b) => a.z - b.z || a.i - b.i).map((x) => x.c);
}
function errMsg(err) {
  return err instanceof Error ? err.message : String(err);
}

// src/core/anim.ts
var EDGE_DIRS = { left: "left", right: "right", top: "top", bottom: "bottom" };
var VERT_DIRS = { top: "top", bottom: "bottom" };
var AXIS_DIRS = { horizontal: "horizontal", vertical: "vertical" };
var SHAPE_DIRS = { in: "in", out: "out" };
var CORNER_DIRS = {
  "down-right": "down-right",
  "down-left": "down-left",
  "up-right": "up-right",
  "up-left": "up-left"
};
var AXIS = { allowed: AXIS_DIRS, label: "horizontal|vertical" };
var SHAPE = { allowed: SHAPE_DIRS, label: "in|out" };
var CORNER = { allowed: CORNER_DIRS, def: "down-right", label: "down-right|down-left|up-right|up-left" };
var DIR_FAMILY = {
  "fly-in": { allowed: EDGE_DIRS, def: "bottom", label: "left|right|top|bottom" },
  "fly-out": { allowed: EDGE_DIRS, def: "bottom", label: "left|right|top|bottom" },
  "wipe-in": { allowed: EDGE_DIRS, def: "bottom", label: "left|right|top|bottom" },
  "wipe-out": { allowed: EDGE_DIRS, def: "bottom", label: "left|right|top|bottom" },
  "float-in": { allowed: VERT_DIRS, def: "bottom", label: "top|bottom" },
  "float-out": { allowed: VERT_DIRS, def: "bottom", label: "top|bottom" },
  "split-in": { ...AXIS, def: "vertical" },
  "split-out": { ...AXIS, def: "vertical" },
  "random-bars-in": { ...AXIS, def: "horizontal" },
  "random-bars-out": { ...AXIS, def: "horizontal" },
  "blinds-in": { ...AXIS, def: "horizontal" },
  "blinds-out": { ...AXIS, def: "horizontal" },
  "checkerboard-in": { ...AXIS, def: "horizontal" },
  "checkerboard-out": { ...AXIS, def: "horizontal" },
  "box-in": { ...SHAPE, def: "in" },
  "box-out": { ...SHAPE, def: "out" },
  "circle-in": { ...SHAPE, def: "in" },
  "circle-out": { ...SHAPE, def: "out" },
  "diamond-in": { ...SHAPE, def: "in" },
  "diamond-out": { ...SHAPE, def: "out" },
  "plus-in": { ...SHAPE, def: "in" },
  "plus-out": { ...SHAPE, def: "out" },
  "strips-in": CORNER,
  "strips-out": CORNER
};
var effectiveDurationMs = (def) => def.durationMs * (def.repeat ?? 1) * (def.autoReverse ? 2 : 1);
var frac = (n) => {
  const s = n.toFixed(5);
  return s === "-0.00000" ? "0.00000" : s;
};
function pathToOoxml(segs, slideWpx, slideHpx) {
  let out = "M 0 0";
  for (const seg of segs) {
    const vals = seg.p.map((v, k) => frac(v / (k % 2 === 0 ? slideWpx : slideHpx)));
    out += ` ${seg.c} ${vals.join(" ")}`;
  }
  return out + " E";
}

// src/render/timing.ts
var unrollPathSegs = (segs, repeat) => {
  const starts = [];
  let cur = [0, 0];
  for (const s of segs) {
    starts.push(cur);
    cur = s.c === "L" ? [s.p[0], s.p[1]] : [s.p[4], s.p[5]];
  }
  const reversed = [];
  for (let i = segs.length - 1; i >= 0; i--) {
    const s = segs[i];
    const [sx, sy] = starts[i];
    reversed.push(
      s.c === "L" ? { c: "L", p: [sx, sy] } : { c: "C", p: [s.p[2], s.p[3], s.p[0], s.p[1], sx, sy] }
    );
  }
  const cycle = [...segs, ...reversed];
  const out = [];
  for (let k = 0; k < Math.max(1, repeat); k++) out.push(...cycle);
  return out;
};
var PRESET = {
  appear: { pid: 1, cls: "entr" },
  disappear: { pid: 1, cls: "exit" },
  "fade-in": { pid: 10, cls: "entr" },
  "fade-out": { pid: 10, cls: "exit" },
  "fly-in": { pid: 2, cls: "entr" },
  "fly-out": { pid: 2, cls: "exit" },
  "wipe-in": { pid: 22, cls: "entr" },
  "wipe-out": { pid: 22, cls: "exit" },
  "float-in": { pid: 42, cls: "entr" },
  "float-out": { pid: 42, cls: "exit" },
  "split-in": { pid: 16, cls: "entr" },
  "split-out": { pid: 16, cls: "exit" },
  "bounce-in": { pid: 26, cls: "entr" },
  "bounce-out": { pid: 26, cls: "exit" },
  "zoom-in": { pid: 23, cls: "entr" },
  "zoom-out": { pid: 23, cls: "exit" },
  "wheel-in": { pid: 21, cls: "entr" },
  "wheel-out": { pid: 21, cls: "exit" },
  "random-bars-in": { pid: 14, cls: "entr" },
  "random-bars-out": { pid: 14, cls: "exit" },
  "blinds-in": { pid: 3, cls: "entr" },
  "blinds-out": { pid: 3, cls: "exit" },
  "checkerboard-in": { pid: 5, cls: "entr" },
  "checkerboard-out": { pid: 5, cls: "exit" },
  "dissolve-in": { pid: 9, cls: "entr" },
  "dissolve-out": { pid: 9, cls: "exit" },
  "box-in": { pid: 4, cls: "entr" },
  "box-out": { pid: 4, cls: "exit" },
  "circle-in": { pid: 6, cls: "entr" },
  "circle-out": { pid: 6, cls: "exit" },
  "diamond-in": { pid: 8, cls: "entr" },
  "diamond-out": { pid: 8, cls: "exit" },
  "plus-in": { pid: 13, cls: "entr" },
  "plus-out": { pid: 13, cls: "exit" },
  "strips-in": { pid: 18, cls: "entr" },
  "strips-out": { pid: 18, cls: "exit" },
  "wedge-in": { pid: 20, cls: "entr" },
  "wedge-out": { pid: 20, cls: "exit" },
  spin: { pid: 8, cls: "emph" },
  grow: { pid: 6, cls: "emph" },
  shrink: { pid: 6, cls: "emph" },
  pulse: { pid: 26, cls: "emph" },
  teeter: { pid: 32, cls: "emph" },
  path: { pid: 0, cls: "path" }
};
var DIR_FLAG = { top: 1, right: 2, bottom: 4, left: 8 };
var SPLIT_SUBTYPE = {
  "split-in": { vertical: 21, horizontal: 26 },
  "split-out": { vertical: 37, horizontal: 42 }
};
var AXIS_SUBTYPE = { horizontal: 10, vertical: 5 };
var SHAPE_SUBTYPE = { in: 16, out: 32 };
var CORNER_SUBTYPE = {
  "up-right": 3,
  "down-right": 6,
  "up-left": 9,
  "down-left": 12
};
var STRIPS_FILTER = {
  "down-right": "strips(downRight)",
  "down-left": "strips(downLeft)",
  "up-right": "strips(upRight)",
  "up-left": "strips(upLeft)"
};
var CHECKER_FILTER = {
  horizontal: "checkerboard(across)",
  vertical: "checkerboard(down)"
};
var WIPE_FILTER = {
  bottom: "wipe(up)",
  top: "wipe(down)",
  left: "wipe(right)",
  right: "wipe(left)"
};
var BARN_FILTER = {
  "split-in": { vertical: "barn(inVertical)", horizontal: "barn(inHorizontal)" },
  "split-out": { vertical: "barn(outVertical)", horizontal: "barn(outHorizontal)" }
};
var FLY_FROM = {
  left: { x: "0-#ppt_w/2", y: "#ppt_y" },
  right: { x: "1+#ppt_w/2", y: "#ppt_y" },
  top: { x: "#ppt_x", y: "0-#ppt_h/2" },
  bottom: { x: "#ppt_x", y: "1+#ppt_h/2" }
};
var FLOAT_FROM = {
  bottom: "#ppt_y+.1",
  top: "#ppt_y-.1"
};
var BOUNCE_IN_PATH = "M -0.25 -0.33333 C -0.195 -0.21 -0.14 -0.075 -0.085 0 C -0.064 -0.14815 -0.042 -0.14815 -0.021 0 C -0.0165 -0.04933 -0.012 -0.04933 -0.0075 0 C -0.005 -0.0164 -0.0025 -0.0164 0 0 E";
var BOUNCE_OUT_PATH = "M 0 0 C 0.0025 -0.0164 0.005 -0.0164 0.0075 0 C 0.012 -0.04933 0.0165 -0.04933 0.021 0 C 0.042 -0.14815 0.064 -0.14815 0.085 0 C 0.14 0.09 0.195 0.45 0.25 1.1 E";
function buildTimingXml(anims, slideWpx, slideHpx) {
  const live = anims.filter((a) => a.spids.length > 0);
  if (live.length === 0) return "";
  const sorted = [...live].sort((a, b) => a.def.order - b.def.order || a.def.index - b.def.index);
  let idCounter = 0;
  const nid = () => ++idCounter;
  nid();
  nid();
  const setVis = (spid, val, delay) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
  const bhvrCTn = (dur, delay, extra = "") => delay ? `<p:cTn id="${nid()}" dur="${dur}" fill="hold"${extra}><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>` : `<p:cTn id="${nid()}" dur="${dur}" fill="hold"${extra}/>`;
  const animEffect = (spid, dir, filter, dur, opts = {}) => `<p:animEffect transition="${dir}" filter="${filter}"><p:cBhvr${opts.tmFilter ? ` tmFilter="${opts.tmFilter}"` : ""}>` + (opts.delay ? `<p:cTn id="${nid()}" dur="${dur}"${opts.extra ?? ""}><p:stCondLst><p:cond delay="${opts.delay}"/></p:stCondLst></p:cTn>` : `<p:cTn id="${nid()}" dur="${dur}"${opts.extra ?? ""}/>`) + `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`;
  const flyAnim = (spid, attr, from, to, dur) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nid()}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const animScaleFromTo = (spid, from, to, dur) => `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr><p:from x="${from}" y="${from}"/><p:to x="${to}" y="${to}"/></p:animScale>`;
  const animScaleBy = (spid, by, dur, extra = "") => `<p:animScale><p:cBhvr>` + bhvrCTn(dur, void 0, extra) + `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr><p:by x="${by}" y="${by}"/></p:animScale>`;
  const animRotBy = (spid, by, dur, delay, extra = "") => `<p:animRot by="${by}"><p:cBhvr>` + bhvrCTn(dur, delay, extra) + `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
  const animMotionPath = (spid, path, dur, extra = "") => `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" ptsTypes=""><p:cBhvr>` + bhvrCTn(dur, void 0, extra) + `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
  const repAttrs = (def) => (def.repeat ? ` repeatCount="${def.repeat * 1e3}"` : "") + (def.autoReverse ? ` autoRev="1"` : "");
  const behaviors = (def, spid, secondary = false) => {
    const d = def.durationMs;
    const edge = def.dir ?? "bottom";
    const axis = def.dir;
    const hide = setVis(spid, "hidden", Math.max(d - 1, 0));
    switch (def.effect) {
      case "appear":
        return setVis(spid, "visible", 0);
      case "disappear":
        return setVis(spid, "hidden", 0);
      case "fade-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", "fade", d);
      case "fade-out":
        return animEffect(spid, "out", "fade", d) + hide;
      case "fly-in": {
        const from = FLY_FROM[edge];
        return setVis(spid, "visible", 0) + flyAnim(spid, "ppt_x", from.x, "#ppt_x", d) + flyAnim(spid, "ppt_y", from.y, "#ppt_y", d);
      }
      case "fly-out": {
        const to = FLY_FROM[edge];
        return flyAnim(spid, "ppt_x", "#ppt_x", to.x, d) + flyAnim(spid, "ppt_y", "#ppt_y", to.y, d) + hide;
      }
      case "wipe-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", WIPE_FILTER[edge], d);
      case "wipe-out":
        return animEffect(spid, "out", WIPE_FILTER[edge], d) + hide;
      case "float-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", "fade", d) + flyAnim(spid, "ppt_y", FLOAT_FROM[edge], "#ppt_y", d);
      case "float-out":
        return animEffect(spid, "out", "fade", d) + flyAnim(spid, "ppt_y", "#ppt_y", FLOAT_FROM[edge], d) + hide;
      case "split-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", BARN_FILTER["split-in"][axis], d);
      case "split-out":
        return animEffect(spid, "out", BARN_FILTER["split-out"][axis], d) + hide;
      case "bounce-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", "fade", Math.max(1, Math.round(d * 0.32))) + animMotionPath(spid, BOUNCE_IN_PATH, d);
      case "bounce-out":
        return animMotionPath(spid, BOUNCE_OUT_PATH, d) + animEffect(spid, "out", "fade", Math.max(1, Math.round(d * 0.2)), { delay: Math.round(d * 0.8) }) + hide;
      case "zoom-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", "fade", d) + animScaleFromTo(spid, 1e4, 1e5, d);
      case "zoom-out":
        return animEffect(spid, "out", "fade", d) + animScaleFromTo(spid, 1e5, 1e4, d) + hide;
      case "wheel-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : "wheel(1)", d);
      case "wheel-out":
        return animEffect(spid, "out", secondary ? "fade" : "wheel(1)", d) + hide;
      case "random-bars-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : `randombar(${axis})`, d);
      case "random-bars-out":
        return animEffect(spid, "out", secondary ? "fade" : `randombar(${axis})`, d) + hide;
      case "blinds-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : `blinds(${axis})`, d);
      case "blinds-out":
        return animEffect(spid, "out", secondary ? "fade" : `blinds(${axis})`, d) + hide;
      case "checkerboard-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : CHECKER_FILTER[axis], d);
      case "checkerboard-out":
        return animEffect(spid, "out", secondary ? "fade" : CHECKER_FILTER[axis], d) + hide;
      case "dissolve-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : "dissolve", d);
      case "dissolve-out":
        return animEffect(spid, "out", secondary ? "fade" : "dissolve", d) + hide;
      case "box-in":
      case "circle-in":
      case "diamond-in":
      case "plus-in": {
        const name = def.effect.slice(0, -3);
        const f = `${name}(${def.dir ?? "in"})`;
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : f, d);
      }
      case "box-out":
      case "circle-out":
      case "diamond-out":
      case "plus-out": {
        const name = def.effect.slice(0, -4);
        const f = `${name}(${def.dir ?? "out"})`;
        return animEffect(spid, "out", secondary ? "fade" : f, d) + hide;
      }
      case "strips-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : STRIPS_FILTER[def.dir ?? "down-right"], d);
      case "strips-out":
        return animEffect(spid, "out", secondary ? "fade" : STRIPS_FILTER[def.dir ?? "down-right"], d) + hide;
      case "wedge-in":
        return setVis(spid, "visible", 0) + animEffect(spid, "in", secondary ? "fade" : "wedge", d);
      case "wedge-out":
        return animEffect(spid, "out", secondary ? "fade" : "wedge", d) + hide;
      case "spin":
        return animRotBy(spid, Math.round((def.rotateDeg ?? 360) * 6e4), d, void 0, repAttrs(def));
      case "grow":
      case "shrink":
        return animScaleBy(spid, Math.round((def.scale ?? (def.effect === "grow" ? 1.5 : 0.67)) * 1e5), d, repAttrs(def));
      case "pulse":
        return animEffect(spid, "out", "fade", d, { tmFilter: "0,0; .2,.5; .8,.5; 1,0", extra: repAttrs(def) }) + animScaleBy(spid, Math.round((def.scale ?? 1.05) * 1e5), Math.max(1, Math.round(d / 2)), ` autoRev="1"` + repAttrs(def));
      case "teeter": {
        const r = Math.round((def.rotateDeg ?? 5) * 6e4);
        const t = (f) => Math.round(d * f);
        return animRotBy(spid, r, Math.max(1, t(0.1))) + animRotBy(spid, -2 * r, Math.max(1, t(0.2)), t(0.2)) + animRotBy(spid, 2 * r, Math.max(1, t(0.2)), t(0.4)) + animRotBy(spid, -2 * r, Math.max(1, t(0.2)), t(0.6)) + animRotBy(spid, r, Math.max(1, t(0.2)), t(0.8));
      }
      case "path": {
        const segs = def.pathSegs ?? [];
        if (def.autoReverse) {
          const unrolled = unrollPathSegs(segs, def.repeat ?? 1);
          return animMotionPath(spid, pathToOoxml(unrolled, slideWpx, slideHpx), effectiveDurationMs(def));
        }
        return animMotionPath(spid, pathToOoxml(segs, slideWpx, slideHpx), d, repAttrs(def));
      }
    }
  };
  const BEHAVIOR_LEVEL_REPEAT = /* @__PURE__ */ new Set(["spin", "grow", "shrink", "pulse", "path"]);
  const presetSubtype = (def) => {
    switch (def.effect) {
      case "fly-in":
      case "fly-out":
      case "wipe-in":
      case "wipe-out":
        return DIR_FLAG[def.dir ?? "bottom"];
      case "split-in":
      case "split-out":
        return SPLIT_SUBTYPE[def.effect][def.dir];
      case "random-bars-in":
      case "random-bars-out":
      case "blinds-in":
      case "blinds-out":
      case "checkerboard-in":
      case "checkerboard-out":
        return AXIS_SUBTYPE[def.dir];
      case "box-in":
      case "circle-in":
      case "diamond-in":
      case "plus-in":
        return SHAPE_SUBTYPE[def.dir ?? "in"];
      case "box-out":
      case "circle-out":
      case "diamond-out":
      case "plus-out":
        return SHAPE_SUBTYPE[def.dir ?? "out"];
      case "strips-in":
      case "strips-out":
        return CORNER_SUBTYPE[def.dir ?? "down-right"];
      case "wheel-in":
      case "wheel-out":
        return 1;
      // spoke count
      default:
        return 0;
    }
  };
  const effectPar = (def, spid, nodeType, secondary = false) => {
    const preset = PRESET[def.effect];
    const pid = (def.effect === "float-in" || def.effect === "float-out") && def.dir === "top" ? 47 : preset.pid;
    const parId = nid();
    let body = behaviors(def, spid, secondary);
    if (def.repeat && !BEHAVIOR_LEVEL_REPEAT.has(def.effect)) {
      body = `<p:par><p:cTn id="${nid()}" dur="${def.durationMs}" repeatCount="${def.repeat * 1e3}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`;
    }
    return `<p:par><p:cTn id="${parId}" presetID="${pid}" presetClass="${preset.cls}" presetSubtype="${presetSubtype(def)}" fill="hold" grpId="0" nodeType="${nodeType}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`;
  };
  const NODE_TYPE = { click: "clickEffect", with: "withEffect", after: "afterEffect" };
  const groups = [];
  let prevStart = 0;
  let groupEnd = 0;
  for (const a of sorted) {
    const opens = groups.length === 0 || a.def.trigger === "click";
    if (opens) {
      groups.push({ auto: a.def.trigger !== "click", items: [] });
      prevStart = 0;
      groupEnd = 0;
    }
    const g = groups[groups.length - 1];
    const start = opens ? a.def.delayMs : a.def.trigger === "with" ? prevStart + a.def.delayMs : groupEnd + a.def.delayMs;
    g.items.push({ def: a.def, spids: a.spids, start });
    prevStart = start;
    groupEnd = Math.max(groupEnd, start + effectiveDurationMs(a.def));
  }
  let groupsXml = "";
  for (const g of groups) {
    const gid = nid();
    let inner = "";
    for (const item of g.items) {
      const iid = nid();
      const nodeType = NODE_TYPE[item.def.trigger];
      const pars = item.spids.map((spid, i) => effectPar(item.def, spid, i === 0 ? nodeType : "withEffect", i > 0)).join("");
      inner += `<p:par><p:cTn id="${iid}" fill="hold"><p:stCondLst><p:cond delay="${item.start}"/></p:stCondLst><p:childTnLst>${pars}</p:childTnLst></p:cTn></p:par>`;
    }
    groupsXml += `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst><p:cond delay="${g.auto ? "0" : "indefinite"}"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
  }
  const bldSpids = [...new Set(sorted.flatMap((a) => a.spids))];
  const bldLst = `<p:bldLst>${bldSpids.map((s) => `<p:bldP spid="${s}" grpId="0" animBg="1"/>`).join("")}</p:bldLst>`;
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groupsXml}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`;
}

// src/render/build-editable.ts
async function buildEditablePptx(input, resolveMedia) {
  if (!Array.isArray(input.slides) || input.slides.length === 0) {
    throw new Error("buildEditablePptx: deck has no slides");
  }
  const warnings = [];
  const pptx = new PptxGenJS2();
  pptx.defineLayout({ name: "CUSTOM", width: pxToInches(input.width), height: pxToInches(input.height) });
  pptx.layout = "CUSTOM";
  const mediaCache = await buildMediaCache(input.slides, warnings, resolveMedia);
  let animations = 0;
  const animHidden = [];
  for (const [i, captured] of input.slides.entries()) {
    const slide = pptx.addSlide();
    const rootBg = parseColor(captured.root.style.backgroundColor);
    if (rootBg && rootBg.alpha === 1) {
      slide.background = { color: rootBg.hex };
      captured.root.style.backgroundColor = "transparent";
    }
    const animEntries = [];
    try {
      renderNodeToPptx(captured.root, {
        slide,
        slideW: input.width,
        slideH: input.height,
        originX: captured.rect.x,
        originY: captured.rect.y,
        mediaCache,
        warnings,
        slideNo: i + 1,
        animEntries,
        fontMap: input.fontMap
      });
    } catch (err) {
      warnings.push(`Slide render aborted: ${err instanceof Error ? err.message : String(err)}`);
    }
    if (captured.notes?.trim()) {
      try {
        slide.addNotes(captured.notes);
      } catch (err) {
        warnings.push(`addNotes failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
    const live = animEntries.filter((e) => e.spids.length > 0);
    for (const dead of animEntries) {
      if (dead.spids.length === 0) {
        animHidden.push(
          `Slide ${i + 1}: data-anim "${dead.def.effect}" element produced no exported shapes \u2014 animation dropped`
        );
      }
    }
    if (live.length > 0) {
      try {
        slide._timingXml = buildTimingXml(
          live.map((e) => ({ def: e.def, spids: e.spids })),
          input.width,
          input.height
        );
        animations += live.length;
      } catch (err) {
        warnings.push(
          `Slide ${i + 1} exported without animations \u2014 timing build failed: ${err instanceof Error ? err.message : String(err)}`
        );
      }
    }
  }
  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return { buffer, bytes: buffer.length, slides: input.slides.length, warnings, animations, animHidden };
}

// src/render/build-screenshot.ts
async function buildScreenshotPptx(slideTiles, notes, input) {
  const warnings = [];
  const pptx = new PptxGenJS2();
  pptx.defineLayout({ name: "CUSTOM", width: pxToInches(input.width), height: pxToInches(input.height) });
  pptx.layout = "CUSTOM";
  for (let i = 0; i < slideTiles.length; i++) {
    const slide = pptx.addSlide();
    for (const tile of slideTiles[i]) {
      slide.addImage({ data: tile.data, x: tile.x, y: tile.y, w: tile.w, h: tile.h });
    }
    const note = notes[i];
    if (note?.trim()) {
      try {
        slide.addNotes(note);
      } catch (err) {
        warnings.push(`addNotes slide ${i + 1}: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  }
  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return { buffer, bytes: buffer.length, slides: slideTiles.length, warnings, animations: 0, animHidden: [] };
}

// src/validate/validate.ts
function validate(setup, perSlide, slideTrees, input, mode) {
  const flags = [];
  if (!setup.fontsReady) {
    flags.push({
      kind: "fonts_timeout",
      message: "document.fonts.ready did not resolve within 8s \u2014 text metrics may use fallback fonts. Check that font URLs are reachable."
    });
  }
  const misses = Array.isArray(setup.fontSwapMisses) ? setup.fontSwapMisses : [];
  if (misses.length > 0) {
    const list = misses.join(", ");
    flags.push({
      kind: "font_swap_failed",
      message: mode === "screenshots" ? `Font swap target(s) ${list} never loaded \u2014 the Google Fonts CSS fetch failed or returned nothing for the family, or the face isn't installed locally. The screenshots were rendered with a fallback font. Retry with a corrected family name or a Google-served family (check fonts.google.com) \u2014 web-safe faces may not exist in this rendering environment either, so they are not a fix here. Tell the user plainly which fonts couldn't be applied.` : `Font swap target(s) ${list} never loaded \u2014 the Google Fonts CSS fetch failed or returned nothing for the family, or the face isn't installed locally. The exported file names these fonts, but layout used a fallback, so text sizing and wrapping may drift. Retry with a corrected family name (check fonts.google.com) or web-safe fonts, and tell the user plainly which fonts couldn't be applied.`
    });
  }
  if (mode === "editable" && input.resetTransformSelector && !setup.resetRect) {
    flags.push({
      kind: "reset_selector_miss",
      message: `resetTransformSelector ${JSON.stringify(input.resetTransformSelector)} matched nothing \u2014 capture may be scaled.`
    });
  } else if (setup.resetRect) {
    const dw = Math.abs(setup.resetRect.w - input.width);
    const dh = Math.abs(setup.resetRect.h - input.height);
    if (dw > 2 || dh > 2) {
      flags.push({
        kind: "slide_size_mismatch",
        message: `resetTransformSelector measures ${Math.round(setup.resetRect.w)}\xD7${Math.round(setup.resetRect.h)} after reset, expected ${input.width}\xD7${input.height}. Check the element isn't constrained by a parent's max-width/overflow.`
      });
    }
  }
  for (let i = 0; i < slideTrees.length; i++) {
    const rect = slideTrees[i].rect;
    if (Math.abs(rect.w - input.width) > 2 || Math.abs(rect.h - input.height) > 2) {
      flags.push({
        kind: "slide_size_mismatch",
        message: `Slide ${i + 1} root measures ${Math.round(rect.w)}\xD7${Math.round(rect.h)}, expected ${input.width}\xD7${input.height}. The selector may be matching a wrapper rather than the slide content, or the deck doesn't fix slide dimensions.`
      });
      break;
    }
  }
  const dupAdjacent = [];
  for (let i = 1; i < perSlide.length; i++) {
    if (perSlide[i].hash === perSlide[i - 1].hash) dupAdjacent.push(i);
  }
  if (dupAdjacent.length > 0) {
    flags.push({
      kind: "duplicate_adjacent",
      message: `Slides ${dupAdjacent.map((i) => `${i}/${i + 1}`).join(", ")} captured identically \u2014 showJs likely failed to navigate. Check the JS actually changes the visible slide; some decks need a longer delay for transitions.`
    });
  }
  const counts = /* @__PURE__ */ new Map();
  for (const r of perSlide) counts.set(r.hash, (counts.get(r.hash) ?? 0) + 1);
  const maxCount = Math.max(0, ...counts.values());
  if (perSlide.length >= 3 && maxCount > perSlide.length / 2) {
    flags.push({
      kind: "duplicate_majority",
      message: `${maxCount}/${perSlide.length} slides captured identically. The deck likely doesn't expose a JS navigation hook, or the showJs is wrong.`
    });
  }
  const notes = setup.notes;
  if (notes.length === 0) {
    flags.push({
      kind: "no_speaker_notes",
      message: "No speaker notes found in the deck (neither data-speaker-notes attributes nor a #speaker-notes JSON block). Expected if the deck has no notes."
    });
  } else {
    if (notes.length !== input.slides.length) {
      flags.push({
        kind: "notes_count_mismatch",
        message: `Speaker notes have ${notes.length} entries but ${input.slides.length} slides were requested \u2014 notes attach by index, so the tail will be missing or misalign.`
      });
    }
    const nonEmpty = notes.filter((n) => n.trim());
    if (nonEmpty.length >= 2 && new Set(nonEmpty).size === 1) {
      flags.push({
        kind: "notes_uniform_nonempty",
        message: `All ${nonEmpty.length} non-empty speaker notes are identical \u2014 likely a placeholder, not real notes.`
      });
    }
  }
  if (mode === "screenshots" && (setup.animCount ?? 0) > 0) {
    flags.push({
      kind: "animations_ignored_screenshots",
      message: `${setup.animCount} data-anim element(s) found, but mode="screenshots" exports flat images \u2014 animations were ignored. Use the editable mode to export them as native PowerPoint animations.`
    });
  }
  const failed = perSlide.reduce((sum, r) => sum + r.imagesFailed, 0);
  if (failed > 0) {
    const waited = perSlide.reduce((sum, r) => sum + r.imagesWaited, 0);
    flags.push({
      kind: "images_failed",
      message: `${failed}/${waited} images failed to decode before capture \u2014 they'll be missing from the export. Usually a 404 src or a CORS-blocked external URL.`
    });
  }
  return flags;
}

// src/orchestrator/driver.ts
var PlaywrightDriver = class _PlaywrightDriver {
  browser;
  context;
  page;
  constructor(browser, context, page) {
    this.browser = browser;
    this.context = context;
    this.page = page;
  }
  static async launch(url, opts) {
    const { chromium } = await import("playwright");
    const browser = await chromium.launch({ headless: true });
    const context = await browser.newContext({
      viewport: { width: opts.width, height: opts.height },
      deviceScaleFactor: opts.deviceScaleFactor ?? 1
    });
    const page = await context.newPage();
    page.setDefaultTimeout(opts.timeout ?? 3e4);
    await page.goto(url, { waitUntil: "load", timeout: opts.timeout ?? 3e4 });
    return new _PlaywrightDriver(browser, context, page);
  }
  /** Full-viewport PNG as a base64 data URL. (←capturePage().toDataURL()) */
  async screenshot() {
    const buf = await this.page.screenshot({ type: "png" });
    return `data:image/png;base64,${buf.toString("base64")}`;
  }
  async setViewportSize(width, height) {
    await this.page.setViewportSize({ width, height });
  }
  async close() {
    try {
      await this.context.close();
    } catch {
    }
    try {
      await this.browser.close();
    } catch {
    }
  }
};

// src/orchestrator/inject.ts
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
var cached = null;
function bundleSource() {
  if (cached == null) {
    const path = fileURLToPath(new URL("./capture.iife.js", import.meta.url));
    cached = readFileSync(path, "utf8");
  }
  return cached;
}
async function injectCaptureBundle(page) {
  await page.addScriptTag({ content: bundleSource() });
  const ok = await page.evaluate(
    () => typeof window.__genpptx === "object"
  );
  if (!ok) throw new Error("capture bundle failed to initialise window.__genpptx");
}
function evaluateWithTimeout(p, ms, label) {
  let timer;
  const guard = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms);
  });
  return Promise.race([p, guard]).finally(() => clearTimeout(timer));
}
function captureBudget(spec) {
  const d = Number(spec.delay);
  return 15e3 + (Number.isFinite(d) ? d : 600);
}
function callSetup(page, input) {
  return evaluateWithTimeout(
    page.evaluate((arg) => window.__genpptx.setup(arg), input),
    15e3,
    "setup"
  );
}
function callCaptureEditable(page, spec, fontSwaps) {
  return evaluateWithTimeout(
    page.evaluate(
      ([s, f]) => window.__genpptx.captureEditable(s, f),
      [spec, fontSwaps]
    ),
    captureBudget(spec),
    "editable capture"
  );
}
function callCaptureScreenshot(page, spec) {
  return evaluateWithTimeout(
    page.evaluate((s) => window.__genpptx.captureScreenshot(s), spec),
    captureBudget(spec),
    "screenshot capture"
  );
}
function callResolveMedia(page, refs) {
  return page.evaluate((r) => window.__genpptx.resolveMedia(r), refs);
}

// src/orchestrator/errors.ts
function stringifyError(err) {
  if (err instanceof Error) return err.message;
  try {
    return String(err);
  } catch {
    return "unknown error";
  }
}
function timeoutHint(message, phase, slideNum) {
  if (!/timed?\s*out|timeout/i.test(message)) return "";
  if (phase === "setup") {
    return ". The page didn't finish loading the deck in time \u2014 confirm the URL serves the deck, then retry.";
  }
  const where = `slide ${slideNum}`;
  return phase === "editable" ? `. Capture stalled on ${where} (slide too heavy for editable capture). Retrying identically will stall again \u2014 tell the user, then retry with mode:'screenshots' for a pixel-perfect non-editable fallback.` : `. Capture stalled on ${where}. Retrying identically may stall again \u2014 tell the user and consider a longer delay.`;
}

// src/orchestrator/run.ts
async function runGenPptx(rawInput, driver) {
  const mode = rawInput.mode ?? "editable";
  const input = {
    ...rawInput,
    hideSelectors: sanitizeStrings(rawInput.hideSelectors),
    googleFontImports: sanitizeStrings(rawInput.googleFontImports),
    fontSwaps: sanitizeFontSwaps(rawInput.fontSwaps)
  };
  if (!Array.isArray(input.slides) || input.slides.length === 0) {
    throw new Error("genPptx: slides[] is empty");
  }
  const { page } = driver;
  await injectCaptureBundle(page);
  if (mode === "screenshots") {
    await driver.setViewportSize(input.width, input.height);
    await page.evaluate(
      () => new Promise(
        (r) => requestAnimationFrame(() => requestAnimationFrame(() => r()))
      )
    );
  }
  let setupRes;
  try {
    setupRes = await callSetup(page, {
      mode,
      width: input.width,
      height: input.height,
      hideSelectors: input.hideSelectors,
      googleFontImports: input.googleFontImports,
      fontSwaps: input.fontSwaps,
      resetTransformSelector: mode === "screenshots" ? void 0 : input.resetTransformSelector
    });
  } catch (err) {
    const msg = stringifyError(err);
    throw new Error(`genPptx: setup failed: ${msg}${timeoutHint(msg, "setup")}`);
  }
  if (mode === "editable") {
    const captures = [];
    const capturedSlides = [];
    const animInvalid = [];
    const animNested = [];
    for (let i = 0; i < input.slides.length; i++) {
      let cap;
      try {
        cap = await callCaptureEditable(page, input.slides[i], input.fontSwaps ?? []);
      } catch (err) {
        const msg = stringifyError(err);
        throw new Error(
          `genPptx: slide ${i + 1}/${input.slides.length} capture failed: ${msg}${timeoutHint(msg, "editable", i + 1)}`
        );
      }
      const captured = { rect: cap.slide.rect, root: cap.slide.root };
      const note = setupRes.notes[i];
      if (note?.trim()) captured.notes = note;
      capturedSlides.push(captured);
      captures.push({ hash: cap.hash, imagesWaited: cap.imagesWaited, imagesFailed: cap.imagesFailed });
      for (const w of cap.animWarnings ?? []) animInvalid.push(`Slide ${i + 1}: ${w}`);
      for (const w of cap.animNested ?? []) animNested.push(`Slide ${i + 1}: ${w}`);
    }
    const build2 = await buildEditablePptx(
      { width: input.width, height: input.height, slides: capturedSlides },
      (refs) => callResolveMedia(page, refs)
    );
    const validation2 = validate(setupRes, captures, capturedSlides, input, "editable");
    const capped = (items) => items.slice(0, 8).join(" | ") + (items.length > 8 ? ` \u2026 and ${items.length - 8} more` : "");
    if (animInvalid.length > 0) {
      validation2.push({ kind: "animation_invalid", message: capped(animInvalid) });
    }
    if (animNested.length > 0) {
      validation2.push({
        kind: "animation_nested",
        message: `${capped(animNested)} \u2014 the innermost data-anim wins for that subtree`
      });
    }
    if (build2.animHidden.length > 0) {
      validation2.push({ kind: "animation_hidden_target", message: capped(build2.animHidden) });
    }
    return {
      buffer: build2.buffer,
      result: {
        bytes: build2.bytes,
        slides: build2.slides,
        warnings: build2.warnings,
        validation: validation2,
        speakerNotes: setupRes.notes.slice(0, input.slides.length),
        animations: build2.animations
      }
    };
  }
  const wIn = pxToInches(input.width);
  const hIn = pxToInches(input.height);
  const tiles = [];
  const hashes = [];
  for (let i = 0; i < input.slides.length; i++) {
    let dataUrl;
    try {
      await callCaptureScreenshot(page, input.slides[i]);
      dataUrl = await driver.screenshot();
    } catch (err) {
      const msg = stringifyError(err);
      throw new Error(
        `genPptx: slide ${i + 1}/${input.slides.length} capture failed: ${msg}${timeoutHint(msg, "screenshots", i + 1)}`
      );
    }
    let v = 5381;
    const mid = dataUrl.length >> 1;
    const slice = dataUrl.slice(mid, mid + 4096);
    for (let k = 0; k < slice.length; k++) v = (v << 5) + v + slice.charCodeAt(k) | 0;
    hashes.push({ hash: v >>> 0, imagesWaited: 0, imagesFailed: 0 });
    tiles.push([{ data: dataUrl, x: 0, y: 0, w: wIn, h: hIn }]);
  }
  const build = await buildScreenshotPptx(tiles, setupRes.notes, input);
  const validation = validate(setupRes, hashes, [], input, "screenshots");
  return {
    buffer: build.buffer,
    result: {
      bytes: build.bytes,
      slides: build.slides,
      warnings: build.warnings,
      validation,
      speakerNotes: setupRes.notes.slice(0, input.slides.length),
      animations: 0
    }
  };
}

// src/orchestrator/output.ts
import { mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";

// src/core/filename.ts
function safeBasename(filename, fallback) {
  const cleaned = (filename ?? "").normalize("NFC").replace(/[^\p{L}\p{N}\-_. ]/gu, "_").replace(/\s+/g, " ").replace(/^[.\s]+|[.\s]+$/g, "");
  return cleaned || fallback;
}

// src/orchestrator/output.ts
async function writeOutput(buffer, outDir, filename) {
  const base = safeBasename(filename, "deck").replace(/\.pptx$/i, "") || "deck";
  const name = `${base}.pptx`;
  const dir = resolve(outDir);
  await mkdir(dir, { recursive: true });
  const path = join(dir, name);
  await writeFile(path, buffer);
  return path;
}

// src/cli.ts
var SETUP_HINT = "cd <skill>/agents/gen-pptx && npm install && npx playwright install chromium";
function usage(msg) {
  process.stderr.write(
    `${msg}

Usage: gen-pptx --url <servedDeckUrl> --config <jsonPath|-> [--out <dir>]
`
  );
  process.exit(64);
}
function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--url") out.url = argv[++i];
    else if (a === "--config") out.config = argv[++i];
    else if (a === "--out") out.out = argv[++i];
    else if (a === "-h" || a === "--help") usage("gen-pptx");
    else usage(`Unknown argument: ${a}`);
  }
  if (!out.url) usage("Missing --url");
  if (!out.config) usage("Missing --config");
  if (!/^https?:\/\//i.test(out.url)) {
    usage("--url must be an http(s) URL (deck-stage / multi-file decks need a served origin, not file://)");
  }
  return out;
}
async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(chunk);
  return Buffer.concat(chunks).toString("utf8");
}
async function preflight() {
  const major = parseInt(process.versions.node.split(".")[0] ?? "0", 10);
  if (major < 18) {
    process.stderr.write(`gen-pptx: node >= 18 required (found ${process.versions.node}).
`);
    process.exit(1);
  }
  let pw;
  try {
    pw = await import("playwright");
  } catch {
    process.stderr.write(`gen-pptx: playwright is not installed.
One-time setup:
  ${SETUP_HINT}
`);
    process.exit(1);
  }
  let exe = "";
  try {
    exe = pw.chromium.executablePath();
  } catch {
  }
  if (!exe || !existsSync(exe)) {
    process.stderr.write(`gen-pptx: Chromium browser is not installed.
One-time setup:
  ${SETUP_HINT}
`);
    process.exit(1);
  }
}
async function main() {
  const args = parseArgs(process.argv.slice(2));
  await preflight();
  const raw = args.config === "-" ? await readStdin() : readFileSync2(resolve2(args.config), "utf8");
  let input;
  try {
    input = JSON.parse(raw);
  } catch (err) {
    usage(`--config is not valid JSON: ${err instanceof Error ? err.message : String(err)}`);
  }
  if (!input || typeof input.width !== "number" || typeof input.height !== "number") {
    usage("config must include numeric width and height");
  }
  const mode = input.mode ?? "editable";
  const driver = await PlaywrightDriver.launch(args.url, {
    width: input.width,
    height: input.height,
    deviceScaleFactor: mode === "screenshots" ? 2 : 1
  });
  try {
    const { result, buffer } = await runGenPptx(input, driver);
    const savedPath = await writeOutput(buffer, args.out ?? process.cwd(), input.filename);
    process.stdout.write(
      JSON.stringify({
        ok: true,
        file: savedPath,
        slides: result.slides,
        animations: result.animations,
        bytes: result.bytes,
        flags: result.validation.map((v) => ({ code: v.kind, message: v.message })),
        warnings: result.warnings,
        speakerNotes: result.speakerNotes
      }) + "\n"
    );
    process.exit(0);
  } catch (err) {
    process.stdout.write(
      JSON.stringify({ ok: false, error: err instanceof Error ? err.message : String(err) }) + "\n"
    );
    process.exit(1);
  } finally {
    await driver.close();
  }
}
main().catch((err) => {
  process.stdout.write(
    JSON.stringify({ ok: false, error: err instanceof Error ? err.message : String(err) }) + "\n"
  );
  process.exit(1);
});
