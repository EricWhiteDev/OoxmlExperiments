/* global Word console */

import { UpperLowerExperiments } from "./UpperLowerExperiments";
import { XDocument, XElement, XAttribute } from "ltxmlts";
import { WmlPackage, W } from "openxmlsdkts";

export type OoxmlSource = "document" | "selection";

export async function getOoxml(
  context: Word.RequestContext,
  source: OoxmlSource,
): Promise<string> {
  let ooxmlResult: OfficeExtension.ClientResult<string>;
  if (source === "selection") {
    ooxmlResult = context.document.getSelection().getOoxml();
  } else {
    ooxmlResult = context.document.body.getOoxml();
  }
  await context.sync();
  return ooxmlResult.value;
}

export async function entireDocumentToUpper() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      await context.sync();

      const xml = await UpperLowerExperiments.entireDocumentToUpper(ooxml.value);

      body.insertOoxml(xml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function entireDocumentToLower() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      await context.sync();

      const xml = await UpperLowerExperiments.entireDocumentToLower(ooxml.value);

      body.insertOoxml(xml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}




export async function getStyleInfo(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const styles = context.document.getStyles();
      styles.load("items");
      await context.sync();

      for (const style of styles.items) {
        style.load([
          "nameLocal",
          "baseStyle",
          "linked",
          "listLevelNumber",
          "type",
        ]);
        style.font.load("name");
        style.paragraphFormat.load([
          "alignment",
          "firstLineIndent",
          "leftIndent",
          "rightIndent",
          "lineSpacing",
          "spaceAfter",
          "spaceBefore",
          "outlineLevel",
        ]);
      }
      await context.sync();

      const lines: string[] = [];
      for (const style of styles.items) {
        lines.push(style.nameLocal);

        let baseStyle = "";
        try { baseStyle = style.baseStyle; } catch (_e) { /* no base style */ }
        lines.push(`  baseStyle: ${baseStyle}`);

        lines.push(`  font name: ${style.font.name}`);
        lines.push(`  linked: ${style.linked}`);
        lines.push(`  listLevelNumber: ${style.listLevelNumber}`);

        const pf = style.paragraphFormat;
        const pfParts: string[] = [];
        if (pf.alignment !== undefined) {
          pfParts.push(`alignment=${pf.alignment}`);
        }
        if (pf.firstLineIndent !== undefined) {
          pfParts.push(`firstLineIndent=${pf.firstLineIndent}`);
        }
        if (pf.leftIndent !== undefined) {
          pfParts.push(`leftIndent=${pf.leftIndent}`);
        }
        if (pf.rightIndent !== undefined) {
          pfParts.push(`rightIndent=${pf.rightIndent}`);
        }
        if (pf.lineSpacing !== undefined) {
          pfParts.push(`lineSpacing=${pf.lineSpacing}`);
        }
        if (pf.spaceBefore !== undefined) {
          pfParts.push(`spaceBefore=${pf.spaceBefore}`);
        }
        if (pf.spaceAfter !== undefined) {
          pfParts.push(`spaceAfter=${pf.spaceAfter}`);
        }
        if (pf.outlineLevel !== undefined) {
          pfParts.push(`outlineLevel=${pf.outlineLevel}`);
        }
        if (pfParts.length > 0) {
          lines.push(`  paragraphFormat: ${pfParts.join(", ")}`);
        }

        lines.push(`  type: ${style.type}`);
        lines.push("");
      }

      return lines.join("\n");
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function getPackageAsXml(source: OoxmlSource): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const ooxml = await getOoxml(context, source);
      const xDoc = XDocument.parse(ooxml);
      return xDoc.toStringWithIndentation();
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function getMainPart(source: OoxmlSource): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const ooxml = await getOoxml(context, source);
      const pkg = await WmlPackage.open(ooxml);
      const mainPart = await pkg.mainDocumentPart();
      const xDoc = await mainPart.getXDocument();
      return xDoc.toStringWithIndentation();
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function getStyleDefPart(source: OoxmlSource): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const ooxml = await getOoxml(context, source);
      const pkg = await WmlPackage.open(ooxml);
      const mainPart = await pkg.mainDocumentPart();
      const stylePart = await mainPart.styleDefinitionsPart();
      if (!stylePart) {
        return null;
      }
      const xDoc = await stylePart.getXDocument();
      return xDoc.toStringWithIndentation();
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function setStyleUsingOoxml(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);

      // Add the HappyBold style to the style definitions part
      const mainPart = await pkg.mainDocumentPart();
      const stylePart = await mainPart.styleDefinitionsPart();
      if (!stylePart) {
        return null;
      }
      const stylesXDoc = await stylePart.getXDocument();
      const stylesRoot = stylesXDoc.root!;

      const happyBoldStyle = new XElement(W.style,
        new XAttribute(W.type, "paragraph"),
        new XAttribute(W.customStyle, "1"),
        new XAttribute(W.styleId, "HappyBold"),
        new XElement(W._name, new XAttribute(W.val, "HappyBold")),
        new XElement(W.basedOn, new XAttribute(W.val, "Normal")),
        new XElement(W.qFormat),
        new XElement(W.rsid, new XAttribute(W.val, "00084F40")),
        new XElement(W.rPr,
          new XElement(W.rFonts,
            new XAttribute(W.ascii, "Courier New"),
            new XAttribute(W.hAnsi, "Courier New"),
          ),
          new XElement(W.b),
          new XElement(W.i),
        ),
      );
      stylesRoot.add(happyBoldStyle);
      stylePart.putXDocument(stylesXDoc);

      // Set the 3rd paragraph's style to HappyBold
      const mainXDoc = await mainPart.getXDocument();
      const mainBody = mainXDoc.root!.element(W.body)!;
      const paragraphs = mainBody.elements(W.p);
      if (paragraphs.length >= 3) {
        const thirdPara = paragraphs[2];
        let pPr = thirdPara.element(W.pPr);
        if (!pPr) {
          pPr = new XElement(W.pPr);
          thirdPara.addFirst(pPr);
        }
        let pStyleEl = pPr.element(W.pStyle);
        if (pStyleEl) {
          pStyleEl.attribute(W.val)!.value = "HappyBold";
        } else {
          pStyleEl = new XElement(W.pStyle, new XAttribute(W.val, "HappyBold"));
          pPr.addFirst(pStyleEl);
        }
      }
      mainPart.putXDocument(mainXDoc);

      // Serialize for display (formatted) and for insertion (unformatted)
      const flatOpc = await pkg.saveToFlatOpcAsync();
      const displayXDoc = XDocument.parse(flatOpc);
      const displayXml = displayXDoc.toStringWithIndentation();

      // Put the modified document back into Word
      body.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();

      return displayXml;
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function changeDefaultStyle(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);
      const mainPart = await pkg.mainDocumentPart();
      const stylePart = await mainPart.styleDefinitionsPart();
      if (!stylePart) {
        return null;
      }
      const stylesXDoc = await stylePart.getXDocument();
      const stylesRoot = stylesXDoc.root!;

      // Find the paragraph style with w:default='1' and remove the default attribute
      const allStyles = stylesRoot.elements(W.style);
      for (const s of allStyles) {
        const typeAttr = s.attribute(W.type);
        const defaultAttr = s.attribute(W._default);
        if (typeAttr && typeAttr.value === "paragraph" && defaultAttr && defaultAttr.value === "1") {
          defaultAttr.remove();
          break;
        }
      }

      // Add the HappyBold style as the new default paragraph style
      const happyBoldStyle = new XElement(W.style,
        new XAttribute(W.type, "paragraph"),
        new XAttribute(W._default, "1"),
        new XAttribute(W.customStyle, "1"),
        new XAttribute(W.styleId, "HappyBold"),
        new XElement(W._name, new XAttribute(W.val, "HappyBold")),
        new XElement(W.basedOn, new XAttribute(W.val, "Normal")),
        new XElement(W.qFormat),
        new XElement(W.rsid, new XAttribute(W.val, "00084F40")),
        new XElement(W.rPr,
          new XElement(W.rFonts,
            new XAttribute(W.ascii, "Courier New"),
            new XAttribute(W.hAnsi, "Courier New"),
          ),
          new XElement(W.b),
          new XElement(W.i),
        ),
      );
      stylesRoot.add(happyBoldStyle);
      stylePart.putXDocument(stylesXDoc);

      // Serialize for display (formatted) and for insertion (unformatted)
      const flatOpc = await pkg.saveToFlatOpcAsync();
      const displayXDoc = XDocument.parse(flatOpc);
      const displayXml = displayXDoc.toStringWithIndentation();

      // Put the modified document back into Word
      body.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();

      return displayXml;
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function setStyleWrong(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);
      const mainPart = await pkg.mainDocumentPart();

      // Set the 3rd paragraph's style to HappyBold (without adding the style definition)
      const mainXDoc = await mainPart.getXDocument();
      const mainBody = mainXDoc.root!.element(W.body)!;
      const paragraphs = mainBody.elements(W.p);
      if (paragraphs.length >= 3) {
        const thirdPara = paragraphs[2];
        let pPr = thirdPara.element(W.pPr);
        if (!pPr) {
          pPr = new XElement(W.pPr);
          thirdPara.addFirst(pPr);
        }
        let pStyleEl = pPr.element(W.pStyle);
        if (pStyleEl) {
          pStyleEl.attribute(W.val)!.value = "HappyBold";
        } else {
          pStyleEl = new XElement(W.pStyle, new XAttribute(W.val, "HappyBold"));
          pPr.addFirst(pStyleEl);
        }
      }
      mainPart.putXDocument(mainXDoc);

      // Serialize for display (formatted) and for insertion (unformatted)
      const flatOpc = await pkg.saveToFlatOpcAsync();
      const displayXDoc = XDocument.parse(flatOpc);
      const displayXml = displayXDoc.toStringWithIndentation();

      // Put the modified document back into Word
      body.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();

      return displayXml;
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function setDocumentBody(xml: string) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertOoxml(xml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function selectRange(
  startParagraphIdx: number,
  startCharIdx: number,
  endParagraphIdx: number,
  endCharIdx: number
) {
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const startPara = paragraphs.items[startParagraphIdx];
      const endPara = paragraphs.items[endParagraphIdx];

      startPara.load("text");
      endPara.load("text");
      await context.sync();

      let range: Word.Range;
      if (startParagraphIdx === endParagraphIdx && startCharIdx === endCharIdx) {
        // Collapse to insertion point
        if (startCharIdx === 0) {
          range = startPara.getRange("Start");
        } else {
          const textBefore = startPara.text.substring(0, startCharIdx);
          const searchResults = startPara.search(textBefore, { matchCase: true, matchWholeWord: false });
          searchResults.load("items");
          await context.sync();
          if (searchResults.items.length > 0) {
            range = searchResults.items[0].getRange("End");
          } else {
            range = startPara.getRange("Start");
          }
        }
      } else {
        // Build start position
        let rangeStart: Word.Range;
        if (startCharIdx === 0) {
          rangeStart = startPara.getRange("Start");
        } else {
          const textBefore = startPara.text.substring(0, startCharIdx);
          const searchResults = startPara.search(textBefore, { matchCase: true, matchWholeWord: false });
          searchResults.load("items");
          await context.sync();
          rangeStart = searchResults.items.length > 0
            ? searchResults.items[0].getRange("End")
            : startPara.getRange("Start");
        }

        // Build end position
        let rangeEnd: Word.Range;
        if (endCharIdx === 0) {
          rangeEnd = endPara.getRange("Start");
        } else {
          const textBefore = endPara.text.substring(0, endCharIdx);
          const searchResults = endPara.search(textBefore, { matchCase: true, matchWholeWord: false });
          searchResults.load("items");
          await context.sync();
          rangeEnd = searchResults.items.length > 0
            ? searchResults.items[0].getRange("End")
            : endPara.getRange("End");
        }

        range = rangeStart.expandTo(rangeEnd);
      }

      range.select();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
