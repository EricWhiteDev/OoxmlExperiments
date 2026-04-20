/* global Word console */

import { UpperLowerExperiments } from "./UpperLowerExperiments";
import { XDocument, XElement, XAttribute, XNamespace } from "ltxmlts";
import { WmlPackage, W, W14 } from "openxmlsdkts";

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

function serializeWithoutTransientAttributes(xDoc: XDocument): string {
  const clone = new XDocument(xDoc);
  const root = clone.root;
  if (root) {
    for (const el of root.descendantsAndSelf()) {
      const toRemove: XAttribute[] = [];
      for (const attr of el.attributes()) {
        const name = attr.name;
        if (name.equals(W14.paraId) || name.equals(W14.textId) || name.localName.startsWith("rsid")) {
          toRemove.push(attr);
        }
      }
      for (const attr of toRemove) {
        attr.remove();
      }
    }
  }
  return clone.toStringWithIndentation();
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

export async function getCustomXmlInfo(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);
      const mainPart = await pkg.mainDocumentPart();
      const customParts = await mainPart.customXmlParts();

      const lines: string[] = [];
      lines.push(`Custom XML Parts: ${customParts.length}`);
      lines.push("");

      for (let i = 0; i < customParts.length; i++) {
        const part = customParts[i];
        lines.push(`--- Part ${i + 1} ---`);
        lines.push(`URI: ${part.getUri()}`);
        lines.push(`Content Type: ${part.getContentType()}`);

        const propsPart = await part.customXmlPropertiesPart();
        if (propsPart) {
          const propsXDoc = await propsPart.getXDocument();
          const propsRoot = propsXDoc.root;
          if (propsRoot) {
            const itemIdAttr = propsRoot.attributes().find(a => a.name.localName === "itemID");
            if (itemIdAttr) {
              lines.push(`Item ID: ${itemIdAttr.value}`);
            }
          }
        }

        const xDoc = await part.getXDocument();
        const root = xDoc.root;
        if (root) {
          lines.push(`Root Element: ${root.name.toString()}`);
          const nsAttrs = root.attributes().filter(a => a.name.localName.startsWith("xmlns") || a.name.namespaceName === "http://www.w3.org/2000/xmlns/");
          if (nsAttrs.length > 0) {
            lines.push("Namespaces:");
            for (const ns of nsAttrs) {
              lines.push(`  ${ns.name.localName}: ${ns.value}`);
            }
          }
        }

        lines.push(`XML:\n${xDoc.toStringWithIndentation()}`);
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
      return serializeWithoutTransientAttributes(xDoc);
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
      return serializeWithoutTransientAttributes(xDoc);
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function getNumPart(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();
      const pkg = await WmlPackage.open(ooxmlResult.value);
      const mainPart = await pkg.mainDocumentPart();
      const numPart = await mainPart.numberingDefinitionsPart();
      if (!numPart) {
        return null;
      }
      const xDoc = await numPart.getXDocument();
      return serializeWithoutTransientAttributes(xDoc);
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
      return serializeWithoutTransientAttributes(xDoc);
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

export async function setParaStyleOnSelection(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const ooxmlResult = selection.getOoxml();
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

      // Set the first paragraph's style to HappyBold
      const mainXDoc = await mainPart.getXDocument();
      const mainBody = mainXDoc.root!.element(W.body)!;
      const paragraphs = mainBody.elements(W.p);
      if (paragraphs.length >= 1) {
        const firstPara = paragraphs[0];
        let pPr = firstPara.element(W.pPr);
        if (!pPr) {
          pPr = new XElement(W.pPr);
          firstPara.addFirst(pPr);
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

      // Replace the selection with the modified XML
      selection.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();

      return displayXml;
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function setRunStyleOnSelection(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const ooxmlResult = selection.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);

      // Add the BoldingRun character style to the style definitions part
      const mainPart = await pkg.mainDocumentPart();
      const stylePart = await mainPart.styleDefinitionsPart();
      if (!stylePart) {
        return null;
      }
      const stylesXDoc = await stylePart.getXDocument();
      const stylesRoot = stylesXDoc.root!;

      const boldingRunStyle = new XElement(W.style,
        new XAttribute(W.type, "character"),
        new XAttribute(W.customStyle, "1"),
        new XAttribute(W.styleId, "BoldingRun"),
        new XElement(W._name, new XAttribute(W.val, "BoldingRun")),
        new XElement(W.basedOn, new XAttribute(W.val, "DefaultParagraphFont")),
        new XElement(W.uiPriority, new XAttribute(W.val, "1")),
        new XElement(W.qFormat),
        new XElement(W.rsid, new XAttribute(W.val, "00936926")),
        new XElement(W.rPr,
          new XElement(W.rFonts,
            new XAttribute(W.ascii, "Courier New"),
            new XAttribute(W.hAnsi, "Courier New"),
          ),
          new XElement(W.b),
          new XElement(W.i),
        ),
      );
      stylesRoot.add(boldingRunStyle);
      stylePart.putXDocument(stylesXDoc);

      // Set the first run's style to BoldingRun
      const mainXDoc = await mainPart.getXDocument();
      const mainBody = mainXDoc.root!.element(W.body)!;
      const runs = mainBody.descendants(W.r);
      if (runs.length >= 1) {
        const firstRun = runs[0];
        let rPr = firstRun.element(W.rPr);
        if (!rPr) {
          rPr = new XElement(W.rPr);
          firstRun.addFirst(rPr);
        }
        let rStyleEl = rPr.element(W.rStyle);
        if (rStyleEl) {
          rStyleEl.attribute(W.val)!.value = "BoldingRun";
        } else {
          rStyleEl = new XElement(W.rStyle, new XAttribute(W.val, "BoldingRun"));
          rPr.addFirst(rStyleEl);
        }
      }
      mainPart.putXDocument(mainXDoc);

      // Serialize for display (formatted) and for insertion (unformatted)
      const flatOpc = await pkg.saveToFlatOpcAsync();
      const displayXDoc = XDocument.parse(flatOpc);
      const displayXml = displayXDoc.toStringWithIndentation();

      // Replace the selection with the modified XML
      selection.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();

      return displayXml;
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function setNumberingStyle(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);
      const mainPart = await pkg.mainDocumentPart();

      // --- Modify paragraphs 4, 5, 6 (0-indexed: 3, 4, 5) ---
      const mainXDoc = await mainPart.getXDocument();
      const mainBody = mainXDoc.root!.element(W.body)!;
      const paragraphs = mainBody.elements(W.p);
      for (let i = 3; i <= 5 && i < paragraphs.length; i++) {
        const para = paragraphs[i];
        let pPr = para.element(W.pPr);
        if (!pPr) {
          pPr = new XElement(W.pPr);
          para.addFirst(pPr);
        }

        // Remove existing w:ind
        const indEl = pPr.element(W.ind);
        if (indEl) {
          indEl.remove();
        }

        // Remove existing pStyle and numPr if present
        const existingPStyle = pPr.element(W.pStyle);
        if (existingPStyle) {
          existingPStyle.remove();
        }
        const existingNumPr = pPr.element(W.numPr);
        if (existingNumPr) {
          existingNumPr.remove();
        }

        // Add pStyle and numPr at the beginning of pPr
        const numPrEl = new XElement(W.numPr,
          new XElement(W.ilvl, new XAttribute(W.val, "0")),
          new XElement(W.numId, new XAttribute(W.val, "2")),
        );
        pPr.addFirst(numPrEl);
        pPr.addFirst(new XElement(W.pStyle, new XAttribute(W.val, "ListParagraph")));
      }
      mainPart.putXDocument(mainXDoc);

      // --- Add abstractNum and num to numbering part ---
      const numPart = await mainPart.numberingDefinitionsPart();
      if (numPart) {
        const numXDoc = await numPart.getXDocument();
        const numberingRoot = numXDoc.root!;

        const W15 = XNamespace.get("http://schemas.microsoft.com/office/word/2012/wordml");
        const W16cid = XNamespace.get("http://schemas.microsoft.com/office/word/2016/wordml/cid");

        const abstractNumEl = new XElement(W.abstractNum,
          new XAttribute(W.abstractNumId, "1"),
          new XAttribute(W15.getName("restartNumberingAfterBreak"), "0"),
          new XElement(W.nsid, new XAttribute(W.val, "4C7C037F")),
          new XElement(W.multiLevelType, new XAttribute(W.val, "hybridMultilevel")),
          new XElement(W.tmpl, new XAttribute(W.val, "98D4747E")),
          new XElement(W.lvl, new XAttribute(W.ilvl, "0"), new XAttribute(W.tplc, "04090017"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerLetter")),
            new XElement(W.lvlText, new XAttribute(W.val, "%1)")),
            new XElement(W.lvlJc, new XAttribute(W.val, "left")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "1440"), new XAttribute(W.hanging, "360")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "1"), new XAttribute(W.tplc, "04090019"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerLetter")),
            new XElement(W.lvlText, new XAttribute(W.val, "%2.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "left")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "2160"), new XAttribute(W.hanging, "360")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "2"), new XAttribute(W.tplc, "0409001B"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerRoman")),
            new XElement(W.lvlText, new XAttribute(W.val, "%3.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "right")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "2880"), new XAttribute(W.hanging, "180")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "3"), new XAttribute(W.tplc, "0409000F"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "decimal")),
            new XElement(W.lvlText, new XAttribute(W.val, "%4.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "left")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "3600"), new XAttribute(W.hanging, "360")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "4"), new XAttribute(W.tplc, "04090019"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerLetter")),
            new XElement(W.lvlText, new XAttribute(W.val, "%5.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "left")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "4320"), new XAttribute(W.hanging, "360")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "5"), new XAttribute(W.tplc, "0409001B"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerRoman")),
            new XElement(W.lvlText, new XAttribute(W.val, "%6.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "right")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "5040"), new XAttribute(W.hanging, "180")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "6"), new XAttribute(W.tplc, "0409000F"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "decimal")),
            new XElement(W.lvlText, new XAttribute(W.val, "%7.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "left")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "5760"), new XAttribute(W.hanging, "360")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "7"), new XAttribute(W.tplc, "04090019"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerLetter")),
            new XElement(W.lvlText, new XAttribute(W.val, "%8.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "left")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "6480"), new XAttribute(W.hanging, "360")),
            ),
          ),
          new XElement(W.lvl, new XAttribute(W.ilvl, "8"), new XAttribute(W.tplc, "0409001B"), new XAttribute(W.tentative, "1"),
            new XElement(W.start, new XAttribute(W.val, "1")),
            new XElement(W.numFmt, new XAttribute(W.val, "lowerRoman")),
            new XElement(W.lvlText, new XAttribute(W.val, "%9.")),
            new XElement(W.lvlJc, new XAttribute(W.val, "right")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "7200"), new XAttribute(W.hanging, "180")),
            ),
          ),
        );

        // Insert as second child of w:numbering
        const children = numberingRoot.elements();
        if (children.length >= 1) {
          children[0].addAfterSelf(abstractNumEl);
        } else {
          numberingRoot.add(abstractNumEl);
        }

        // Add w:num as last child
        const numEl = new XElement(W.num,
          new XAttribute(W.numId, "2"),
          new XAttribute(W16cid.getName("durableId"), "708186538"),
          new XElement(W.abstractNumId, new XAttribute(W.val, "1")),
        );
        numberingRoot.add(numEl);

        numPart.putXDocument(numXDoc);
      }

      // --- Ensure ListParagraph style exists in styles part ---
      const stylePart = await mainPart.styleDefinitionsPart();
      if (stylePart) {
        const stylesXDoc = await stylePart.getXDocument();
        const stylesRoot = stylesXDoc.root!;
        const allStyles = stylesRoot.elements(W.style);
        let hasListParagraph = false;
        for (const s of allStyles) {
          const styleIdAttr = s.attribute(W.styleId);
          if (styleIdAttr && styleIdAttr.value === "ListParagraph") {
            hasListParagraph = true;
            break;
          }
        }
        if (!hasListParagraph) {
          const listParaStyle = new XElement(W.style,
            new XAttribute(W.type, "paragraph"),
            new XAttribute(W.styleId, "ListParagraph"),
            new XElement(W._name, new XAttribute(W.val, "List Paragraph")),
            new XElement(W.basedOn, new XAttribute(W.val, "Normal")),
            new XElement(W.uiPriority, new XAttribute(W.val, "34")),
            new XElement(W.qFormat),
            new XElement(W.rsid, new XAttribute(W.val, "00440BBD")),
            new XElement(W.pPr,
              new XElement(W.ind, new XAttribute(W.left, "720")),
              new XElement(W.contextualSpacing),
            ),
          );
          stylesRoot.add(listParaStyle);
          stylePart.putXDocument(stylesXDoc);
        }
      }

      // Serialize for display (formatted with transient attrs removed) and for insertion (unformatted)
      const flatOpc = await pkg.saveToFlatOpcAsync();
      const displayXDoc = XDocument.parse(flatOpc);
      const displayXml = serializeWithoutTransientAttributes(displayXDoc);

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

const LEGAL_WORDS = [
  "whereas", "hereinafter", "pursuant", "thereof", "hereof", "therein", "herein",
  "notwithstanding", "indemnify", "covenant", "obligation", "liability", "warranty",
  "representation", "jurisdiction", "arbitration", "adjudication", "enforceable",
  "severability", "termination", "confidential", "proprietary", "disclosure",
  "assignment", "sublicense", "intellectual", "infringement", "indemnification",
  "negligence", "damages", "remedies", "injunctive", "equitable", "monetary",
  "consideration", "counterpart", "execution", "performance", "compliance",
  "regulatory", "statutory", "contractual", "provisions", "stipulation",
  "acknowledgment", "affirmative", "mitigation", "subrogation", "indemnitor",
  "indemnitee", "guarantor", "surety", "collateral", "lien", "encumbrance",
  "subordination", "successor", "assignee", "licensor", "licensee", "grantor",
  "grantee", "lessor", "lessee", "mortgagor", "mortgagee", "obligor", "obligee",
  "plaintiff", "defendant", "claimant", "respondent", "arbitrator", "mediator",
  "adjudicator", "tribunal", "forum", "venue", "jurisdiction", "governing",
  "applicable", "enforceable", "irrevocable", "unconditional", "absolute",
  "perpetual", "exclusive", "nonexclusive", "transferable", "nontransferable",
  "revocable", "voidable", "material", "substantial", "reasonable", "customary",
  "commercially", "expeditiously", "forthwith", "promptly", "immediately",
  "contemporaneously", "simultaneously", "severally", "jointly", "explicitly",
  "implicitly", "expressly", "constructively",
];

function xmlSpace(): XAttribute {
  return new XAttribute(XNamespace.xml.getName("space"), "preserve");
}

function makeRunWithRpr(rPr: XElement | null, text: string): XElement {
  if (rPr) {
    return new XElement(W.r, new XElement(rPr), new XElement(W.t, xmlSpace(), text));
  }
  return new XElement(W.r, new XElement(W.t, xmlSpace(), text));
}

function addRandomRevTracking(mainBody: XElement): void {
  let idCounter = 0;
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");

  // Collect paragraphs with > 5 words, up to 20
  const eligibleParas: XElement[] = [];
  for (const para of mainBody.elements(W.p)) {
    const text = para.descendants(W.t).map((t) => t.value).join(" ");
    const wordCount = text.split(/\s+/).filter((w) => w.length > 0).length;
    if (wordCount > 5) {
      eligibleParas.push(para);
      if (eligibleParas.length === 20) {
        break;
      }
    }
  }

  for (const para of eligibleParas) {
    // --- Insert w:del ---
    const runsForDel = para.elements(W.r);

    type DelPoint =
      | { type: "between"; afterRun: XElement }
      | { type: "within"; run: XElement; pos: number };

    const delPoints: DelPoint[] = [];

    // Between adjacent runs
    for (let i = 0; i < runsForDel.length - 1; i++) {
      delPoints.push({ type: "between", afterRun: runsForDel[i] });
    }

    // Within a run at each space character
    for (const run of runsForDel) {
      const tEl = run.element(W.t);
      if (!tEl) {
        continue;
      }
      const text = tEl.value;
      for (let i = 0; i < text.length; i++) {
        if (text[i] === " ") {
          delPoints.push({ type: "within", run, pos: i + 1 });
        }
      }
    }

    if (delPoints.length > 0) {
      const point = delPoints[Math.floor(Math.random() * delPoints.length)];
      const delEl = new XElement(W.del,
        new XAttribute(W.id, String(idCounter++)),
        new XAttribute(W.author, "Eric White"),
        new XAttribute(W.date, now),
        new XElement(W.r,
          new XElement(W.delText, xmlSpace(), LEGAL_WORDS[Math.floor(Math.random() * LEGAL_WORDS.length)] + " "),
        ),
      );

      if (point.type === "between") {
        point.afterRun.addAfterSelf(delEl);
      } else {
        const run = point.run;
        const rPr = run.element(W.rPr);
        const text = run.element(W.t)!.value;
        const leftRun = makeRunWithRpr(rPr, text.slice(0, point.pos));
        const rightRun = makeRunWithRpr(rPr, text.slice(point.pos));
        run.replaceWith(leftRun, delEl, rightRun);
      }
    }

    // --- Insert w:ins ---
    // Re-collect runs (excludes runs inside w:del just added)
    const runsForIns = para.elements(W.r);

    type WordEntry = { run: XElement; wordStart: number; wordEnd: number };
    const wordEntries: WordEntry[] = [];

    for (const run of runsForIns) {
      const tEl = run.element(W.t);
      if (!tEl) {
        continue;
      }
      const text = tEl.value;
      const wordRe = /\S+\s*/g;
      let match: RegExpExecArray | null;
      while ((match = wordRe.exec(text)) !== null) {
        wordEntries.push({ run, wordStart: match.index, wordEnd: match.index + match[0].length });
      }
    }

    if (wordEntries.length > 0) {
      const entry = wordEntries[Math.floor(Math.random() * wordEntries.length)];
      const { run, wordStart, wordEnd } = entry;
      const rPr = run.element(W.rPr);
      const text = run.element(W.t)!.value;
      const wordText = text.slice(wordStart, wordEnd);
      const wordRun = makeRunWithRpr(rPr, wordText);

      const insEl = new XElement(W.ins,
        new XAttribute(W.id, String(idCounter++)),
        new XAttribute(W.author, "Eric White"),
        new XAttribute(W.date, now),
        wordRun,
      );

      if (wordStart === 0 && wordEnd === text.length) {
        // Word covers entire run — wrap the run directly
        run.replaceWith(insEl);
      } else {
        const pieces: XElement[] = [];
        if (wordStart > 0) {
          pieces.push(makeRunWithRpr(rPr, text.slice(0, wordStart)));
        }
        pieces.push(insEl);
        if (wordEnd < text.length) {
          pieces.push(makeRunWithRpr(rPr, text.slice(wordEnd)));
        }
        run.replaceWith(...pieces);
      }
    }
  }
}

export async function swapEveryOtherPara(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxmlResult.value);
      const mainPart = await pkg.mainDocumentPart();
      const mainXDoc = await mainPart.getXDocument();
      const mainBody = mainXDoc.root!.element(W.body)!;

      // Snapshot all child elements before mutating
      const children = mainBody.elements();
      const sectPr = children.find((el) => el.name.localName === "sectPr") ?? null;
      const nonSectPr = children.filter((el) => el.name.localName !== "sectPr");

      const newChildren = new Array<XElement>();
      const count = nonSectPr.length - 1;

      for (let i = 0; i < count; i += 2) {
        newChildren.push(children[i + 1]);
        newChildren.push(children[i]);
      }

      if (count % 2 === 1) {
        newChildren.push(children[count]);
      }
      newChildren.push(sectPr);

      mainBody.replaceNodes(newChildren);

      addRandomRevTracking(mainBody);

      mainPart.putXDocument(mainXDoc);

      // Formatted display: clone main XDoc, strip transient attrs, indent
      const displayXml = serializeWithoutTransientAttributes(mainXDoc);

      // Serialize without formatting and replace the document body
      const flatOpc = await pkg.saveToFlatOpcAsync();
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
