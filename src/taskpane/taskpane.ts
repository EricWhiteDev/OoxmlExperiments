/* global Word console */

import { UpperLowerExperiments } from "./UpperLowerExperiments";
import { XDocument } from "ltxmlts";
//import { WmlPackage } from "openxmlsdkts";

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
        if (pf.alignment !== undefined) pfParts.push(`alignment=${pf.alignment}`);
        if (pf.firstLineIndent !== undefined) pfParts.push(`firstLineIndent=${pf.firstLineIndent}`);
        if (pf.leftIndent !== undefined) pfParts.push(`leftIndent=${pf.leftIndent}`);
        if (pf.rightIndent !== undefined) pfParts.push(`rightIndent=${pf.rightIndent}`);
        if (pf.lineSpacing !== undefined) pfParts.push(`lineSpacing=${pf.lineSpacing}`);
        if (pf.spaceBefore !== undefined) pfParts.push(`spaceBefore=${pf.spaceBefore}`);
        if (pf.spaceAfter !== undefined) pfParts.push(`spaceAfter=${pf.spaceAfter}`);
        if (pf.outlineLevel !== undefined) pfParts.push(`outlineLevel=${pf.outlineLevel}`);
        if (pfParts.length > 0) lines.push(`  paragraphFormat: ${pfParts.join(", ")}`);

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

export async function getEntireDocument(): Promise<string | null> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      await context.sync();

      const xDoc = XDocument.parse(ooxml.value);
      return xDoc.toStringWithIndentation();
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
