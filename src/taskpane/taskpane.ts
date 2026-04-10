/* global Word console */

import { UpperLowerExperiments } from "./UpperLowerExperiments";
import { BoldRunsExperiment } from "./BoldRunsExperiment";
import { TestDocuments } from "./TestDocuments";
import { WmlPackage, W } from "openxmlsdkts";
import { XElement } from "ltxmlts";
import { base64Image } from "./TestImages";

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

export async function appendToSecondParagraph() {
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load();
      await context.sync();

      paragraphs.items[1].insertText(
        " New sentence in the paragraph.",
        Word.InsertLocation.end
      );
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function insertParagraph() {
  try {
    await Word.run(async (context) => {
      const docBody = context.document.body;
      docBody.insertParagraph("Hello World", Word.InsertLocation.start);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function intenseReference() {
  try {
    await Word.run(async (context) => {
      const firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.styleBuiltIn = Word.BuiltInStyleName.intenseReference;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function getAndSetOoxml() {
  try {
    await Word.run(async (context) => {
      const firstParagraph = context.document.body.paragraphs.getFirst();
      const ooxml = firstParagraph.getOoxml();
      await context.sync();

      const xml = await BoldRunsExperiment.boldAllRuns(ooxml.value);

      firstParagraph.insertOoxml(xml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function setFont() {
  try {
    await Word.run(async (context) => {
      const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      secondParagraph.font.set({
        name: "Courier New",
        bold: true,
        size: 18
      });
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}



export async function ooxmlPartialPara(): Promise<{ title: string; text: string } | null> {
  try {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const firstPara = paragraphs.items[0];
      const thirdPara = paragraphs.items[2];

      firstPara.load("text");
      thirdPara.load("text");
      await context.sync();

      const firstMid = Math.floor(firstPara.text.length / 2);
      const thirdMid = Math.floor(thirdPara.text.length / 2);

      const startRange = firstPara.getRange("Start").expandTo(firstPara.getRange("End"));
      const searchStart = firstPara.search(firstPara.text.substring(0, firstMid), { matchCase: true });
      searchStart.load("items");
      await context.sync();

      const selectionStart = searchStart.items.length > 0
        ? searchStart.items[0].getRange("End")
        : firstPara.getRange("Start");

      const searchEnd = thirdPara.search(thirdPara.text.substring(0, thirdMid), { matchCase: true });
      searchEnd.load("items");
      await context.sync();

      const selectionEnd = searchEnd.items.length > 0
        ? searchEnd.items[0].getRange("End")
        : thirdPara.getRange("End");

      const selectedRange = selectionStart.expandTo(selectionEnd);
      selectedRange.select();
      const ooxml = selectedRange.getOoxml();
      await context.sync();

      const pkg = await WmlPackage.open(ooxml.value);
      const mainPart = await pkg.mainDocumentPart();
      const xDoc = await mainPart.getXDocument();
      const indentedXml = xDoc.toStringWithIndentation();
      const root = xDoc.root!;
      const body = root.element(W.body)!;

      const runs = body.descendants(W.r);
      for (const run of runs) {
        let rPr = run.element(W.rPr);
        if (!rPr) {
          rPr = new XElement(W.rPr);
          run.addFirst(rPr);
        }
        rPr.add(new XElement(W.b));
      }

      mainPart.putXDocument(xDoc);
      const modifiedXml = await pkg.saveToFlatOpcAsync();

      selectedRange.insertOoxml(modifiedXml, Word.InsertLocation.replace);
      await context.sync();

      return { title: "Main Document Part", text: indentedXml };
    });
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function modContentControl() {
  try {
    await Word.run(async (context) => {
      const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
      serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function createContentControl() {
  try {
    await Word.run(async (context) => {
      const serviceNameRange = context.document.getSelection();
      const serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Service Name";
      serviceNameContentControl.tag = "serviceName";
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function modTableCell22() {
  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) return;

      const cell = tables.items[0].getCell(1, 1);
      cell.body.clear();
      cell.body.insertText("Biz Buzz", Word.InsertLocation.start);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function addTable() {
  try {
    await Word.run(async (context) => {
      const tableData = [
        ["Name", "ID", "Birth City"],
        ["Bob", "434", "Chicago"],
        ["Sue", "719", "Havana"],
      ];

      const range = context.document.getSelection();
      const paragraph = range.paragraphs.getFirst();
      paragraph.insertTable(tableData.length, tableData[0].length, Word.InsertLocation.after, tableData);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function modTableCell() {
  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) return;

      const cell = tables.items[0].getCell(0, 0);
      cell.body.clear();
      cell.body.insertText("Foo Bar", Word.InsertLocation.start);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function addHtml() {
  try {
    await Word.run(async (context) => {
      const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
      blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function addImage() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
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
