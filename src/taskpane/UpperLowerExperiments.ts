import { WmlPackage, W } from "openxmlsdkts";

export class UpperLowerExperiments {
  static async entireDocumentToUpper(flatOpc: string): Promise<string> {
    const pkg = await WmlPackage.open(flatOpc);
    const mainPart = await pkg.mainDocumentPart();
    const xDoc = await mainPart.getXDocument();
    const root = xDoc.root;

    const body = root.element(W.body);
    const textElements = body.descendants(W.t);
    for (const tElement of textElements) {
      tElement.value = tElement.value.toUpperCase();
    }

    mainPart.putXDocument(xDoc);
    const result = await pkg.saveToFlatOpcAsync();
    return result;
  }

  static async entireDocumentToLower(flatOpc: string): Promise<string> {
    const pkg = await WmlPackage.open(flatOpc);
    const mainPart = await pkg.mainDocumentPart();
    const xDoc = await mainPart.getXDocument();
    const root = xDoc.root;

    const body = root.element(W.body);
    const textElements = body.descendants(W.t);
    for (const tElement of textElements) {
      tElement.value = tElement.value.toLowerCase();
    }

    mainPart.putXDocument(xDoc);
    const result = await pkg.saveToFlatOpcAsync();
    return result;
  }
}
