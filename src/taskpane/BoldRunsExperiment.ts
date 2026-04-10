import { WmlPackage, W } from "openxmlsdkts";
import { XElement } from "ltxmlts";

export class BoldRunsExperiment {
  static async boldAllRuns(flatOpc: string): Promise<string> {
    const pkg = await WmlPackage.open(flatOpc);
    const mainPart = await pkg.mainDocumentPart();
    const xDoc = await mainPart.getXDocument();
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
    return await pkg.saveToFlatOpcAsync();
  }
}
