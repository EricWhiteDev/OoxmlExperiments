import React from "react";
import {
  Button, Dropdown, Option, makeStyles,
  Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions,
} from "@fluentui/react-components";
import { entireDocumentToUpper, entireDocumentToLower, appendToSecondParagraph, insertParagraph, intenseReference, getAndSetOoxml, setFont, ooxmlPartialPara, modContentControl, createContentControl, modTableCell, modTableCell22, addTable, addHtml, addImage, setDocumentBody } from "../taskpane";
import { TestDocuments } from "../TestDocuments";


const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    padding: "16px",
    gap: "10px",
    boxSizing: "border-box",
    height: "100vh",
  },
  row: {
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    gap: "6px",
  },
  dropdown: {
    minWidth: 0,
    flexShrink: 0,
  },
  spacer: {
    flex: 1,
  },
  separator: {
    border: "none",
    borderTop: "1px solid #e0e0e0",
    marginTop: "0",
    marginBottom: "0",
  },
  textViewerSection: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minHeight: 0,
    flex: 1,
  },
  textViewerTitle: {
    fontSize: "12pt",
    fontWeight: "bold",
    margin: "0",
  },
  textDisplay: {
    flex: 1,
    fontFamily: "monospace",
    fontSize: "12px",
    whiteSpace: "pre",
    overflow: "auto",
    backgroundColor: "#f5f5f5",
    padding: "8px",
    borderRadius: "4px",
    border: "1px solid #e0e0e0",
    minHeight: "100px",
  },
});

const App: React.FC = () => {
  const styles = useStyles();

  const onClickButton4 = () => entireDocumentToUpper();
  const onClickButton5 = () => entireDocumentToLower();
  const onClickBodyParagraphs = () => appendToSecondParagraph();
  const onClickInsertParagraph = () => insertParagraph();
  const [intenseDialogOpen, setIntenseDialogOpen] = React.useState(false);
  const onClickIntenseReference = async () => {
    await intenseReference();
    setIntenseDialogOpen(true);
  };
  const onClickGetAndSetOoxml = () => getAndSetOoxml();
  const onClickSetFont = () => setFont();
  const [textViewerTitle, setTextViewerTitle] = React.useState("");
  const [textViewerText, setTextViewerText] = React.useState("");
  const onClickOoxmlPartialPara = async () => {
    const result = await ooxmlPartialPara();
    if (result) {
      setTextViewerTitle(result.title);
      setTextViewerText(result.text);
    }
  };

  return (
    <div className={styles.root}>
      <Dropdown
        className={styles.dropdown}
        placeholder="Select a document..."
        onOptionSelect={(_e, data) => {
          const xml = TestDocuments.testDocumentsMap.get(data.optionValue as string);
          if (xml) {
            setDocumentBody(xml);
          }
        }}
      >
        {Array.from(TestDocuments.testDocumentsMap.keys()).map((key) => (
          <Option key={key} value={key}>{key}</Option>
        ))}
      </Dropdown>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickButton4}>To Upper Case</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickButton5}>To Lower Case</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickBodyParagraphs}>insertText</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickInsertParagraph}>insertParagraph</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickIntenseReference}>intenseReference</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetAndSetOoxml}>getAndSetOoxml</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickSetFont}>setFont</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickOoxmlPartialPara}>ooxmlPartialPara</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => addImage()}>addImage</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => addHtml()}>addHtml</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => addTable()}>addTable</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => modTableCell()}>modTableCell-1-1</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => modTableCell22()}>modTableCell-2-2</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => createContentControl()}>createContentControl</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => modContentControl()}>modContentControl</Button>
      </div>
      <Dialog open={intenseDialogOpen} onOpenChange={(_e, data) => setIntenseDialogOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>intenseReference</DialogTitle>
            <DialogContent>Changes the color, sets to caps, and sets to bold</DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setIntenseDialogOpen(false)}>OK</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      <div className={styles.spacer} />
      <hr className={styles.separator} />
      <div className={styles.textViewerSection}>
        <div className={styles.textViewerTitle}>{textViewerTitle}</div>
        <div className={styles.textDisplay}>{textViewerText}</div>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => navigator.clipboard.writeText(textViewerText)}>Copy to Clipboard</Button>
      </div>
    </div>
  );
};

export default App;
