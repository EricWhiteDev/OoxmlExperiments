import React from "react";
import {
  Button, Dropdown, Option, makeStyles, RadioGroup, Radio,
} from "@fluentui/react-components";
import { entireDocumentToUpper, entireDocumentToLower, getPackageAsXml, getMainPart, getStyleDefPart, getNumPart, getStyleInfo, setStyleUsingOoxml, setParaStyleOnSelection, setRunStyleOnSelection, setStyleWrong, changeDefaultStyle, setDocumentBody } from "../taskpane";
import type { OoxmlSource } from "../taskpane";
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
  const [ooxmlSource, setOoxmlSource] = React.useState<OoxmlSource>("document");
  const [textViewerTitle, setTextViewerTitle] = React.useState("");
  const [textViewerText, setTextViewerText] = React.useState("");
  const onClickGetPackageAsXml = async () => {
    const result = await getPackageAsXml(ooxmlSource);
    if (result) {
      setTextViewerTitle("Package as XML");
      setTextViewerText(result);
    }
  };
  const onClickGetMainPart = async () => {
    const result = await getMainPart(ooxmlSource);
    if (result) {
      setTextViewerTitle("Main XDoc");
      setTextViewerText(result);
    }
  };
  const onClickGetStyleDefPart = async () => {
    const result = await getStyleDefPart(ooxmlSource);
    if (result) {
      setTextViewerTitle("Style Definitions Part");
      setTextViewerText(result);
    }
  };
  const onClickGetNumPart = async () => {
    const result = await getNumPart();
    if (result) {
      setTextViewerTitle("Numbering Part");
      setTextViewerText(result);
    }
  };
  const onClickSetRunStyleOnSelection = async () => {
    const result = await setRunStyleOnSelection();
    if (result) {
      setTextViewerTitle("Package After Setting Selection");
      setTextViewerText(result);
    }
  };
  const onClickSetParaStyleOnSelection = async () => {
    const result = await setParaStyleOnSelection();
    if (result) {
      setTextViewerTitle("Package After Setting Selection");
      setTextViewerText(result);
    }
  };
  const onClickChangeDefaultStyle = async () => {
    const result = await changeDefaultStyle();
    if (result) {
      setTextViewerTitle("Package After Change of Default Style");
      setTextViewerText(result);
    }
  };
  const onClickSetStyleWrong = async () => {
    const result = await setStyleWrong();
    if (result) {
      setTextViewerTitle("Package After Wrong Mod");
      setTextViewerText(result);
    }
  };
  const onClickSetStyleUsingOoxml = async () => {
    const result = await setStyleUsingOoxml();
    if (result) {
      setTextViewerTitle("Package After Mod");
      setTextViewerText(result);
    }
  };
  const onClickGetStyleInfo = async () => {
    const result = await getStyleInfo();
    if (result) {
      setTextViewerTitle("Styles");
      setTextViewerText(result);
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
      <div className={styles.spacer} />
      <hr style={{ border: "none", borderTop: "1px solid #e0e0e0", margin: 0 }} />
      <RadioGroup
        layout="horizontal"
        value={ooxmlSource}
        onChange={(_e, data) => setOoxmlSource(data.value as OoxmlSource)}
      >
        <Radio value="document" label="Get Entire Document" />
        <Radio value="selection" label="Get Selection" />
      </RadioGroup>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetPackageAsXml}>Get Package as XML</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetMainPart}>Get Main Part</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetStyleDefPart}>Get Style Def Part</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetNumPart}>Get Num Part</Button>
      </div>
      <hr style={{ border: "none", borderTop: "1px solid #e0e0e0", margin: 0 }} />
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickSetStyleUsingOoxml}>Set Style using Ooxml</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start", backgroundColor: "#d32f2f" }} onClick={onClickSetStyleWrong}>Set Style Wrong</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickChangeDefaultStyle}>Chg Dflt</Button>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start", backgroundColor: "#d32f2f" }} onClick={onClickSetParaStyleOnSelection}>Set Para Style on Sel</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickSetRunStyleOnSelection}>Set Run Style on Sel</Button>
      </div>
      <hr style={{ border: "none", borderTop: "1px solid #e0e0e0", margin: 0 }} />
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetStyleInfo}>Get Style Info from API</Button>
      </div>
      <div className={styles.textViewerSection}>
        <div className={styles.textViewerTitle}>{textViewerTitle}</div>
        <div className={styles.textDisplay}>{textViewerText}</div>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={() => navigator.clipboard.writeText(textViewerText)}>Copy to Clipboard</Button>
      </div>
    </div>
  );
};

export default App;
