import React from "react";
import {
  Button, Dropdown, Option, makeStyles,
} from "@fluentui/react-components";
import { entireDocumentToUpper, entireDocumentToLower, getEntireDocument, getStyleInfo, setDocumentBody } from "../taskpane";
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
  const [textViewerTitle, setTextViewerTitle] = React.useState("");
  const [textViewerText, setTextViewerText] = React.useState("");
  const onClickGetEntireDocument = async () => {
    const result = await getEntireDocument();
    if (result) {
      setTextViewerTitle("Entire Document");
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
      <div className={styles.row}>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetEntireDocument}>Get Entire Document</Button>
        <Button appearance="primary" style={{ fontSize: "8pt", minWidth: 0, padding: "2px 6px", alignSelf: "flex-start" }} onClick={onClickGetStyleInfo}>Get Style Info</Button>
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
