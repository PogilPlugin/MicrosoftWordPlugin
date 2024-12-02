import * as React from "react";
import { createDocs, markSelection } from "../taskpane";
import { makeStyles, Button } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    height: "97vh",
    width: "100%",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px",
  },
  topButton: {
    alignSelf: "center",
    margin: "5px",
    color: "white",
    padding: "12px 20px",
    fontSize: "1rem",
  },
  middleButtons: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "10px",
    justifyContent: "center",
    alignItems: "center",
  },
  button: {
    padding: "8px 12px",
    fontSize: "0.85rem",
    textAlign: "center",
  },
  bottomButtons: {
    display: "flex",
    justifyContent: "center",
    gap: "5px",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  // Function to insert text at the current cursor position in Word
  const insertTextAtCursor = async (text: string) => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace); // Replace or insert at cursor position
      await context.sync();
    });
  };

  return (
    <div className={styles.root}>
      {/* Top Button */}
      <Button
        appearance="primary"
        className={styles.topButton}
        size="large"
        onClick={markSelection}
      >
        Mark Selection As Teacher Content
      </Button>

      {/* Middle Buttons */}
      <div className={styles.middleButtons}>
        <Button
          appearance="primary"
          className={styles.button}
          onClick={() => insertTextAtCursor("{{TEACHER START}}")}
        >
          Teacher Start
        </Button>
        <Button
          appearance="primary"
          className={styles.button}
          onClick={() => insertTextAtCursor("{{TEACHER STOP}}")}
        >
          Teacher Stop
        </Button>
        <Button
          appearance="primary"
          className={styles.button}
          onClick={() => insertTextAtCursor("{{STUDENT START}}")}
        >
          Student Start
        </Button>
        <Button
          appearance="primary"
          className={styles.button}
          onClick={() => insertTextAtCursor("{{STUDENT STOP}}")}
        >
          Student Stop
        </Button>
      </div>

      {/* Bottom Buttons */}
      <div className={styles.bottomButtons}>
        <Button
          appearance="primary"
          className={styles.button}
          size="small"
          onClick={() => createDocs("student")}
        >
          Create Student Document
        </Button>
        <Button
          appearance="primary"
          className={styles.button}
          size="small"
          onClick={() => createDocs("teacher")}
        >
          Create Teacher Document
        </Button>
      </div>
    </div>
  );
};

export default App;
