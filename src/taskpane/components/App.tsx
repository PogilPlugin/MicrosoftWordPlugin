import * as React from "react";
import { createDocs, markSelection } from "../taskpane";
import { makeStyles, Button } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    height: "100vh",
    width: "100%",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px",
  },
  //mark as selection button style
  topButton: {
    alignSelf: "center",
    margin: "5px",
    backgroundColor: "red",
    color: "white",
    padding: "12px 20px",
    fontSize: "1rem",
    ":hover": {
      backgroundColor: "darkred",
    },
  },

  //create docs buttons
  bottomButtons: {
    display: "flex",
    justifyContent: "center",
    gap: "5px",
  },
  button: {
    padding: "8px 12px",
    fontSize: "0.85rem",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  const handleCreateStudentDoc = async () => {
    await createDocs("student");  // student doc button 
  };

  const handleCreateTeacherDoc = async () => {
    await createDocs("teacher"); // teacher doc button 
  };

  return (
    <div className={styles.root}>
      <Button
        appearance="secondary"
        className={styles.topButton}
        size="large" // because alone in horizantally 
        onClick={markSelection}
      >
        Mark Selection As Teacher Content
      </Button>
      <div className={styles.bottomButtons}>
        <Button
          appearance="primary"
          className={styles.button}
          size="small" //because there is 2 button in x axis
          onClick={handleCreateStudentDoc}
        >
          Create Student Document
        </Button>
        <Button
          appearance="primary"
          className={styles.button}
          size="small" // because there is 2 button in some spot
          onClick={handleCreateTeacherDoc}
        >
          Create Teacher Document
        </Button>
      </div>
    </div>
  );
};

export default App;
