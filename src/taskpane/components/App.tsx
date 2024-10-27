import * as React from "react";
import { createWindow, convertToPdf } from "../taskpane";
import { makeStyles, Button, Label,} from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    height: "100vh",
    width: "100%",
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  button: {
    margin: "10px",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <form >
        <input id="studentDocCheckbox" type="checkbox" /> <Label>Create Student Document</Label>
        <input id="teacherDocCheckbox" type="checkbox" /> <Label>Create Teacher Document</Label>
      </form>
      
      <Button appearance="primary" className={styles.button} size="large" onClick={createWindow}>
        Create Document
      </Button>
      <Button id="convertToPdf" onClick={convertToPdf}>PDF</Button>

    </div>
  );
};

export default App;
