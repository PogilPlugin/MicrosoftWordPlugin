import * as React from "react";
import { createWindow } from "../taskpane";
import { makeStyles, Button } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    height: "100vh",
  },
  button: {
    margin: "25%",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <p>Input a file name.</p>
      <form>
        <input type="text" id="file"/>
      </form>
      <Button appearance="primary" className={styles.button} size="large" onClick={createWindow}>
        Create Document
      </Button>
    </div>
  );
};

export default App;
