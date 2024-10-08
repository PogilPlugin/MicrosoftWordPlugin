import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import {Button} from '@fluentui/react-components';
import { createDocument } from "../taskpane";
import { ProgressPlugin } from "webpack";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Button appearance="primary" disabled={false} size="large" onClick={createDocument}>
        Generate
      </Button>
    </div>
  );
};

export default App;
