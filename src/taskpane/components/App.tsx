import * as React from "react";
import {createDocs, markSelection} from "../taskpane";
import { makeStyles, Button, Label, Text } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    height: "100vh",
    width: "100%",
    display: "flex",
    flexWrap: 'wrap',
    flexDirection: 'column'
  },
  button: {
    margin: "10px",
  },
  flex: {
    flex: '1'
  },
  section: {
    display: 'block',
    margin: '10px',
    textAlign: 'center',
  },

});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <form className={styles.flex}>
        <div className={styles.section}>
          <input id="studentDocCheckbox" type="checkbox" />
          <Label>Create Student Document</Label>
        </div>
        <div className={styles.section}>
          <input id="teacherDocCheckbox" type="checkbox" />
          <Label>Create Teacher Document</Label>
        </div>
      </form>

      <div className={styles.flex}>
        <Text id='notificationText' className={styles.section} ></Text>
      </div>

      <Button appearance="primary" className={styles.button} size="large" onClick={createDocs}>
        Create Documents
      </Button>

      <Button appearance="secondary" className={styles.button} size="large" onClick={markSelection}>
        Mark Selection As Teacher Content
      </Button>

    </div>
  );
};

export default App;
