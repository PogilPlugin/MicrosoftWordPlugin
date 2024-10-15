import * as React from "react";
import { createWindow, convertToPdf  } from "../taskpane"; //converttopdf addedd
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
//it is for checking to console message because i got error( it is good now)
const handlePdfClick = () => {
  console.log("PDF button clocked.");
  convertToPdf();
};

// convertopdf button added
const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <p>Input a file name.</p>
      <form>
        <input type="text" id="file"/>
      </form>
      <Button appearance="primary" className={styles.button} size="large" onClick={createWindow}>
        Create Converted PDF 
      </Button>
      <button id="convertToPdf" onClick={handlePdfClick}>PDF'ye Dönüştür</button>

    </div>
  );
};

export default App;
