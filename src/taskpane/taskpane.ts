/* global Word, console */
import { jsPDF } from "jspdf";
import html2pdf from "html2pdf.js";

// state
let XML_data = '';

export async function createWindow() {
  // take text from word
  await Word.run(getData); // get data
  await Word.run(makeNewDocument); // make new doc
}

// Word belgesinin içeriğini almak
const getData = async (context) => {
  const body = context.document.body;
  const bodyOOXML = body.getOoxml(); //get ooml from word doc
  await context.sync();

  XML_data = bodyOOXML.value; // keep ooml data
};

// new word doc and add content
const makeNewDocument = async (context) => {
  const newDocument = context.application.createDocument(); // Yeni belge oluştur
  await context.sync();

  const newDocBody = newDocument.body;
  await context.sync();

  newDocBody.insertOoxml(XML_data, Word.InsertLocation.start); // OOXML verisini yeni belgenin başına ekle
  await context.sync();

  newDocument.open(); // open new file
  await context.sync();
};

// Word doc to pdf convert
export async function convertToPdf() { // func
  try {
    //get content from word
    await Word.run(async (context) => { 
      const body = context.document.body;
      const htmlContent = body.getHtml(); // take word doc to html format
      await context.sync(); //office js processing 
  /////////////////////////////////////////////////////////////////
      // HTMLto  PDF with  html2pdf////////////
      const pdfOptions = {
        margin: 1,
        filename: 'document.pdf', // i need to do set up with eli code
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
      };
  //////////////////////////////////////////////////////////////
      // take html content and make pdf 
      html2pdf().from(htmlContent.value).set(pdfOptions).save();
      //html2pdf() - library
      //from(htmlContent.value),- take html content and styles from converted html by word
      //set -- convert pdf 
      //save -- downlad
      console.log("pdf succesfull donwload");
    });
  } catch (error) {
    console.error('pdf download ERROR:', error);
  }
  ////////////////////////////////////////////////////////////////
}
