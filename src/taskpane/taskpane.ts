/* 
Solution Explanation:
This solution ensures that separate filters for both Student and Teacher are applied to HTML content specifically for PDF conversion. 

Function Descriptions:
- `createDocs`: Main function that coordinates document creation based on user choices.
- `getCheckboxes`: Retrieves the selected options (Student, Teacher, PDF) from the user interface.
- `getXml` and `getHtml`: Fetches the document content in XML and HTML formats.
- `filteroutStudHligh`, `filteroutTeachHlight`: Filters for Student and Teacher in XML and HTML content for Word.
- `pdfStudentFilter`, `pdfTeacherFilter`: Filters for Student and Teacher in HTML content specifically for PDF conversion.
- `makeDocument`: Creates a new Word document with the specified content.
- `makePDF`: Generates and saves a PDF from HTML content with the applied filters.

--YIGIT TURAN
*/

/* global Word, console */
import html2pdf from "html2pdf.js";

// Variables to store document content
let xmlData: string; // Word document content in OOXML format
let htmlData: string; // Word document content in HTML format

// Stores user selections for document creation options
let checkboxes: { student: boolean; teacher: boolean; pdf: boolean } = { student: false, teacher: false, pdf: false };

// Initializes add-in when Office is ready
Office.onReady(() => {});

// Main function to create documents based on user selections
async function createDocs() {
  console.log("Begin: Document creation process");

  await Word.run(async (context) => {
    // Retrieve the checkbox selections from the UI
    await getCheckboxes(context);

    // Exit if no document type is selected
    if (!anyDocs()) {
      console.log("No Documents selected.");
      return;
    }

    console.log("Fetching document content...");

    // Fetch document content in both XML and HTML formats
    await getXml(context);
    await getHtml(context);

    // Apply filtering for Word or PDF based on selected options
    if (checkboxes.student) {
      xmlData = filteroutStudHligh(xmlData, "xml");
      htmlData = checkboxes.pdf ? pdfStudentFilter(htmlData) : filteroutStudHligh(htmlData, "html");
    }
    if (checkboxes.teacher) {
      xmlData = filteroutTeachHlight(xmlData, "xml");
      htmlData = checkboxes.pdf ? pdfTeacherFilter(htmlData) : filteroutTeachHlight(htmlData, "html");
    }

    // Create PDF if selected and skip creating Word document
    if (checkboxes.pdf) {
      console.log('Creating PDF document only:');
      makePDF(htmlData); // Use the filtered HTML data for PDF generation
    } else {
      // Create Word document if PDF is not selected
      console.log('Creating Word document only:');
      makeDocument(context, xmlData);
    }
  });

  console.log("End: Document creation process");
}

// Retrieves user checkbox selections from the document creation form
const getCheckboxes = async (context) => {
  const studentDocCheckbox = <HTMLInputElement>document.getElementById("studentDocCheckbox");
  const teacherDocCheckbox = <HTMLInputElement>document.getElementById("teacherDocCheckbox");
  const pdfDocCheckbox = <HTMLInputElement>document.getElementById("pdfDocCheckbox");
  await context.sync();

  checkboxes.student = studentDocCheckbox.checked;
  checkboxes.teacher = teacherDocCheckbox.checked;
  checkboxes.pdf = pdfDocCheckbox.checked;

  console.log(`Checkbox settings: { 'student': ${checkboxes.student}, 'teacher': ${checkboxes.teacher}, 'pdf': ${checkboxes.pdf} }`);
};

// Checks if any document creation option is selected
const anyDocs = (): boolean => {
  return checkboxes.teacher || checkboxes.student || checkboxes.pdf;
};

// Fetches the document content in XML format
const getXml = async (context) => {
  const body: Word.Body = context.document.body;
  const bodyOOXML = body.getOoxml();
  await context.sync();

  xmlData = bodyOOXML.value;
};

// Fetches the document content in HTML format
const getHtml = async (context) => {
  const body: Word.Body = context.document.body;
  const bodyHTML = body.getHtml();
  await context.sync();

  htmlData = bodyHTML.value;
};

// Filters out highlighted (yellow background) elements completely for Student in XML or HTML format for Word
const filteroutStudHligh = (content: string, format: string): string => {
  if (format === "xml") {
    return content.replace(/<w:r[^>]*><w:rPr><w:highlight w:val="yellow"\/><\/w:rPr>[\s\S]*?<\/w:r>/g, "");
  } else {
    let tempDiv = document.createElement("div");
    tempDiv.innerHTML = content;
    tempDiv.querySelectorAll("span[style*='background-color: yellow']").forEach((element) => {
      element.remove();
    });
    return tempDiv.innerHTML;
  }
};

// Removes only the yellow highlight, keeping the text, for Teacher in XML or HTML format for Word
const filteroutTeachHlight = (content: string, format: string): string => {
  if (format === "xml") {
    return content.replace(/<w:rPr><w:highlight w:val="yellow"\/><\/w:rPr>/g, "<w:rPr></w:rPr>");
  } else {
    let tempDiv = document.createElement("div");
    tempDiv.innerHTML = content;
    tempDiv.querySelectorAll<HTMLElement>("span[style*='background-color: yellow']").forEach((element) => {
      element.style.backgroundColor = "";
    });
    return tempDiv.innerHTML;
  }
};

// Filters out highlighted (yellow background) elements completely for Student in HTML format for PDF
const pdfStudentFilter = (content: string): string => {
  let tempDiv = document.createElement("div");
  tempDiv.innerHTML = content;

  // Removes elements with yellow background color in the HTML content
  tempDiv.querySelectorAll("span[style*='background-color: yellow']").forEach((element) => {
    element.remove();
  });

  return tempDiv.innerHTML;
};

// Removes only the yellow highlight, keeping the text, for Teacher in HTML format for PDF
const pdfTeacherFilter = (content: string): string => {
  let tempDiv = document.createElement("div");
  tempDiv.innerHTML = content;

  // Selects elements with yellow background color and removes only the background color style
  tempDiv.querySelectorAll<HTMLElement>("span[style*='background-color: yellow']").forEach((element) => {
    element.style.backgroundColor = "";
  });

  return tempDiv.innerHTML;
};

// Creates a new Word document with the specified content
const makeDocument = async (context, content) => {
  const doc = context.application.createDocument();
  await context.sync();

  const docBody: Word.Body = doc.body;
  await context.sync();

  // Inserts content into the new document
  docBody.insertOoxml(content, Word.InsertLocation.start);
  await context.sync();

  // Opens the new document in Word
  doc.open();
  await context.sync();
};

// Generates and saves a PDF from HTML content
const makePDF = async (content) => {
  const pdfOptions = {
    margin: 1,
    filename: "document.pdf",
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
  };

  // Create a temporary element to hold the content for conversion
  let tempElement = document.createElement("div");
  tempElement.innerHTML = content;

  // Pass the temporary element to html2pdf for PDF conversion
  html2pdf().from(tempElement).set(pdfOptions).save();
};

export default createDocs;
