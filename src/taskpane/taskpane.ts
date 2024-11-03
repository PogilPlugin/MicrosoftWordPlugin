/* 
Solution Explanation:
This solution ensures that separate filters for both Student and Teacher are applied to HTML content specifically for PDF conversion. 

Function Descriptions:
- `createDocs`: Main function that coordinates document creation based on user choices.
- `getCheckboxes`: Retrieves the selected options (Student, Teacher, PDF) from the user interface.
- `getXml` and `getHtml`: Fetches the document content in XML and HTML formats.
- `filterStudentXML`, `filterTeacherXML`: Filters for Student and Teacher in XML and HTML content for Word.
- `pdfStudentFilter`, `pdfTeacherFilter`: Filters for Student and Teacher in HTML content specifically for PDF conversion.
- `makeDocument`: Creates a new Word document with the specified content.
- `makePDF`: Generates and saves a PDF from HTML content with the applied filters.

--YIGIT TURAN
*/


/* global Word console */
import html2pdf from "html2pdf.js";

let xmlData: string;
let htmlData: string;
let checkboxes: { student: boolean; teacher: boolean; pdf: boolean } = { student: false, teacher: false, pdf: false };

async function createDocs() {
  console.log("Begin: ");

  await Word.run(async (context) => {
    
    await getCheckboxes(context);

    if (!anyDocs()) {
      console.log("No Documents selected.");
      return;
    }

    if (checkboxes.pdf) {
      console.log('Making new pdf doc:')
      await getHtml(context);

      if (checkboxes.student)
        makePDF(pdfStudentFilter());
      if (checkboxes.teacher)
        makePDF(pdfTeacherFilter());

    } else {
      console.log('Making new word doc:')
      await getXml(context);

      if (checkboxes.student)
        makeDocument(context, filterStudentXML());
      if (checkboxes.teacher)
        makeDocument(context, filterTeacherXML());

    }
  });

  console.log("End;");
}

const getCheckboxes = async (context) => {
  const studentDocCheckbox = <HTMLInputElement>document.getElementById("studentDocCheckbox");
  const teacherDocCheckbox = <HTMLInputElement>document.getElementById("teacherDocCheckbox");
  const pdfDocCheckbox = <HTMLInputElement>document.getElementById("pdfDocCheckbox");
  await context.sync();

  checkboxes.student = studentDocCheckbox.checked;
  checkboxes.teacher = teacherDocCheckbox.checked;
  checkboxes.pdf = pdfDocCheckbox.checked;
  console.log(`settings: { 'student': ${checkboxes.student}, 'teacher': ${checkboxes.teacher},  'pdf': ${checkboxes.pdf},}`);
};

const anyDocs = (): boolean => {
  return checkboxes.teacher || checkboxes.student;
};

const logToDoc = async (context, str: string) => {
  const body: Word.Body = context.document.body;
  body.insertText(str, Word.InsertLocation.start);
};

const getXml = async (context) => {
  const body: Word.Body = context.document.body;
  const bodyOOXML = body.getOoxml();
  await context.sync();

  xmlData = bodyOOXML.value;
};

const getHtml = async (context) => {
  const body: Word.Body = context.document.body;
  const bodyHTML = body.getHtml();
  await context.sync();

  htmlData = bodyHTML.value;
};

const makeDocument = async (context, content: string) => {
  const doc = context.application.createDocument();
  await context.sync();

  const docBody: Word.Body = doc.body;
  await context.sync();

  docBody.insertOoxml(content, Word.InsertLocation.start);
  await context.sync();

  doc.open();
  await context.sync();
};

const makePDF = async (content: string) => {
  const pdfOptions = {
    margin: 1,
    filename: "document.pdf",
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
  };

  html2pdf().from(content).set(pdfOptions).save();
};


// Filters out highlighted (yellow background) elements completely for Student in XML or HTML format for Word
const filterStudentXML = (): string => {
  return xmlData.replace(/<w:r[^>]*><w:rPr><w:highlight w:val="yellow"\/><\/w:rPr>[\s\S]*?<\/w:r>/g, "");
};

// Removes only the yellow highlight, keeping the text, for Teacher in XML or HTML format for Word
const filterTeacherXML = (): string => {
  return xmlData.replace(/<w:rPr><w:highlight w:val="yellow"\/><\/w:rPr>/g, "<w:rPr></w:rPr>");
};

// Filters out highlighted (yellow background) elements completely for Student in HTML format for PDF
const pdfStudentFilter = (): string => {
  let tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlData;

  // Removes elements with yellow background color in the HTML content
  tempDiv.querySelectorAll("span[style*='background-color: yellow']").forEach((element) => {
    element.remove();
  });

  return tempDiv.innerHTML;
};

// Removes only the yellow highlight, keeping the text, for Teacher in HTML format for PDF
const pdfTeacherFilter = (): string => {
  let tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlData;

  // Selects elements with yellow background color and removes only the background color style
  tempDiv.querySelectorAll<HTMLElement>("span[style*='background-color: yellow']").forEach((element) => {
    element.style.backgroundColor = "";
  });

  return tempDiv.innerHTML;
};

export default createDocs;
