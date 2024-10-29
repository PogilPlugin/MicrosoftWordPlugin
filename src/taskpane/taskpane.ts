/* global Word console */
import html2pdf from "html2pdf.js";

let xmlData: string;
let htmlData: string;
let checkboxes: { student: boolean; teacher: boolean; pdf: boolean } = { student: false, teacher: false, pdf: false };

// add in parsing

// run on startup
Office.onReady(() => {
  
});

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
        makePDF(xmlData);
      if (checkboxes.teacher)
        makePDF(xmlData);

    } else {
      console.log('Making new word doc:')
      await getXml(context);

      if (checkboxes.student)
        makeDocument(context, xmlData);
      if (checkboxes.teacher)
        makeDocument(context, xmlData);

    }

    return;
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

const makeDocument = async (context, content) => {
  const doc = context.application.createDocument();
  await context.sync();

  const docBody: Word.Body = doc.body;
  await context.sync();

  docBody.insertOoxml(content, Word.InsertLocation.start);
  await context.sync();

  doc.open();
  await context.sync();
};

const makePDF = async (content) => {
  const pdfOptions = {
    margin: 1,
    filename: "document.pdf",
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
  };

  html2pdf().from(content).set(pdfOptions).save();
};

export default createDocs;
