/* global Word console */
import html2pdf from "html2pdf.js";

// state
let XML_data : string;

async function createWindow() {

  await Word.run(getData);

  await Word.run(getCheckboxes);

}

const getCheckboxes = async (context) => { 
  const studentDocCheckbox = <HTMLInputElement>document.getElementById("studentDocCheckbox");
  const teacherDocCheckbox = <HTMLInputElement>document.getElementById("teacherDocCheckbox");
  await context.sync();

  if (studentDocCheckbox.checked)
    await logToDoc(context, "student");
    await Word.run(makeStudentDocument);

  if (teacherDocCheckbox.checked)
    await logToDoc(context, "teacher");
    await Word.run(makeTeacherDocument);
}

const logToDoc = async (context, str: string) => {
  const body: Word.Body = context.document.body;
  body.insertText(str, Word.InsertLocation.start);

}

const getData = async (context) => {
  const body: Word.Body = context.document.body;
  const bodyOOXML = body.getOoxml();

  await context.sync();

  XML_data =  bodyOOXML.value;
}

const makeStudentDocument = async (context) => {
  const studentDocument = context.application.createDocument(); 
  await context.sync();
  
  const studentDocumentBody: Word.Body = studentDocument.body;
  await context.sync();

  studentDocumentBody.insertOoxml(await parseXML(), Word.InsertLocation.start);
  await context.sync();
  
  studentDocument.open();
  await context.sync();
}

const makeTeacherDocument = async (context) => {
  const teacherDocument = context.application.createDocument(); 
  await context.sync();
  
  const teacherDocumentBody: Word.Body = teacherDocument.body;
  await context.sync();

  teacherDocumentBody.insertOoxml(await parseXML(), Word.InsertLocation.start);
  await context.sync();
  
  teacherDocument.open();
  await context.sync();
}

const parseXML = async () => {
  return XML_data; 
}

// Word doc to pdf convert
async function convertToPdf() {
  try {
    await Word.run(async (context) => { 
      const body = context.document.body;
      const htmlContent = body.getHtml(); 
      await context.sync(); 
     
      const pdfOptions = {
        margin: 1,
        filename: 'document.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
      };
 
      html2pdf().from(htmlContent.value).set(pdfOptions).save();
      
    });
  } catch (error) {
  }
}

export {convertToPdf, createWindow};
