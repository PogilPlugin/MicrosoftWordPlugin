/* global Word console */

// state
let XML_data : string;

export async function createWindow() {
  // Write text to the document.
    await Word.run(getData);

    const checkboxes = <HTMLInputElement>document.getElementById("checkboxes");

    console.log(checkboxes);

    //if ()
    //  await Word.run(makeStudentDocument);
    
    // if ()
    //   await Word.run(makeTeacherDocument);

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