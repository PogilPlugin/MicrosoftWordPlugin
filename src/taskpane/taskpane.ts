/* global Word console */
// state
let XML_data : string;
let newDocument: Word.DocumentCreated;

export async function createWindow() {
  // Write text to the document.
    await Word.run(getData);
    await Word.run(makeNewDocument);

    
}

const getData = async (context) => {
  const body: Word.Body = context.document.body;
  const bodyOOXML = body.getOoxml();

  await context.sync();

  XML_data =  bodyOOXML.value;
}

const makeNewDocument = async (context) => {
  newDocument = context.application.createDocument(); 
  await context.sync();
  
  const newDocBody: Word.Body = newDocument.body;
  await context.sync();

  newDocBody.insertOoxml(XML_data, Word.InsertLocation.start);
  await context.sync();
  
  newDocument.open();
}