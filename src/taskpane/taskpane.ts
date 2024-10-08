/* global Word console */

// npm run start:desktop -- --app word --document "C:\Temp\Test.docx"

export async function createDocument() {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      console.log("Hello World");

      const externalDoc: Word.DocumentCreated = context.application.createDocument();
      await context.sync();

      const currentDocBody: Word.Body = context.document.body;
      currentDocBody.load("text");
      await context.sync();

      externalDoc.body.insertParagraph(currentDocBody.text, Word.InsertLocation.start);
      externalDoc.save('Save','C:\\Temp\\A.docx');

      externalDoc.open();

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
