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


let xmlData: string;
let checkboxes: { student: boolean; teacher: boolean;} = { student: false, teacher: false};

async function createDocs() {
  console.log("Begin: ");

  await Word.run(async (context) => {
    
    await getCheckboxes(context);

    if (!anyDocs()) {
      console.log("No Documents selected.");
      notify("Error: No Documents selected.")
      return;
    }

    notify("");

    console.log('Making new word doc:')
    await getXml(context);

    if (checkboxes.student)
      makeDocument(context, filterStudentXML());
    if (checkboxes.teacher)
      makeDocument(context, filterTeacherXML());
    
  });

  console.log("End;");
}

const getCheckboxes = async (context) => {
  const studentDocCheckbox = <HTMLInputElement>document.getElementById("studentDocCheckbox");
  const teacherDocCheckbox = <HTMLInputElement>document.getElementById("teacherDocCheckbox");
  await context.sync();

  checkboxes.student = studentDocCheckbox.checked;
  checkboxes.teacher = teacherDocCheckbox.checked;
  console.log(`settings: { 'student': ${checkboxes.student}, 'teacher': ${checkboxes.teacher}}`);
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



// Filters out highlighted (yellow background) elements completely for Student in XML or HTML format for Word
// TODO: make work inside table 
const filterStudentXML = (): string => {
  return xmlData.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:highlight w:val="cyan"\/>[\s\S]*?<\/w:p>/, "");
};

// Removes only the yellow highlight, keeping the text, for Teacher in XML or HTML format for Word
const filterTeacherXML = (): string => {
  return xmlData.replace(/<w:rPr><w:highlight w:val="cyan"\/><\/w:rPr>/g, "<w:rPr></w:rPr>");
};


function notify(message: string) {
  const text = <HTMLElement>document.getElementById("notificationText");
  text.innerText = message;
}


async function markSelection() {
  await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load(['font', 'isEmpty']);
      await context.sync();

      if (!selection.isEmpty) 
        selection.font.highlightColor = 'Turquoise';
      await context.sync();

      
  }).catch((error) => {
      console.error("Error:", error);
  });
}
export {createDocs, markSelection};
