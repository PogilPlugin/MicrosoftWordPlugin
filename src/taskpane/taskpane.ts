let xmlData: string;

/**
 * `createDocs` get action depends on doc type.
 * `type' can be student or teacher - 2 choices.
 */
async function createDocs(type?: "student" | "teacher") {
  console.log("Begin: ", type);

  await Word.run(async (context) => {
    if (!type) {
      console.log("No document type selected.");
      notify("Error: No document type selected.");
      return;
    }

    notify("");

    console.log(`Making ${type} word doc:`);
    await getXml(context);

    if (type === "student") {
      makeDocument(context, filterStudentXML());
    } else if (type === "teacher") {
      makeDocument(context, filterTeacherXML());
    }
  });

  console.log("End;");
}

/**
 * Take XML data from Word document
 */
const getXml = async (context: Word.RequestContext) => {
  const body: Word.Body = context.document.body;
  const bodyOOXML = body.getOoxml();
  await context.sync();

  xmlData = bodyOOXML.value;
};

/**
 * Create Word document and save content
 */
const makeDocument = async (context: Word.RequestContext, content: string) => {
  const doc = context.application.createDocument();
  await context.sync();

  const docBody: Word.Body = doc.body;
  await context.sync();

  docBody.insertOoxml(content, Word.InsertLocation.start);
  await context.sync();

  doc.open();
  await context.sync();
};

/**
 * Filter XML for student document (highlighted text is removed)
 */
const filterStudentXML = (): string => {
  const Parser: DOMParser = new DOMParser();
  let xmlDocument: XMLDocument = Parser.parseFromString(xmlData, "application/xml");
  const body = xmlDocument.getElementsByTagName("w:body")[0];

  // Remove content before {{STUDENT START}}
  const studentStart = Array.from(xmlDocument.getElementsByTagName("w:t")).find((element) =>
    element.textContent?.includes("{{STUDENT START}}")
  );
  if (studentStart) {
    let currentNode = body.firstChild;
    while (currentNode && currentNode !== studentStart.parentElement?.parentElement) {
      const nextNode = currentNode.nextSibling;
      body.removeChild(currentNode);
      currentNode = nextNode;
    }
    studentStart.parentElement?.parentElement?.remove();
  }

  // Remove content after {{STUDENT STOP}}
  const studentStop = Array.from(xmlDocument.getElementsByTagName("w:t")).find((element) =>
    element.textContent?.includes("{{STUDENT STOP}}")
  );
  if (studentStop) {
    let currentNode = studentStop.parentElement?.parentElement?.nextSibling;
    while (currentNode) {
      const nextNode = currentNode.nextSibling;
      body.removeChild(currentNode);
      currentNode = nextNode;
    }
    studentStop.parentElement?.parentElement?.remove();
  }

  // Remove highlighted content 
  const highlights = Array.from(xmlDocument.getElementsByTagName("w:highlight"));
  highlights.forEach((element) => {
    if (element.outerHTML.includes("cyan")) {
      const parentParagraph = element.parentElement?.parentElement;
      if (parentParagraph) {
        while (parentParagraph.hasChildNodes()) {
          parentParagraph.lastChild?.remove();
        }
      }
    }
  });

  // Remove tags themselves
  removeTags(xmlDocument, ["{{TEACHER START}}", "{{TEACHER STOP}}", "{{STUDENT START}}", "{{STUDENT STOP}}"]);

  const Serializer: XMLSerializer = new XMLSerializer();
  return Serializer.serializeToString(xmlDocument);
};

/**
 * Filter XML for teacher document (highlighted text's highlight is removed)
 */
const filterTeacherXML = (): string => {
  const Parser: DOMParser = new DOMParser();
  let xmlDocument: XMLDocument = Parser.parseFromString(xmlData, "application/xml");
  const body = xmlDocument.getElementsByTagName("w:body")[0];

  // Remove content before {{TEACHER START}}
  const teacherStart = Array.from(xmlDocument.getElementsByTagName("w:t")).find((element) =>
    element.textContent?.includes("{{TEACHER START}}")
  );
  if (teacherStart) {
    let currentNode = body.firstChild;
    while (currentNode && currentNode !== teacherStart.parentElement?.parentElement) {
      const nextNode = currentNode.nextSibling;
      body.removeChild(currentNode);
      currentNode = nextNode;
    }
    teacherStart.parentElement?.parentElement?.remove();
  }

  // Remove content after {{TEACHER STOP}}
  const teacherStop = Array.from(xmlDocument.getElementsByTagName("w:t")).find((element) =>
    element.textContent?.includes("{{TEACHER STOP}}")
  );
  if (teacherStop) {
    let currentNode = teacherStop.parentElement?.parentElement?.nextSibling;
    while (currentNode) {
      const nextNode = currentNode.nextSibling;
      body.removeChild(currentNode);
      currentNode = nextNode;
    }
    teacherStop.parentElement?.parentElement?.remove();
  }

  // Remove highlight only (content remains)
  const highlights = Array.from(xmlDocument.getElementsByTagName("w:highlight"));
  highlights.forEach((element) => {
    if (element.outerHTML.includes("cyan")) {
      element.remove(); // Only remove highlight node
    }
  });

  // Remove tags themselves
  removeTags(xmlDocument, ["{{TEACHER START}}", "{{TEACHER STOP}}", "{{STUDENT START}}", "{{STUDENT STOP}}"]);

  const Serializer: XMLSerializer = new XMLSerializer();
  return Serializer.serializeToString(xmlDocument);
};

/**
 * Remove specific tags from XML document
 */
const removeTags = (xmlDocument: XMLDocument, tags: string[]) => {
  tags.forEach((tag) => {
    const elements = Array.from(xmlDocument.getElementsByTagName("w:t")).filter((element) =>
      element.textContent?.includes(tag)
    );
    elements.forEach((element) => {
      element.parentElement?.parentElement?.remove();
    });
  });
};

/**
 * Feedback message to user
 */
function notify(message: string) {
  const text = document.getElementById("notificationText") as HTMLElement;
  if (text) {
    text.innerText = message;
  }
}

/**
 * Mark selection as teacher content
 */
async function markSelection() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load(["font", "isEmpty"]);
    await context.sync();

    if (!selection.isEmpty) {
      selection.font.highlightColor = "null";
    }
    await context.sync();
  }).catch((error) => {
    console.error("Error:", error);
  });
}

export { createDocs, markSelection };
