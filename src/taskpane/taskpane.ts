let xmlData: string;

/**
 * `createDocs` get action depends doc type.
 * `type' can be student or teacher - 2 choice .
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
 * take xml data from word doc
 */
const getXml = async (context: Word.RequestContext) => {
  const body: Word.Body = context.document.body;
  const bodyOOXML = body.getOoxml();
  await context.sync();

  xmlData = bodyOOXML.value;
};

/**
 * create word doc and save content
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
 * make filter out for student document ( highlighted text is deleting)
 */
const filterStudentXML = (): string => {
  const Parser: DOMParser = new DOMParser();
  let xmlDocument: XMLDocument = Parser.parseFromString(xmlData, "application/xml");

  const xmlCollection = xmlDocument.getElementsByTagName("w:highlight");

  Array.from(xmlCollection).forEach((element) => {
    if (element.outerHTML.includes("cyan")) {
      const p = element.parentElement?.parentElement;
      if (p) {
        while (p.hasChildNodes()) {
          p.lastChild?.remove();
        }
      }
    }
  });

  const Serializer: XMLSerializer = new XMLSerializer();
  const parsedXml: string = Serializer.serializeToString(xmlDocument);

  return parsedXml;
};

/**
 *make filter out for teacher document ( protect text - filterout higligh).
 */
const filterTeacherXML = (): string => {
  const Parser: DOMParser = new DOMParser();
  let xmlDocument: XMLDocument = Parser.parseFromString(xmlData, "application/xml");

  const xmlCollection = xmlDocument.getElementsByTagName("w:highlight");

  Array.from(xmlCollection).forEach((element) => {
    if (element.outerHTML.includes("cyan")) {
      element.remove();
    }
  });

  const Serializer: XMLSerializer = new XMLSerializer();
  const parsedXml: string = Serializer.serializeToString(xmlDocument);

  return parsedXml;
};

/**
 * feedback message to user.
 */
function notify(message: string) {
  const text = document.getElementById("notificationText") as HTMLElement;
  if (text) {
    text.innerText = message;
  }
}

/**
 *make text as a teacher content.
 */
async function markSelection() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load(["font", "isEmpty"]);
    await context.sync();

    if (!selection.isEmpty) {
      selection.font.highlightColor = "Turquoise";
    }
    await context.sync();
  }).catch((error) => {
    console.error("Error:", error);
  });
}

export { createDocs, markSelection };
