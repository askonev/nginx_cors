var uniqueId =
  Date.now().toString(36) + Math.random().toString(36).substring(2);

var ip = "192.168.0.102";
var callbackurl = ""; // 'https://a54124be-9d23-4b1d-b43d-de66863b94f5.mock.pstmn.io' // 'https://eowzrjdoaq8tuyh.m.pipedream.net';

var file = "empty.docx";
// var file = "empty.pptx";
// var url = 'https://testing-documentserver-files.s3.amazonaws.com/public_documents/empty.docx'
var url = `http://${ip}:7080/files/${file}`;

// console.log(uniqueId)
// console.log(`file url: ${url}`);

var _type = "desktop";
// var _type = "mobile";

function createConnector() {
  window.connector = docEditor.createConnector();
}

window.docEditor = new DocsAPI.DocEditor("placeholder", {
  type: _type,
  document: {
    fileType: "docx",
    // fileType: "pptx", 
    info: {
      owner: "John Smith",
      favorite: true,
      folder: "Example Files",
      sharingSettings: [
        {
          permissions: "Full Access",
          user: "John Smith",
        },
        {
          isLink: true,
          permissions: "Read Only",
          user: "External link",
        },
      ],
    },
    key: uniqueId.toString(),
    title: "empty.docx",
    // title: "empty.pptx",
    url: url,
    permissions: {
      edit: true,
      download: true,
      review: true,
      comment: true,
    },
  },
  documentType: "word",
  // documentType: "slide",
  editorConfig: {
    mode: "edit",
    customization: {
      zoom: 100,
      integrationMode: "embed",
    },
    callbackUrl: callbackurl,
  },
  height: "100%",
  width: "100%",
  events: {
    onDocumentReady: createConnector,
    onMetaChange: onMetaChange,
  },
});
