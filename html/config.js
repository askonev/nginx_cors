var uniqueId = Date.now().toString(36) + 
                          Math.random().toString(36).substring(2);

var ip = "192.168.0.104";
var callbackurl = ""; // 'https://a54124be-9d23-4b1d-b43d-de66863b94f5.mock.pstmn.io' // 'https://eowzrjdoaq8tuyh.m.pipedream.net';

// var _file = "empty.docx";
var _file = "empty.xlsx"
// var _file = "empty.pptx";

var _filetype = _file.split('.').pop()

// var url = 'https://testing-documentserver-files.s3.amazonaws.com/public_documents/empty.docx'
var _url = `http://${ip}:7080/files/${_file}`;

// var _documentType = "word"
var _documentType =  "cell"
// var _documentType = "slide"
// var _docuemntType = "pdf"

var _type = "desktop";
// var _type = "mobile";

/////////////////////////////////////////////////////////

function createConnector() {
  window.connector = docEditor.createConnector();
}

/////////////////////////////////////////////////////////

window.docEditor = new DocsAPI.DocEditor("placeholder", {
  type: _type,
  document: {
    fileType: _filetype,
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
    title: _file,
    url: _url,
    permissions: {
      edit: true,
      download: true,
      review: true,
      comment: true,
    },
  },
  documentType: _documentType,
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
