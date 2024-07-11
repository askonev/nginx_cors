
window.host_ip = '192.168.0.103'

var config = function(type) {
  switch (type) {
    case 'docx':
      var _file = "empty.docx";
      // 'https://testing-documentserver-files.s3.amazonaws.com/public_documents/empty.docx'
      var _documentType = "word";
      break;
    case 'xlsx':
      var _file = "empty.xlsx";
      var _documentType = "cell";
      break;
    case 'pptx':
      var _file = "empty.pptx";
      var _documentType = "slide";
      break;
    case 'pdf':
      var _documentType = "pdf";
      var _file = "sample.pdf"
      break;
  }

  return {
    ip: '192.168.0.103',
    uuid:
      Date.now().toString(36) +
      Math.random().toString(36).substring(2).toString(),
    source: _file,
    extension: _file.split('.').pop(),
    url: `http://${window.host_ip}:9090/files/${_file}`,
    type: _documentType,
    platform: 'desktop', // "mobile"
    mode: 'edit',
  };
};

config = config('docx');

/////////////////////////////////////////////////////////

function createConnector() {
  window.connector = docEditor.createConnector();
}

window.docType = config.type;

/////////////////////////////////////////////////////////

window.docEditor = new DocsAPI.DocEditor("placeholder", {
  type: config.platform,
  document: {
    fileType: config.extension,
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
    key: config.uuid,
    title: config.source,
    url: config.url,
    permissions: {
      edit: true,
      download: true,
      review: true,
      comment: true,
    },
  },
  documentType: config.type,
  editorConfig: {
    mode: config.mode,
    customization: {
      zoom: 100,
      integrationMode: "embed",
    },
    callbackUrl: '',
  },
  height: "100%",
  width: "100%",
  events: {
    onDocumentReady: createConnector,
    onMetaChange: onMetaChange,
  },
});
