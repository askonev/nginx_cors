
window.host_ip = '192.168.4.138'

var config = function (type) {
  switch (type) {
    case 'docx':
      var _file = "empty.docx";
      // var _file = "docx/with_TOC.docx";
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
    ip: '192.168.4.138',
    uuid:
      Date.now().toString(36) +
      Math.random().toString(36).substring(2).toString(),
    // uuid: 'BCFA2CED',
    source: _file,
    extension: _file.split('.').pop(),
    url: `http://${window.host_ip}:9090/files/${_file}`,
    type: _documentType,
    platform: 'desktop', // "mobile"
    mode: 'edit',
    lang: 'en'
  };
};

config = config('docx');

/////////////////////////////////////////////////////////

// EVENTS

var onAppReady = function () {
  massage = " _              _   _  _ ___  _  _     _                             \n"
  massage += "/ \\ |\\ | | \\_/ / \\ |_ |_  |  /  |_    | \\  _   _     ._ _   _  ._ _|_\n"
  massage += "\\_/ | \\| |_ |  \\_/ |  |  _|_ \\_ |_    |_/ (_) (_ |_| | | | (/_ | | |_\n"
  massage += "|_  _| o _|_  _  ._    o  _    ._ _   _.  _|    |                    \n"
  massage += "|_ (_| |  |_ (_) |     | _>    | (/_ (_| (_| \\/ o                   \n"
  massage += "                                             /                       \n"
  console.log(massage);
};

function createConnector() {
  window.connector = docEditor.createConnector();

  // if (typeof window.connector == 'underfined') {
  //   alert('connector is not defined');
  // }
  var expr = typeof window.connector
  switch(expr) {
    case 'underfined':
      console.error('[ERROR] connector is not defined')
      break;
    case 'object':
      console.log('[LOG] connector exist')
      break;
    default:
      console.log(`Sorry, we are out of ${expr}.`)
  }
}

// direct way
var onMetaChange = function (event) {
  console.log('onMetaChange log:' + event);
  console.log(event.data.title);
};


function onOutdatedVersion() {
  console.log('Event: onOutdatedVersion')
  // location.reload(true);
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
    // callbackUrl: 'http://192.168.4.138:9090',
    user: {
      group: "Group1,Group2",
      id: "78e1e841",
      // image: "https://example.com/url-to-user-avatar.png",
      name: "Smith Johan"
    },
    mode: config.mode,
    lang: config.lang,
    customization: {
      zoom: 120,
      integrationMode: "embed",
    },
  },
  height: "100%",
  width: "100%",
  events: {
    "onAppReady": onAppReady,
    "onDocumentReady": createConnector,
    "onMetaChange": onMetaChange,
    "onOutdatedVersion": onOutdatedVersion
  },
});
