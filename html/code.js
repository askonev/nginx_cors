// function TrackRevisionON() {
//     connector.callCommand(function () {
//         var odoc = Api.GetDocument();
//         odoc.SetTrackRevisions(true);
//         console.log('on track')
//     })
// }

// function TrackRevisionOFF() {
//     connector.callCommand(function () {
//         var odoc = Api.GetDocument();
//         odoc.SetTrackRevisions(false);
//         console.log('off track')
//     })
// }


// docEditor

function insertImage() {
  window.docEditor.insertImage({
    "c": "add",
    "fileType": "png",
    "url": `http://${window.host_ip}:9090/images/dev.png`
  });
}

// CDE

function getVersion() {
  connector.executeMethod("GetVersion", [], function (version) {
    console.log(version);
  });
}

function addHello() {
  connector.callCommand(
    function () {
      var oDocument = Api.GetDocument();
      var oParagraph = Api.CreateParagraph();
      oParagraph.AddText("Hello");
      oDocument.InsertContent([oParagraph]);
      Api.AddComment(oParagraph, "text", "author");
      return "text & comment added";
    },
    function (callback_arg) { console.log("test:", callback_arg); },
    false
  );
}

function addImage() {
  connector.callCommand(function () {
    var oDocument = Api.GetDocument()
    var oParagraph = oDocument.GetElement(0)
    var oDrawing = Api.CreateImage("https://legacy-api.onlyoffice.com/content/img/docbuilder/examples/coordinate_aspects.png",
      60 * 36000,
      35 * 36000)
    oParagraph.AddDrawing(oDrawing)
  });
}

function GetRewiewReport() {
  connector.callCommand(function () {
    var odoc = Api.GetDocument();
    // odoc.SetTrackRevisions(true);
    var report = odoc.GetReviewReport();
    var opar = Api.CreateParagraph();
    if (typeof report["Anonymous"] === "object") {
      for (var i = 0; i < report["Anonymous"].length; i++) {
        var change_info = report["Anonymous"][i];
        console.log(change_info);
      }
      opar.AddText("Anonymous: " + report["Anonymous"][0]);
      odoc.Push(opar);
      // console.log("Anonymous: " + report["Anonymous"]);
    } else {
      console.log("there are no changes");
    }
    // odoc.SetTrackRevisions(false);
  });
}

function pullMetaChange() {
  // Define the URL and request data
  const url = "http://192.168.0.153:3000/coauthoring/CommandService.ashx";
  const data = {
    c: "meta",
    key: uniqueId,
    meta: {
      title: "New title",
    },
  };

  // Create the request options
  const requestOptions = {
    method: "POST",
    mode: "no-cors", // Set the mode to 'no-cors'
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(data),
  };

  fetch(url, requestOptions)
    .then((response) => {
      if (!response.ok) {
        throw new Error("Network response was not ok");
      }
      return response.json(); // Parse the response as JSON if needed
    })
    .then((data) => {
      console.log(data); // Handle the response data here
    })
    .catch((error) => {
      console.error("Error:", error);
    });
}

function getSelectionType() {
  connector.executeMethod("GetSelectionType", [], function (sType) {
    console.log(sType);
  });
}

function callCommand() {

  var bRecalculate = true

  connector.callCommand(function () {
    var oDocument = Api.GetDocument();
    var oParagraph = Api.CreateParagraph();
    oParagraph.AddText('insert text');
    oDocument.InsertContent([oParagraph], false, { "KeepTextOnly": true });
  },
    null,
    bRecalculate)
}

// var html = `<div><p style="background-color:aliceblue;padding:25px;">test</p></div>`
// var html = "<p><b>Plugin methods for OLE objects</b></p><ul><li>AddOleObject</li><li>EditOleObject</li></ul>"
// var html = '<div><br/></div><div style=""><br>RADIOGRAPHIES</div><div><br/></div><div style=""><b><u>Indication</u></b><br>Bilan d\'un traumatisme.</div><div><br/></div><div style=""><b><u>Technique</u></b><br>Face, profil et 3/4</div><div><br/></div><div style=""><b><u>Résultat</u></b><br>Pas de lésion osseuse traumatique. Pas d’anomalie focalisée de la structure osseuse.<br>Bonne congruence articulaire.</div>'
// var html = `<div>RADIOGRAPHIES</div>`
// var html = string_html.replace("div", "test")
// var html = "<html><body><table style = 'border: 2px solid black;border-collapse: collapse;width: 100%;'> <thead> <tr style = 'border: 1px solid black; padding: 8px;text-align: left;font-weight: bold;'> <th>合同编号</th> <th>合同名称</th> <th>甲方</th> <th>乙方</th> <th>签订日期</th> </tr> </thead> <tbody> <tr> <td>CT2023001</td> <td>软件开发合同</td> <td>A公司</td> <td>B公司</td> <td>2023-01-01</td> </tr> <tr> <td>CT2023002</td> <td>货物采购合同</td> <td>C公司</td> <td>D公司</td> <td>2023-02-15</td> </tr> </tbody> </table></body></html>"
function pasteHTML() {
  var html = "<p style='text-align: left;'>text</p>"
  // var html = "<p style='text-align: center;'>text</p>"
  console.log(html)
  connector.executeMethod("PasteHtml", [html], null);
}

function searchNext() {
  connector.executeMethod('SearchNext', [
    {
      searchString: 'Hello',
      matchCase: false,
    },
    true
  ],
    null)
}

function moveToNextReviewChange() {
  connector.executeMethod("MoveToNextReviewChange", [false]);
}

function insertOleObject() {
  var bSelect = true	// Defines if the OLE object will be selected after inserting into the document (true) or not (false).
  connector.executeMethod("InsertOleObject", [
    {
      "Data": "{data}",
      "ImageData": "data:image/png;base64,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",
      "ApplicationId": "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}",
      "Width": 70,
      "Height": 70,
      "WidthPix": 60 * 36000,
      "HeightPix": 60 * 36000
    },
    bSelect]);
}

function getCurrentSentence() {
  connector.executeMethod("GetCurrentSentence", ["entirely"], function (res) {
    console.log(res)
  });
}

function moveToNextReviewChange() {
  connector.executeMethod("MoveToNextReviewChange", [true]);
}

function moveToPreviousReviewChange() {
  connector.executeMethod("MoveToNextReviewChange", [false]);
}

// Comments

function addCommentViaElement() {
  connector.callCommand(function () {
    var oDocument = Api.GetDocument();
    var oParagraph = oDocument.GetElement(0);
    oParagraph.AddText("This is just a sample text");
    Api.AddComment(oParagraph, "comment", "Makoto Senpai");
  });
}

function setUserId() {
  connector.callCommand(function () {
    var oDocument = Api.GetDocument();
    var aComments = oDocument.GetAllComments();

    aComments.forEach(element => {
      element.SetUserId('uid-2');
      element.SetAuthorName('Popato Markes')
    });
  });
}

// Content Controles

function addBlockLvlSdt() {
  // console.log(uniqueId)

  // 0 - only deleting
  // 1 - disable deleting or editing
  // 2 - only editing
  // 3 - full access

  var config = {
    type: 1, //  1 (block), 2 (inline)
    property: {
      Appearance: 1,
      Id: 1001,
      Lock: 3,
      Tag: "BSDT",
      PlaceHolderText: "BlockLvlSdt",
    },
  };

  connector.executeMethod(
    "AddContentControl",
    [config.type, config.property],
    (callback_arg) => {
      console.log(callback_arg);
    }
  );
}

function addInlineLvlSdt() {
  var config = {
    type: 2, // 1 (block), 2 (inline)
    property: {
      Appearance: 1,
      Id: 1002,
      Lock: 3,
      Tag: "ISDT",
      PlaceHolderText: "InlineLvlSdt",
    },
  };

  connector.executeMethod(
    "AddContentControl",
    [config.type, config.property],
    (callback_arg) => {
      console.log(callback_arg);
    }
  );
}

function getAllContentControls() {
  connector.executeMethod("GetAllContentControls", [], (callback_arg) => {
    if (typeof callback_arg[0] != "undefined") {
      for (var i = 0; i < callback_arg.length; i++) {
        console.log(i, callback_arg[i])
      }
    }
  });
}

function getCurrentContentControlPr() {
  connector.executeMethod("GetCurrentContentControlPr", [], (callback) => {
    console.log(callback);

    // var arrDocuments = [
    //   {
    //     Props: {
    //       Id: 100,
    //       InternalId: callback.InternalId,
    //       Tag: "CC_Tag",
    //       Lock: 3,
    //       PlaceHolderText: "custom",
    //     },
    //     Script:
    //       "var oParagraph = Api.CreateParagraph();oParagraph.AddText('Hello world!');Api.GetDocument().InsertContent([oParagraph]);",
    //   },
    // ];
    // connector.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
  });
}

function setPlaseHolder() {
  connector.executeMethod("GetCurrentContentControlPr", [], (callback) => {
    console.log(callback);

    var arrDocuments = [
      {
        Props: {
          InternalId: callback.InternalId,
          PlaceHolderText: "CUSTOM",
        },
      },
    ];
    connector.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
  });
}

function insertAndRemoveCC() {
  var file = 'Lorem_Ipsum.docx';
  var dir = 'template';

  var oControlPrContent = {
    Props: {
      Id: 1,
      Tag: "text block",
      Lock: 3,
    },
    Url: `http://192.168.4.142:9090/files/${dir}/${file}`,
    Format: "docx",
  };

  const arrDocuments = [oControlPrContent];

  connector.executeMethod(
    "InsertAndReplaceContentControls",
    [arrDocuments],
    (returnValue) => {
      // console.log(returnValue);
      // Remove content control
      // connector.executeMethod("RemoveContentControl", [returnValue[0].InternalId]);
    }
  );
}

function insertAndReplaceProps() {
  connector.executeMethod(
    "GetCurrentContentControl",
    [],
    function (InternalId) {
      if (InternalId) {
        var arrDocuments = [
          {
            Props: {
              InternalId: InternalId,
              Id: 100,
              Tag: "Tag",
              Lock: 1,
              Alias: "alias",
              PlaceHolderText: "custom_placeholder",
              Appearance: 1,
              Color: { R: 255, G: 129, B: 44 },
            },
            Script: "var oParagraph = Api.CreateParagraph();\n" +
              "oParagraph.AddText('Updated container');\n" +
              "Api.GetDocument().InsertContent([oParagraph], false);\n",
          },
        ];
        connector.executeMethod("InsertAndReplaceContentControls", [
          arrDocuments,
        ]);
      } else {
        console.log("Please select CC");
      }
    }
  );
}

function attachOnChangeContentControl() {
  connector.attachEvent("onChangeContentControl", function () {
    console.log("event: onChangeContentControl");
  });
}

function remove_cc() {

  // connector.executeMethod("GetAllContentControls", null, (ccs) => {
  //   console.log(ccs)

  //   cc = ccs[0]

  //   connector.executeMethod('RemoveContentControl', [cc.InternalId])

  // });

  connector.executeMethod("RemoveContentControl", []);
}

function startAction() {
  function setPasswordByFile(type, value) {
    console.log(type, value)
  }

  // var flag = 'Information'
  var flag = 'Block'

  connector.executeMethod("StartAction", [`${flag}`, "Message 1"], function () {
    setPasswordByFile("sha256", "123456");

    setTimeout(function () {
      connector.executeMethod("EndAction", [`${flag}`, "Message 1"]);
      console.log("End Action")
    }, 2000);
  });
}

function blockUnblockAllCC() {
  connector.callCommand(
    function () {
      let oDocument = Api.GetDocument();
      let ctls = oDocument.GetAllContentControls();
      for (let i = 0; i < ctls.length; i++) {
        let lt = ctls[i].GetLock();
        console.log(i, lt)
        if (lt == 'contentLocked' || lt == 'sdtContentLocked' || lt == 'sdtLocked') {
          ctls[i].SetLock("");
        } else {
          // "contentLocked" | "sdtContentLocked" | "sdtLocked"
          ctls[i].SetLock("sdtContentLocked");
        }
        console.log(ctls[i].GetLock());
      }
    },
    null,
    false
  );
}

// Events

function onEncryption(event) {
  console.log(event);
  connector.executeMethod("OnEncryption", [
    {
      type: "generatePassword",
      password: "123456",
      docinfo: "{docinfo}",
    },
  ]);
}

function onBlureCC() {
  connector.attachEvent('onBlurContentControl', function (oPr) {
    if (oPr) {
      console.log(oPr)
    }
    console.log('event: onBlurContentControl')
  })
}

function onFocusCC() {
  connector.attachEvent('onFocusContentControl', function (oPr) {
    if (oPr) {
      console.log(oPr)
    }
    console.log('event: onFocusContentControl')
  })
}

function onChangeContentControl() {
  connector.attachEvent('onChangeContentControl', function (control) {
    console.log("event: onChangeContentControl", JSON.stringify(control));
  });
}

function dettach_onChangeContentControl() {
  connector.detachEvent("onChangeContentControl");
}

function onClick() {
  connector.attachEvent("onClick", function (isSelectionUse) {
    console.log(`event: ${isSelectionUse}`);
    // var selectedOleObject = connector.executeMethod("GetSelectedOleObjects");

    // debugger
  });
}

function onContextMenuClick() {
  connector.attachEvent("onContextMenuClick", function (id) {
    switch (id) {
      case 'onConvert': {
        console.log('onConvert')
      }
      case 'onChat': {
        console.log('onChat')
      }

    }
  });
}

function onContextMenuShow() {
  var settings = {
    guid: connector.guid,
    items: [
      {
        id: 'onConvert',
        text: 'context item',
        disabled: false,
        separator: true
      },
      {
        id: 'onChat',
        text: 'chat item',
        disabled: false,
        separator: true
      },
    ]
  }
  // debugger
  connector.attachEvent('onContextMenuShow', function (options) {
    // console.log('[onContextMenuShow]', options)
    if (!options) return;
    if (options.type === 'Selection' || options.type === 'Target') {
      this.executeMethod('AddContextMenuItem', [settings]);
      // console.log('onContextMenuShow')
      // connector.executeMethod("InputText", ["clicked: onContextMenuShow"]);
    }
  });
}

// CSE

function addComment() {
  connector.callCommand(function () {
    Api.AddComment("text", "author");
  });
}

function setAscScope() {
  Asc.scope.data = "Hello world!";
  connector.callCommand(function () {
    var oWorksheet = Api.GetActiveSheet();
    oWorksheet.GetRange("B1").SetValue(Asc.scope.data);
    var data = oWorksheet.GetRange("B1").GetValue();
    return data
  },
    (data) => { console.log(data) }),
    isNoCalc = true
}

// CPE

function createSlide() {
  connector.callCommand(
    function () {
      var oPresentation = Api.GetPresentation();
      var oSlide = Api.CreateSlide();
      var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
      var oGs2 = Api.CreateGradientStop(
        Api.CreateRGBColor(255, 111, 61),
        100000
      );
      var oFill = Api.CreateRadialGradientFill([oGs1, oGs2]);
      oSlide.SetBackground(oFill);
      oPresentation.AddSlide(oSlide);
    },
    function () {
      console.log("callback command");
    }
  );
}

// Common

function getAllComments() {
  switch (window.docType) {
    case "word":
      console.log("GetAllComments");
      connector.executeMethod("GetAllComments", [], (callback_arg) => {
        console.log(callback_arg);
      });
      break;

    case "cell":
      console.log("GetAllComments");
      connector.callCommand(
        function () {
          // var oComments = Api.GetComments();
          var oComments = Api.GetAllComments();
          var obj = {};
          if (oComments.length > 0) {
            for (var i = 0; i < oComments.length; i++) {
              obj[`${i}`] = {
                "Text": `${oComments[i].GetText()}`,
                "Id": `${oComments[i].GetId()}`,
                "AuthorName": `${oComments[i].GetAuthorName()}`,
                "UserId": `${oComments[i].GetUserId()}`,
              }
            }
          } else {
            obj["text"] = 'no comments'
          }
          return obj;
        },
        function (result) {
          console.log(result);
        },
        null,
        isNoCalc = true
      );
      break;

    case "slide":
      console.log("GetAllComments");
      break;
  }
}


function getSelectedText() {

  var numbering = {
    "NewLine": true,
    "NewLineParagraph": true,
    "Numbering": true,
    "Math": false,
    "TableCellSeparator": ';',
    "TableRowSeparator": '_',
    "ParaSeparator": '\n',
    "TabSymbol": String.fromCharCode(9)
  }
  connector.executeMethod("GetSelectedText", [numbering], function (data) {
    console.log(data)
  });
}
