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

// EVENTS

// direct way
var onMetaChange = function (event) {
  console.log(event);
  console.log(event.data.title);
};

// CDE


function getVersion() {
  connector.executeMethod ("GetVersion", [], function (version) {
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
    function (callback_arg) {
      console.log("test:", callback_arg);
    }
  );
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

function getSelectionType() {
  connector.executeMethod("GetSelectionType", [], function (sType) {
    console.log(sType);
  });
}

function callCommand() {

  var recalculate = true

  connector.callCommand(function () {
    var oDocument = Api.GetDocument();
    var oParagraph = Api.CreateParagraph();
    oParagraph.AddText('insert text');
    oDocument.InsertContent([oParagraph], false, { "KeepTextOnly": true });
  }, recalculate)
}

// var html = `<div><p style="background-color:aliceblue;padding:25px;">test</p></div>`
// var html = "<p><b>Plugin methods for OLE objects</b></p><ul><li>AddOleObject</li><li>EditOleObject</li></ul>"
// var html = '<div><br/></div><div style=""><br>RADIOGRAPHIES</div><div><br/></div><div style=""><b><u>Indication</u></b><br>Bilan d\'un traumatisme.</div><div><br/></div><div style=""><b><u>Technique</u></b><br>Face, profil et 3/4</div><div><br/></div><div style=""><b><u>Résultat</u></b><br>Pas de lésion osseuse traumatique. Pas d’anomalie focalisée de la structure osseuse.<br>Bonne congruence articulaire.</div>'
// var html = `<div>RADIOGRAPHIES</div>`
// var html = string_html.replace("div", "test")
function pasteHTML() {
  var html = "<html><body><table style = 'border: 2px solid black;border-collapse: collapse;width: 100%;'> <thead> <tr style = 'border: 1px solid black; padding: 8px;text-align: left;font-weight: bold;'> <th>合同编号</th> <th>合同名称</th> <th>甲方</th> <th>乙方</th> <th>签订日期</th> </tr> </thead> <tbody> <tr> <td>CT2023001</td> <td>软件开发合同</td> <td>A公司</td> <td>B公司</td> <td>2023-01-01</td> </tr> <tr> <td>CT2023002</td> <td>货物采购合同</td> <td>C公司</td> <td>D公司</td> <td>2023-02-15</td> </tr> </tbody> </table></body></html>"
  console.log(html)
  connector.executeMethod("PasteHtml", [html]);
}
// Content Controles

function getAllContentControls() {
  connector.executeMethod("GetAllContentControls", [], (callback_arg) => {
    if (typeof callback_arg[0] != "undefined") {
      for (var i = 0; i < callback_arg.length; i++) {
        console.log(i);
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
      Id: 123,
      Lock: 3,
      Tag: "{TAG}",
      PlaceHolderText: "BlockLvlSdt",
    },
  };

  connector.executeMethod(
    "AddContentControl",
    [config.type, config.property],
    (callback_arg) => {
      // console.log(callback_arg);
    }
  );
}

function addInlineLvlSdt() {
  var config = {
    type: 2, // 1 (block), 2 (inline)
    property: {
      Appearance: 1,
      Id: 321,
      Lock: 3,
      Tag: "{TAG}",
      PlaceHolderText: "InlineLvlSdt",
    },
  };

  connector.executeMethod(
    "AddContentControl",
    [config.type, config.property],
    (callback_arg) => {
      // console.log(callback_arg);
    }
  );
}

function insertAndRemoveCC() {
  var file = "Lorem_Ipsum.docx";

  var oControlPrContent = {
    Props: {
      Id: 1,
      Tag: "text block",
      Lock: 3,
    },
    Url: `http://192.168.4.138:9090/files/bugs/2.docx`,
    Format: "docx",
  };

  const arrDocuments = [oControlPrContent];

  connector.executeMethod(
    "InsertAndReplaceContentControls",
    [arrDocuments],
    (returnValue) => {
      console.log(returnValue);
      // Remove content control
      // connector.executeMethod("RemoveContentControl",
      //                         [returnValue[0].InternalId]
      //                       );
    }
  );
}

function getCurrentAndInsertAndReplaceProps() {
  window.connector.executeMethod(
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
              Color: { R: 100, G: 100, B: 100 },
            },
            Script:
              "var oParagraph = Api.CreateParagraph();oParagraph.AddText('Updated container');Api.GetDocument().InsertContent([oParagraph]);",
          },
        ];

        window.connector.executeMethod("InsertAndReplaceContentControls", [
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
  connector.executeMethod("RemoveContentControl", []);
}

// CSE

function addComment() {
  connector.callCommand(function () {
    Api.AddComment("text", "author");
  });
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
          // var oWorksheet = Api.GetActiveSheet();
          // oWorksheet.GetRange('A1').SetValue(Asc.scope.text)
          var oComments = Api.GetComments();
          var obj = {};
          obj["text"] = oComments[0].GetText();
          obj["AuthorName"] = oComments[0].GetAuthorName();
          return obj;
        },
        function (result) {
          console.log(result);
        },
        true
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
