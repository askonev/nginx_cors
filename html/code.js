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

// Content Controles

function getAllContentControls() {
  connector.executeMethod("GetAllContentControls", [], (callback_arg) => {
    if (typeof callback_arg[0] != "undefined") {
      // console.log(callback_arg[0].Tag);
      for (var i = 0; i < callback_arg.length; i++) {
        // console.log(i)
        console.log(callback_arg[i].Tag);
      }
    }
  });
}

function getCurrentContentControl() {
  window.connector.executeMethod("GetCurrentContentControlPr", [], (sdt) => {
    console.log(sdt);
    console.log(sdt.Tag);

    // debugger

    // window.Asc.scope = {
    //   ccTag : sdt.Tag
    // };

    // console.log(sdt)

    connector.callCommand(
      function () {
        console.log("call callCommand()");
        // var oDocument = Api.GetDocument();
        // console.log(oDocument)
        // var aContentControls = oDocument.GetAllContentControls();
        // console.log(aContentControls)
        // // var aContentControls = Api.pluginMethod_GetAllContentControls();

        // for (var oContentControl of aContentControls) {
        //   if (oContentControl.GetTag() === Asc.scope.ccTag) {
        //     var oRange = oContentControl.GetRange();
        //     var text = oRange.GetText();
        //     console.log(text)
        //     return text;
        //     }
        //   }
      },
      false,
      true,
      function (text) {
        // console.log(text)
      }
    );
  });
}

function addBlockLvlSdt() {
  // console.log(uniqueId)

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
  var file = "shape.docx";

  var oControlPrContent = {
    Props: {
      Id: 1,
      Tag: "text block",
      Lock: 3,
    },
    Url: `http://192.168.4.138:7080/files/template/${file}`,
    Format: "docx",
  };

  const arrDocuments = [oControlPrContent];

  connector.executeMethod(
    "InsertAndReplaceContentControls",
    [arrDocuments],
    (returnValue) => {
      console.log(returnValue);
      // Remove content control
      connector.executeMethod("RemoveContentControl", [
        returnValue[0].InternalId,
      ]);
    }
  );
}

function insertAndReplaceProps() {
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

function remove() {
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

// Universal

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
