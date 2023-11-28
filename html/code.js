// function TrackRevisionON() {
//     connector.callCommand(function () {
//         let odoc = Api.GetDocument();
//         odoc.SetTrackRevisions(true);
//         console.log('on track')
//     })
// }

// function TrackRevisionOFF() {
//     connector.callCommand(function () {
//         let odoc = Api.GetDocument();
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

// through connector
function attachOnChangeContentControl() {
  connector.attachEvent("onChangeContentControl", function () {
    console.log("event: onChangeContentControl");
  });
}

// METHODS

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

function GetRewiewReport() {
  connector.callCommand(function () {
    var odoc = Api.GetDocument();
    // odoc.SetTrackRevisions(true);
    var report = odoc.GetReviewReport();
    var opar = Api.CreateParagraph();
    if (typeof report["Anonymous"] === "object") {
      for (let i = 0; i < report["Anonymous"].length; i++) {
        let change_info = report["Anonymous"][i];
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

function insertAndReplaceProps() {
  window.connector.executeMethod(
    "GetAllContentControls",
    null,
    function (cc_list) {
      console.log("content control list", cc_list);

      var sIternalId = cc_list[0].InternalId.toString(); // first Content Control

      var arrDocuments = [
        {
          Props: {
            InternalId: sIternalId,
            Id: 100,
            Tag: "Tag",
            Lock: 3,
            Alias: "alias",
            PlaceHolderText: "custom_placeholder",
            Appearance: 1,
            Color: { R: 100, G: 100, B: 100 },
          },
          // Script:
          //   "var oParagraph = Api.CreateParagraph();oParagraph.AddText('Updated container');Api.GetDocument().InsertContent([oParagraph]);",
        },
      ];

      window.connector.executeMethod("InsertAndReplaceContentControls", [
        arrDocuments,
      ]);
    }
  );
}

function getAllComments() {
  console.log("GetAllComments");
  connector.executeMethod("GetAllComments", [], (callback_arg) => {
    console.log(callback_arg);
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

function addBlockLvlSdt() {
  // console.log(uniqueId)

  let config = {
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
  let config = {
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
