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

function insertAndReplace() {
  window.connector.executeMethod(
    "GetAllContentControls",
    null,
    function (cc_list) {
      console.log(cc_list);

      var sIternalId = cc_list[0].InternalId.toString(); // first LvlSdt

      var arrDocuments = [
        {
          Props: {
            InternalId: sIternalId,
            Id: 100,
            Tag: "CC_Tag",
            Lock: 3,
          },
          Script:
            "var oParagraph = Api.CreateParagraph();oParagraph.AddText('Hello world!');Api.GetDocument().InsertContent([oParagraph]);",
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

function attachOnChangeContentControl() {
  connector.attachEvent("onChangeContentControl", function () {
    console.log("event: onChangeContentControl");
  });
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
