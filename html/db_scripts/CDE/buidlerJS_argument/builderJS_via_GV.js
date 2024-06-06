
// db --argument=\{\"file\"\:\"/home/askonew/Downloads/chrome.tmp/new.docx\"\,\"name\"\:\"new.docx\"\}

builderJS.CreateFile("docx");
    var sFile = Argument["file"];
    // console.log('Arg:')
    // console.log(sFile)
    GlobalVariable["sFile"] = sFile;
    // console.log('GV:')
    // console.log(GlobalVariable["sFile"]);
builderJS.CloseFile();

// builderJS.OpenFile(GlobalVariable["sFile"]);
builderJS.CreateFile("docx")
    console.log('Result:')
    console.log(GlobalVariable["sFile"]);
// builderJS.SaveFile("docx", "new.docx")
builderJS.CloseFile();
