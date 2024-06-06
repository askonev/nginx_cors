
// db --argument=\{\"file\"\:\"/home/askonew/Downloads/chrome.tmp/new.docx\"\,\"name\"\:\"new.docx\"\}
// db --argument=\{\"format\"\:\"docx\"\}


// * incorrect
// var sFile = Argument["file"]
// var sName = Argument["name"]
// builderJS.OpenFile(sFile, sName);

// * incorrect
// var format = Argument["format"]
// builderJS.CreateFile(format);

// builderJS.OpenFile("/home/askonew/Downloads/chrome.tmp/new.docx", "new.docx");
// builderJS.CreateFile(AVS_OFFICESTUDIO_FILE_DOCUMENT_DOCX);
builderJS.CreateFile("docx");
    let doc = Api.GetDocument();
    let par = doc.GetElement(0)
        par.AddText("test text")
        var sFile = Argument["file"]
        var sName = Argument["name"]
        par.AddText(sFile)
        par.AddText(sName)
    console.log(doc.GetElement(0).GetText())

// builderJS.SetTmpFolder("folder");
// builderJS.SaveFile(AVS_OFFICESTUDIO_FILE_DOCUMENT_DOCX, "path", "x2t_additons_as_xml");
// builderJS.SaveFile("docx", "some.docx");
builderJS.CloseFile();
