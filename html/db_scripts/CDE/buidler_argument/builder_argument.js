
// --argument=\{\"filename\"\:\"arg_simple.docx\"\,\"key\"\:\"onlyoffice\"\,\"key2\"\:\"docbuilder\"\}

builder.CreateFile("docx");
    let oDoc = Api.GetDocument(0);
    let oPar = oDoc.GetElement(0);
        oPar.AddText(Argument["key"]);
        oPar.AddLineBreak();
        oPar.AddText(Argument["key2"]);
    let fileName = Argument["filename"]
builder.SaveFile("docx", "jsValue(fileName)");
builder.CloseFile();
