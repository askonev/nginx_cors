builder.CreateFile("docx");
    let oDoc = Api.GetDocument(0);
    let oPar = oDoc.GetElement(0);
        oPar.AddText('Simple text');
builder.SaveFile("docx", "simple.docx");
builder.CloseFile();