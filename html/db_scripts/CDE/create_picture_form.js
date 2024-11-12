builder.CreateFile("docx");
    var oDocument = Api.GetDocument();
    var oPictureForm = Api.CreatePictureForm({"tip": "Upload your photo", "required": true, "placeholder": "Photo", "scaleFlag": "tooBig", "lockAspectRatio": true, "respectBorders": false, "shiftX": 50, "shiftY": 50});
        oPictureForm.SetImage(native.GetImageUrl("/confidential/personal/personal_image.png"));
    var oParagraph = oDocument.GetElement(0);
        oParagraph.AddElement(oPictureForm);
builder.SaveFile("docx", "GetImage.docx");
builder.CloseFile();
