builderJS.OpenFile("http://192.168.0.105:7080/files/comparer/basic_agreement.docx");
    var file = builderJS.OpenTmpFile("http://192.168.0.105:7080/files/fake");
    // var file = builderJS.OpenTmpFile("http://192.168.0.105:7080/files/comparer/small-docx-file.docx");
        AscCommonWord.CompareDocuments(Api, file, null);
        file.Close();
builderJS.SaveFile("docx", "Result.docx");
builderJS.CloseFile();
