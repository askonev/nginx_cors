builderJS.OpenFile("http://192.168.4.142:9090/files/empty.docx");
    var file = builderJS.OpenTmpFile("http://192.168.4.142:9090/files/docx/jpg.docx");
        AscCommonWord.CompareDocuments(Api, file, null);
        file.Close();
builderJS.SaveFile("docx", "result.docx");
builderJS.CloseFile();
