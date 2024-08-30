builderJS.OpenFile("http://192.168.1.220:9090/files/empty.docx");
    var file = builderJS.OpenTmpFile("http://192.168.1.220:9090/files/docx/docx_with_jpg.docx");
        AscCommonWord.CompareDocuments(Api, file, null);
        file.Close();
builderJS.SaveFile("docx", "Result.docx");
builderJS.CloseFile();
