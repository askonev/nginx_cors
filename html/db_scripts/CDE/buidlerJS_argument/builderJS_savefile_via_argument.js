
// documentbuilder --argument=\{\"file\"\:\"custom_name.docx\"\}

builderJS.CreateFile("docx");
    var sFile = Argument["file"];
console.log('Arg:')
console.log(sFile)
builderJS.SaveFile("docx", sFile)
builderJS.CloseFile();
