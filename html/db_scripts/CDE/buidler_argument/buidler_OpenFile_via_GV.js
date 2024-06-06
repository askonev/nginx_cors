// db --argument=\{\"key\"\:\"http\://192.168.0.153\:7080/files/empty.docx\"\} buidler_OpenFile_via_GV.js
// db --argument=\{\"address\"\:\"arg_simple.docx\"\} buidler_OpenFile_via_GV.js

builder.CreateFile("docx");
    var address = Argument["address"];
    // console.log(address)
    GlobalVariable["data"] = address;
    console.log(GlobalVariable["data"]);
builder.CloseFile();

// incorrect 
// builder.OpenFile(GlobalVariable["data"]);
builder.OpenFile("arg_simple.docx");
    console.log(GlobalVariable["data"]);
builder.SaveFile("docx", "new.docx")
builder.CloseFile();
