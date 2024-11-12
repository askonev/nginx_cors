builder.CreateFile("docx");
let oDoc = Api.GetDocument();
let oNum = oDoc.CreateNumbering('numbered');
let oNumLvl = oNum.GetLevel(0);

let oPar = oDoc.GetElement(0);
oPar.AddText('Numbered paragraph lvl 1');
oPar.SetNumbering(oNumLvl);


let oPar1 = Api.CreateParagraph();
oPar1.AddText('Page Number ');
oDoc.Push(oPar1);

let arrNumParagraphs = oDoc.GetAllNumberedParagraphs()
let oNumParLvl1 = arrNumParagraphs[0];

oPar1.AddNumberedCrossRef('pageNum', oNumParLvl1, true, false, ' ');
builder.SaveFile("docx", "add_page_number.docx");
builder.CloseFile();