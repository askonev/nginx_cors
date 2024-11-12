builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oTable = Api.CreateTable(3, 3);
var oCell = oTable.GetRow(0).GetCell(0);
    oCell.SetWidth("twips", 4200);
    oCell.GetContent().GetElement(0).AddText("text2");
    oCell = oTable.GetRow(0).GetCell(1);
    oCell.SetWidth("twips", 1200);
    oCell = oTable.GetRow(0).GetCell(2);
    oCell.GetContent().GetElement(0).AddText("text3");
var oTextPr_1 = Api.CreateTextPr();
    oTextPr_1.SetFontSize(14 * 2);
    // Set font size 14 to Row(0)
    oTable.GetRow(0).SetTextPr(oTextPr_1);
    oCell = oTable.GetRow(2).GetCell(0);
    oCell.GetContent().GetElement(0).AddText("text4");
var oTextPr_2 = Api.CreateTextPr();
    oTextPr_2.SetFontSize(10 * 2);
    // Set font size 10 to Row(2)
    oTable.GetRow(2).SetTextPr(oTextPr_2);
    oTable.SetWidth("percent", 100);
    oDocument.Push(oTable);
builder.SaveFile("docx", "script_sup.docx");
builder.CloseFile();