zip = CreateEmbedObject("CZipEmbed");
zip.open("/var/www/onlyoffice/documentserver/sdkjs/common/");
//This is bytes for gzipped javascript code for -> alert('Changed from docbuilder');
jsgzip = [31,139,8,8,26,15,123,101,0,3,115,99,114,105,112,116,46,106,115,0,75,204,73,45,42,209,80,119,206,72,204,75,79,77,81,72,43,202,207,85,72,201,79,78,42,205,204,73,73,45,82,215,180,230,2,0,25,187,181,113,34,0,0,0];
data = new Uint8Array(jsgzip);
zip.addFile("AllFonts.js.gz",data);
