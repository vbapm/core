// Recreate multilingual fixture .bas files with correct encodings
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "../tests/__fixtures__/multilingual");

const fixtures = [
  {
    dir: "cp1252",
    encoding: "win1252",
    content:
      'Attribute VB_Name = "Bonjour"\r\n' +
      "Option Explicit\r\n" +
      "\r\n" +
      "' Français : é è ê ë à â ä ù û ü ç É È Ê Ë À Â Ä Ù Û Ü Ç\r\n" +
      "' Allemand : ä ö ü ß Ä Ö Ü\r\n" +
      "' Espagnol : á é í ó ú ñ ¿ ¡\r\n" +
      "\r\n" +
      "Public Function Hello() As String\r\n" +
      '    Hello = "Voilà les caractères accentués !"\r\n' +
      "End Function\r\n",
  },
  {
    dir: "cp1251",
    encoding: "win1251",
    content:
      'Attribute VB_Name = "Privet"\r\n' +
      "Option Explicit\r\n" +
      "\r\n" +
      "' Русский : Привет мир! Как дела?\r\n" +
      "' А Б В Г Д Е Ё Ж З И Й К Л М Н О П Р С Т У Ф Х Ц Ч Ш Щ Ъ Ы Ь Э Ю Я\r\n" +
      "' а б в г д е ё ж з и й к л м н о п р с т у ф х ц ч ш щ ъ ы ь э ю я\r\n" +
      "\r\n" +
      "Public Function Hello() As String\r\n" +
      '    Hello = "Привет мир!"\r\n' +
      "End Function\r\n",
  },
  {
    dir: "cp1250",
    encoding: "win1250",
    content:
      'Attribute VB_Name = "Witaj"\r\n' +
      "Option Explicit\r\n" +
      "\r\n" +
      "' Polski : ą ć ę ł ń ó ś ź ż Ą Ć Ę Ł Ń Ó Ś Ź Ż\r\n" +
      "' Česky : ě š č ř ž ý á í é Ě Š Č Ř Ž Ý Á Í É\r\n" +
      "\r\n" +
      "Public Function Hello() As String\r\n" +
      '    Hello = "Witaj świecie! ą ć ę ł ń ó ś ź ż"\r\n' +
      "End Function\r\n",
  },
];

let iconv;
try {
  iconv = require("iconv-lite");
} catch {
  console.log("iconv-lite not available, using latin1 fallback (CP1252 only)");
}

for (const { dir, encoding, content } of fixtures) {
  let buf;
  if (iconv) {
    buf = iconv.encode(content, encoding);
  } else {
    // Fallback: latin1 works for CP1252 but not others
    buf = Buffer.from(content, "latin1");
  }
  const filePath = path.join(root, dir, "src", "Hello.bas");
  fs.writeFileSync(filePath, buf);
  console.log(`${dir}: ${buf.length} bytes (${encoding})`);
}

console.log("Done");
