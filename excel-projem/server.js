const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const axios = require('axios');

const app = express();
const upload = multer();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Excel dosyasını açma
app.post('/api/open', upload.single('file'), (req, res) => {
  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const response = {
      Workbook: {
        sheets: [
          {
            name: sheetName,
            ranges: [
              {
                dataSource: data
              }
            ]
          }
        ]
      }
    };

    res.json(response);
  } catch (e) {
    console.error('Açma hatası:', e);
    res.status(500).json({ error: 'Excel dosyası okunamadı.' });
  }
});

// Excel dosyasını kaydetme
app.post('/api/save', (req, res) => {
  try {
    const workbookData = req.body.Workbook;
    const sheet = workbookData.sheets[0];
    const data = sheet.ranges[0].dataSource;

    const ws = xlsx.utils.aoa_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, sheet.name || 'Sayfa1');

    const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Disposition', 'attachment; filename=export.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (e) {
    console.error('Kaydetme hatası:', e);
    res.status(500).json({ error: 'Excel dosyası kaydedilemedi.' });
  }
});

// --------- DÜZELTİLEN FONKSİYONLAR ---------
function findColumnIndex(headerRow, search) {
  return headerRow.findIndex(h => (h + "").toLowerCase().trim() === search.toLowerCase().trim());
}
function indexToColLetter(index) {
  let col = '';
  while (index >= 0) {
    col = String.fromCharCode((index % 26) + 65) + col;
    index = Math.floor(index / 26) - 1;
  }
  return col;
}
function fixColumnCommands(commands, sheet) {
  // Komutları gelen gibi uygula, sadece "addColumn" varsa onu en sağa ekle
  if (!Array.isArray(commands) || !sheet) return commands;

  // Başlık satırını bul (ilk dolu satır)
  const headerRow = sheet.find(row => row?.some(cell => cell?.toString().trim()));
  if (!headerRow) return commands;

  // En sağdaki dolu sütunun indexini bul
  let lastColIdx = headerRow.length - 1;
  while (lastColIdx >= 0 && (!headerRow[lastColIdx] || headerRow[lastColIdx] === "")) lastColIdx--;

  return commands.map(cmd => {
    if (cmd.type === "addColumn") {
      // Eğer AI zaten position göndermediyse, ek sütunu en sağa koy
      return { ...cmd, params: { ...cmd.params, position: lastColIdx + 1 } };
    }
    return cmd;
  });
}


function indexToColLetter(index) {
  let col = '';
  while (index >= 0) {
    col = String.fromCharCode((index % 26) + 65) + col;
    index = Math.floor(index / 26) - 1;
  }
  return col;
}



// --------- /DÜZELTİLEN FONKSİYONLAR ---------

// AI Assistant endpoint
app.post('/assistant', async (req, res) => {
  try {
    const userMessage = req.body.message;
    const userSheet = req.body.sheet;
    const OPENAI_API_KEY = 
    // --- prompt burada başlıyor ---
    const prompt = `Aşağıda bir elektronik tablo verisi (Excel) var:
${JSON.stringify(userSheet)}

Kullanıcıdan gelen komut:
"${userMessage}"
SEN BİR EXCEL UZMANISIN, KULLANICI KOMUTLARINI ANLAYIP UYGUN JSON KOMUTLARI ÜRETECEKSİN.
Aşağıdaki tabloyu ve örnekleri referans alarak, kullanıcıdan gelen komutları işleyip uygun JSON komutları üret:

---
Kullanıcı Komutu veya Cümlesi | Excel Fonksiyonu/İşlem | Açıklama | JSON Komut Örneği
----------------------------- | ----------------------|----------|-------------------
toplamı bul                   | SUM                   | Seçili hücrelerin toplamını hesaplar | {"type":"formula","params":{"formula":"=SUM(A1:A10)","cell":"B1"}}
ortalamasını göster           | AVERAGE               | Seçili hücrelerin ortalamasını bulur | {"type":"formula","params":{"formula":"=AVERAGE(A1:A10)","cell":"B1"}}
en büyük değeri bul           | MAX                   | En büyük değeri verir | {"type":"formula","params":{"formula":"=MAX(A1:A10)","cell":"B1"}}
en küçük değeri yaz           | MIN                   | En küçük değeri verir | {"type":"formula","params":{"formula":"=MIN(A1:A10)","cell":"B1"}}
kaç tane sayı var             | COUNT                 | Sayısal değerlerin adedini bulur | {"type":"formula","params":{"formula":"=COUNT(A1:A10)","cell":"B1"}}
telefonları filtrele          | filter                | Belirtilen aralıkta filtre uygular | {"type":"filter","params":{"range":"A1:A20","criteria":"Telefon"}}
boşları temizle               | dataClean             | Boş hücreleri temizler | {"type":"dataClean","params":{"action":"fillEmpty","range":"A1:A20","value":"-"}}
tekrar edenleri sil           | dataClean             | Yinelenen değerleri siler | {"type":"dataClean","params":{"action":"removeDuplicates","range":"A1:A20"}}
harfleri birleştir            | CONCATENATE           | Metinleri birleştirir | {"type":"formula","params":{"formula":"=CONCATENATE(A1,B1)","cell":"C1"}}
metni büyüt                   | UPPER                 | Metni büyük harfe çevirir | {"type":"formula","params":{"formula":"=UPPER(A1)","cell":"B1"}}
eğer 10'dan büyükse evet yaz  | IF                    | Koşullu değer üretir | {"type":"formula","params":{"formula":"=IF(A1>10,\\"Evet\\",\\"Hayır\\")","cell":"B1"}}
satır ekle                    | addRow                | Yeni satır ekler | {"type":"addRow","params":{"rowIndex":5}}
tüm yazıları kalın yap        | cellFormat            | Yazı tipini kalınlaştır | {"type":"cellFormat","params":{"format":{"fontWeight":"bold"},"range":"A1:Z100"}}
yazı boyutunu 18 yap          | cellFormat            | Font boyutunu ayarlar | {"type":"cellFormat","params":{"format":{"fontSize":"18pt"},"range":"A1:Z100"}}
ortala                        | cellFormat            | Yazıları ortalar | {"type":"cellFormat","params":{"format":{"textAlign":"center"},"range":"A1:Z100"}}
kenarlık ekle                 | cellFormat            | Hücre kenarlığı ekler | {"type":"cellFormat","params":{"format":{"border":"1px solid #000"},"range":"A1:Z100"}}
tablonun grafiğini oluştur    | createChart           | Tablodaki verilerle grafik çizer | {"type":"createChart","params":{"chartType":"pie","range":"A1:B10"}}
pdf olarak dışa aktar         | exportPDF             | PDF'ye aktarır | {"type":"exportPDF","params":{}}
sayfayı kaydet                | save                  | Kaydeder | {"type":"save","params":{}}
aç                            | open                  | Dosya açar | {"type":"open","params":{}}

---
Ek örnekler:
Kullanıcı: "Tüm metni italik yap"
JSON: {"type":"cellFormat","params":{"format":{"fontStyle":"italic"},"range":"A1:Z100"}}
Kullanıcı: "Sadece ikinci satırı sil"
JSON: {"type":"deleteRow","params":{"rowIndex":1}}
Kullanıcı: "yazı boyutunu 20 yap"
JSON: {"type":"cellFormat","params":{"format":{"fontSize":"20pt"},"range":"A1:Z100"}}
Tabloya yeni bir sütun eklemen gerektiğinde, yeni sütunu her zaman mevcut sütunların EN SAĞINA ekle.
Örneğin, tablonun en sağında “Toplam” gibi bir sütun varsa, “Durum” sütunu ondan SONRA gelmeli.
Sadece veri satırlarına formül uygula.
Formül yazarken başlık satırına kesinlikle dokunma, formüllerin hücre referansı her zaman ilk veri satırından başlamalı (ör: başlık 2. satırdaysa formüller 3. satırdan başlar).
-**ÇOK ÖNEMLİ:** Eğer tablonun en sağında zaten başka veri sütunu varsa, “Durum” başlığını onun bir sağındaki sütuna ekle.Durum başlığı, mevcut sütunların en sağında olmalı.Aksi durum kabul edilemez! DURUM HER DAİM EN SAĞDA OLACAK!.
-**ÇOK ÖNEMLİ:** Sakın güncel tablodaki verilerin yerlerini değiştirme, eğer kullanıcı "adeti 5'ten fazla olanlara fazla, az olanlara az yaz" gibi bir şey derse, bu durumda yeni bir sütun ekle ve formülü o sütuna uygula sakın olaki aktif verileri yani adet,fiyat gibi sütünların yerlerini değiştirme!.Fazla veya az yazılacak sütun, mevcut sütunların en sağında olmalı.
Kurallar:
- Tablo başlıklarının (ör: Ürün, Adet, Fiyat, NUFUS ) gibi veriler hangi satırda olduğunu analiz et.
- Başlık satırına asla formül veya veri yazma; işlemleri yalnızca veri satırlarına uygula.
- Eğer kullanıcı "5 ve 5'ten fazla" diyorsa, matematiksel olarak ">=" kullan, formül buna göre oluşturulsun.
- Sütun ekleme, formül uygulama gibi adımlar gerektiğinde, bunları sırasıyla JSON dizisi olarak döndür.

- Tabloyu ve başlıkları dikkatlice incele, kullanıcı komutunu analiz et.
- Komut çok açık olmasa bile, kullanıcı ne yapmak istiyor tahmin et ve uygulanabilir en mantıklı çözümü üret.
- Kullanıcı sohbet dilinde, eksik, çok günlük, hatalı ya da saçma bir şey de yazsa, sen asla hata verme! Her zaman makul bir çözüm bul ve uygula.
- Gerekirse varsayım yap, önce veri temizlemesi, ek sütun/formül, hata kontrolü vs. adımları uygula.
- Birden fazla adım gerekiyorsa, her adımı sıralı bir JSON dizi olarak döndür:

- Tek adım gerekiyorsa yine {"type":"formula", ...} ile dön.
- "type" parametresine uygun türü yaz ("formula", "cellFormat", "dataClean", "sort", "filter", "createChart", "generateReport", "macro", "addColumn", "addRow", "custom" vs).
- Eğer kullanıcıdan gelen komut gereksiz, boş, ya da şaka/mantıksız ise de, olabilecek en mantıklı ve uygulanabilir (gerekirse uydurulmuş) bir Excel işlemi üret.
- Gerektiğinde, başlıkları analiz ederek hangi sütunun hangi veriyi tuttuğunu otomatik bul.
- Eğer kullanıcıdan gelen komut, tabloyu analiz etme, özetleme, raporlama gibi bir şeyse, JSON formatında bir rapor döndür:
- Eğer kullanıcıdan gelen komut tablo üzerinde otomatik bir işlem gerektiriyorsa, JSON komut(lar)ı ile dön.
- Eğer kullanıcı yazı boyutunu arttır derse 8,9,10,11,12,14,16,18,20,22,24,26,28,30 gibi değerlerden birini kullan.
- Eğer kullanıcı sadece bir bilgi veya analiz istiyorsa (ör: "En pahalı ürün nedir?", "En fazla adede sahip ürünü yaz" gibi), SADECE düz metin olarak tek satırlık cevabını ver, JSON döndürme.
- **ÇOK ÖNEMLİ:** Yanıtta SADECE GEÇERLİ ve SAF JSON kodu döndür, başında veya sonunda hiçbir açıklama, etiket, başlık, metin, karakter veya satır kullanma. Yalnızca doğrudan geçerli JSON string'i dön!

Sen bir insan gibi, gerçek bir Excel uzmanı gibi cevap ver; **asla pes etme, her durumda çözüm üret!**
`;
    // --- prompt burada bitiyor ---

    const response = await axios.post(
      'https://api.openai.com/v1/chat/completions',
      {
        model: 'gpt-4-1106-preview',
        messages: [
          { role: 'system', content: prompt }
        ],
        temperature: 0.0
      },
      {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${OPENAI_API_KEY}`
        }
      }
    );

    const content = response.data.choices?.[0]?.message?.content ?? "";
    console.log("OPENAI YANITI:", content);
    let jsonCommand = null;
    try { jsonCommand = JSON.parse(content); } catch {}
    // EN SAĞ SÜTUN DÜZELTME!
    if (jsonCommand && Array.isArray(jsonCommand) && req.body.sheet) {
  jsonCommand = fixColumnCommands(jsonCommand, req.body.sheet);

    }
    if (jsonCommand) res.json({ jsonCommand });
    else res.json({ command: content.trim() });
  } catch (e) {
    console.error('Assistant Hatası:', e?.response?.data || e.message);
    res.status(500).json({ error: 'Assistant API hatası.' });
  }
});

app.listen(3000, () => {
  console.log('✅ Sunucu http://localhost:3000 üzerinde çalışıyor');
});
