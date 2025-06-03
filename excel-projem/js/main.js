console.log("main.js yüklendi");

document.addEventListener('DOMContentLoaded', function () {
    const btn = document.getElementById('testChartBtn');
    if (!btn) {
        console.error("testChartBtn bulunamadı!");
        return;
    }
    console.log("testChartBtn bulundu, click event bağlanıyor");
    btn.onclick = function () {
        console.log("Butona tıklandı!");
        const labels = ['Telefon', 'Tablet', 'Laptop'];
        const data = [50000, 15000, 24000];
        drawChart(labels, data, 'pie', 'Ürün Bazında Satışlar');
    }
});

// AI Panel işlemlerini yapan kod
// (Senin index.html içinde zaten AI komutlarını çalıştıran scriptin var, 
//  burada sadece JSON olmayan yanıtlar için net hata gösterimi yapılacak.)

function processAIResult(aiResult) {
    let cmd = aiResult.jsonCommand;
    if (typeof cmd === "string") {
        cmd = cmd.replace(/```json|```js|```|\n/g, '').trim();
        try {
            cmd = JSON.parse(cmd);
        } catch (e) {
            appendToChat('❌', 'YANIT GEÇERLİ JSON DEĞİL: ' + (aiResult.command || ""));
            return;
        }
    }
    if (cmd) {
        runAICommand(cmd, true);
    } else {
        appendToChat('❌', 'Beklenen formatta JSON komutu alınamadı.');
    }
}
