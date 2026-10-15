function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holdingsSheet = ss.getSheetByName("보유종목_상세");
  const holdings = [];
  if (holdingsSheet) {
    const data = holdingsSheet.getDataRange().getValues();
    let headerIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].join("").indexOf("운용 계좌") !== -1 || data[i].join("").indexOf("종목명") !== -1) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx !== -1) {
      for (let r = headerIdx + 1; r < data.length; r++) {
        const row = data[r];
        const name = String(row[2] || row[3] || "").trim();
        if (!name || name === "합계" || name.indexOf("총") === 0) continue;
        const offset = (typeof row[0] === "number" || String(row[1]).indexOf("계좌") !== -1) ? 0 : 1;
        const qty = Number(row[4 + offset]) || 0;
        if (qty <= 0) continue;
        holdings.push({
          no: row[0 + offset],
          account: String(row[1 + offset]),
          name: String(row[2 + offset]),
          category: String(row[3 + offset]),
          qty: qty,
          buyPrice: Number(row[5 + offset]) || 0,
          curPrice: Number(row[6 + offset]) || 0,
          totalBuy: Number(row[7 + offset]) || 0,
          totalEval: Number(row[8 + offset]) || 0,
          pnl: Number(row[9 + offset]) || 0,
          returnRate: Number(row[10 + offset]) || 0,
          memo: String(row[11 + offset] || "")
        });
      }
    }
  }
  const dailySheet = ss.getSheetByName("일별_자산추이");
  const dailyTrend = [];
  if (dailySheet) {
    const dData = dailySheet.getDataRange().getValues();
    for (let i = 0; i < dData.length; i++) {
      const row = dData[i];
      for (let c = 0; c < 3; c++) {
        const val = row[c];
        if (val instanceof Date || (typeof val === "string" && /^\d{4}-\d{2}-\d{2}$/.test(val.trim()))) {
          const dateStr = (val instanceof Date) ? Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd") : val.trim();
          const assetVal = Number(row[c + 2]) || 0;
          if (assetVal > 0) {
            dailyTrend.push({
              date: dateStr,
              time: String(row[c + 1] || "07:00"),
              totalAsset: assetVal,
              diff: Number(row[c + 3]) || 0,
              diffRate: row[c + 4],
              note: String(row[c + 5] || "")
            });
          }
          break;
        }
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    updatedAt: Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"),
    holdings: holdings,
    dailyTrend: dailyTrend
  })).setMimeType(ContentService.MimeType.JSON);
}