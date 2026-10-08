const { google } = require('googleapis');

const jsonResponse = (res, statusCode, data) => {
  res.status(statusCode).json(data);
};

// Google Sheet ထဲရှိ Tab နာမည်ကို အလိုအလျောက် ရှာဖွေပေးသည့် Helper Function
async function getTabTitle(sheets, spreadsheetId, preferredName) {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetList = meta.data.sheets || [];
    
    const matched = sheetList.find(s => 
      s.properties.title.trim().toLowerCase() === preferredName.trim().toLowerCase()
    );
    if (matched) {
      return matched.properties.title;
    }
    
    if (sheetList.length > 0) {
      return sheetList[0].properties.title;
    }
  } catch (e) {}
  return preferredName;
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    let privateKey = process.env.PRIVATE_KEY || '';
    if (privateKey.startsWith('"') && privateKey.endsWith('"')) {
      privateKey = privateKey.slice(1, -1);
    }
    privateKey = privateKey.replace(/\\n/g, '\n');

    const auth = new google.auth.JWT(
      process.env.CLIENT_EMAIL,
      null,
      privateKey,
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth });
    
    const SHEET_ID_KPI = process.env.SHEET_ID_KPI;
    const SHEET_ID_EMP = process.env.SHEET_ID_EMP;

    // =========================================================================
    // 🌟 GET REQUEST (Employee ID Lookup)
    // =========================================================================
    if (req.method === 'GET') {
      const { action, id } = req.query;

      if (action === 'getEmployee' || id) {
        const searchId = id ? String(id).trim() : "";
        if (!searchId) {
          return jsonResponse(res, 200, { status: "error", message: "ID လိုအပ်ပါသည်။" });
        }

        try {
          const empTab = await getTabTitle(sheets, SHEET_ID_EMP, "Mgr Fb-30%");
          const empRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID_EMP,
            range: `'${empTab}'!A2:B3000`,
            valueRenderOption: 'FORMATTED_VALUE'
          });

          const rows = empRes.data.values || [];
          const cleanSearch = searchId.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

          for (let i = 0; i < rows.length; i++) {
            const rawName = rows[i][0] ? String(rows[i][0]).trim() : "";
            const rawId = rows[i][1] ? String(rows[i][1]).trim() : "";

            const cleanIdCol = rawId.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();
            const cleanNameCol = rawName.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

            if (cleanIdCol === cleanSearch && cleanSearch !== "") {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: rawName });
            }
            if (cleanNameCol === cleanSearch && cleanSearch !== "") {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: rawId });
            }
          }
          
          return jsonResponse(res, 200, { status: "error", message: "Employee ID မရှိသေးပါ" });

        } catch (e) {
          return jsonResponse(res, 200, { status: "error", message: "Employee Sheet Error: " + e.message });
        }
      }

      return jsonResponse(res, 200, { status: "running", message: "Vercel KPI API is active!" });
    }

    // =========================================================================
    // 🌟 POST REQUEST (KPI Submission & Duplicate Check)
    // =========================================================================
    if (req.method === 'POST') {
      let bodyData = req.body;
      if (typeof bodyData === 'string') {
        try { bodyData = JSON.parse(bodyData); } catch (e) {}
      }

      const { date, serviceId, customerName, financeName, financeId, reason, duplicateReason, isConfirmed } = bodyData || {};

      if (!date || !serviceId || !financeName) {
        return jsonResponse(res, 200, { status: "error", message: "လိုအပ်သော Parameter များ မပါဝင်ပါ။" });
      }

      const kpiTab = await getTabTitle(sheets, SHEET_ID_KPI, "KPI Calculation");

      const kpiRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID_KPI,
        range: `'${kpiTab}'!A:F`
      }).catch((err) => {
        throw new Error("KPI Sheet ဖတ်မရပါ။ Error: " + err.message);
      });

      const allRows = kpiRes.data.values || [];
      const isConfirmMode = (String(isConfirmed).toLowerCase() === "true");
      
      let count = 0;
      let otherFinances = [];
      const inputDate = new Date(date).toDateString();

      const startIndex = Math.max(1, allRows.length - 400);

      for (let i = startIndex; i < allRows.length; i++) {
        const row = allRows[i];
        if (row[1]) {
          const rowDate = new Date(row[1]).toDateString();

          if (rowDate === inputDate && String(row[2]).trim() === String(serviceId).trim()) {
            count++;
            
            const savedFinanceName = row[4] ? String(row[4]).trim() : "";
            const savedFinanceId = row[5] ? String(row[5]).trim() : "";
            
            if (savedFinanceId !== String(financeId).trim()) {
              if (savedFinanceName && !otherFinances.includes(savedFinanceName)) {
                otherFinances.push(savedFinanceName);
              }
            }
          }
        }
      }

      if (count > 0 && !isConfirmMode) {
        let msg = "";
        if (otherFinances.length > 0) {
          msg = `ဒီ Service ID ကို Finance ${otherFinances.join(", ")} မှ တင်ထားပြီးသားဖြစ်ပြီး သင်သည် ဒီနေ့ (${count + 1}) ကြိမ်မြောက်တင်ခြင်းဖြစ်သောကြောင့် Duplicate Reason ထည့်ပေးပါ။`;
        } else {
          msg = `ဤService ID ကို ဒီနေ့တည်းမှာသင်ကိုယ်တိုင်တင်ပြီးသားဖြစ်နေသောကြောင့် သင်သည် (${count + 1}) ကြိမ်မြောက်တင်ခြင်းဖြစ်ပါသည်။ Duplicate Reason ထည့်ပါ။`;
        }
        return jsonResponse(res, 200, { status: "duplicate_prompt", message: msg });
      }

      // 🇲🇲 မြန်မာစံတော်ချိန် (UTC+6:30) ဖြင့် Timestamp ပြုလုပ်ခြင်း
      const now = new Date();
      const mmTime = new Date(now.getTime() + (6.5 * 60 * 60 * 1000));
      const month = mmTime.getUTCMonth() + 1;
      const day = mmTime.getUTCDate();
      const year = mmTime.getUTCFullYear();
      const hours = mmTime.getUTCHours();
      const minutes = String(mmTime.getUTCMinutes()).padStart(2, '0');
      const seconds = String(mmTime.getUTCSeconds()).padStart(2, '0');

      const timestamp = `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;

      const appendValues = [[timestamp, date, serviceId, customerName, financeName, financeId, reason, duplicateReason || ""]];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID_KPI,
        range: `'${kpiTab}'!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: appendValues }
      });

      return jsonResponse(res, 200, { status: "success", message: "အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။" });
    }

  } catch (error) {
    return jsonResponse(res, 500, { status: "error", message: error.message || String(error) });
  }
};
