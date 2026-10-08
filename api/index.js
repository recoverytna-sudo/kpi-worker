const { google } = require('googleapis');

const jsonResponse = (res, statusCode, data) => {
  res.status(statusCode).json(data);
};

module.exports = async (req, res) => {
  // CORS Headers
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
          const empRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID_EMP,
            range: "'Mgr Fb-30%'!A2:B3000",
            valueRenderOption: 'FORMATTED_VALUE'
          });

          const rows = empRes.data.values || [];
          
          // ရှာဖွေမည့် ID မှ ဂဏန်း/စာသားသီးသန့် စစ်ထုတ်ခြင်း
          const cleanSearch = searchId.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

          for (let i = 0; i < rows.length; i++) {
            const rawName = rows[i][0] ? String(rows[i][0]).trim() : "";
            const rawId = rows[i][1] ? String(rows[i][1]).trim() : "";

            // Sheet ထဲရှိ Column B (ID) မှ ကော်မာနှင့် သင်္ကေတများ ရှင်းလင်းခြင်း
            const cleanIdCol = rawId.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();
            const cleanNameCol = rawName.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

            // ID ဖြင့် ရှာဖွေခြင်း
            if (cleanIdCol === cleanSearch && cleanSearch !== "") {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: rawName });
            }
            // Name ဖြင့် ရှာဖွေခြင်း
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

      const kpiRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID_KPI,
        range: "'KPI Calculation'!A:F"
      }).catch((err) => {
        throw new Error("KPI Sheet ဖတ်မရပါ။ Tab နာမည် 'KPI Calculation' ဟုတ်မဟုတ် စစ်ဆေးပါ။ Error: " + err.message);
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

      const timestamp = new Date().toISOString();
      const appendValues = [[timestamp, date, serviceId, customerName, financeName, financeId, reason, duplicateReason || ""]];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID_KPI,
        range: "'KPI Calculation'!A1",
        valueInputOption: 'USER_ENTERED',
        resource: { values: appendValues }
      });

      return jsonResponse(res, 200, { status: "success", message: "အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။" });
    }

  } catch (error) {
    return jsonResponse(res, 500, { status: "error", message: error.message || String(error) });
  }
};
