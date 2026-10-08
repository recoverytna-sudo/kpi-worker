const { google } = require('googleapis');

// -------------------------------------------------------------------------
// Helper Function: Respond JSON
// -------------------------------------------------------------------------
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
    // Google Service Account Authentication
    const auth = new google.auth.JWT(
      process.env.CLIENT_EMAIL,
      null,
      process.env.PRIVATE_KEY ? process.env.PRIVATE_KEY.replace(/\\n/g, '\n') : '',
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth });
    
    // Environment Variables မှ Sheet IDs ယူခြင်း
    const SHEET_ID_KPI = process.env.SHEET_ID_KPI;
    const SHEET_ID_EMP = process.env.SHEET_ID_EMP;
    const SHEET_NAME_KPI = "KPI Calculation";
    const SHEET_NAME_EMP = "Mgr Fb-30%";

    // =========================================================================
    // 🌟 GET REQUEST HANDLER (Employee Lookup)
    // =========================================================================
    if (req.method === 'GET') {
      const { action, id } = req.query;

      if (action === 'getEmployee' || id) {
        const searchId = id ? String(id).trim() : "";
        if (!searchId) {
          return jsonResponse(res, 200, { status: "error", message: "ID လိုအပ်ပါသည်။" });
        }

        try {
          // Employee Sheet မှ Data ဆွဲယူခြင်း
          const empRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID_EMP,
            range: `${SHEET_NAME_EMP}!A2:B1000`
          });

          const rows = empRes.data.values || [];
          for (let i = 0; i < rows.length; i++) {
            const nameCol = rows[i][0] ? String(rows[i][0]).trim() : "";
            const idCol = rows[i][1] ? String(rows[i][1]).trim() : "";

            if (idCol === searchId) {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: nameCol });
            }
            if (nameCol === searchId) {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: idCol });
            }
          }
          
          return jsonResponse(res, 200, { status: "error", message: "Employee ID မရှိသေးပါ" });

        } catch (e) {
          return jsonResponse(res, 200, { status: "error", message: "Employee Sheet ကို ဖတ်မရပါ။ Access ပေးထားခြင်း ရှိ/မရှိ စစ်ဆေးပါ။" });
        }
      }

      return jsonResponse(res, 200, { status: "running", message: "Vercel API is working!" });
    }

    // =========================================================================
    // 🌟 POST REQUEST HANDLER (KPI Submission & Duplicate Check)
    // =========================================================================
    if (req.method === 'POST') {
      const { date, serviceId, customerName, financeName, financeId, reason, duplicateReason, isConfirmed } = req.body;

      if (!date || !serviceId || !financeName) {
        return jsonResponse(res, 200, { status: "error", message: "လိုအပ်သော Parameter များ မပါဝင်ပါ။" });
      }

      // 1. KPI Sheet မှ ဒေတာများကို ဆွဲယူ၍ Duplicate စစ်ဆေးခြင်း
      const kpiRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID_KPI,
        range: `${SHEET_NAME_KPI}!A:F`
      }).catch(async (err) => {
        // Sheet မရှိသေးပါက Error တက်မည်။ (ဒီနေရာမှာ Vercel မှနေ၍ Sheet အသစ်ဆောက်မပေးနိုင်ပါ။)
        // ထို့ကြောင့် Google Sheet ဘက်တွင် 'KPI Calculation' အမည်ဖြင့် Tab တစ်ခု ကြိုတင်ဆောက်ထားရန် လိုအပ်ပါသည်။
        throw new Error("KPI Calculation Sheet ကို ဖတ်မရပါ။ Sheet နာမည် မှန်ကန်မှု ရှိ/မရှိ စစ်ဆေးပါ။");
      });

      const allRows = kpiRes.data.values || [];
      const isConfirmMode = (String(isConfirmed).toLowerCase() === "true");
      
      let count = 0;
      let otherFinances = [];
      const inputDate = new Date(date).toDateString();

      // နောက်ဆုံး Row 400 ကိုပဲ စစ်ဆေးမည်
      const startIndex = Math.max(1, allRows.length - 400);

      for (let i = startIndex; i < allRows.length; i++) {
        const row = allRows[i];
        if (row[1]) {
          const rowDate = new Date(row[1]).toDateString();

          // Same Date နှင့် Same Service ID တူနေလျှင်
          if (rowDate === inputDate && String(row[2]).trim() === String(serviceId).trim()) {
            count++;
            
            const savedFinanceName = row[4] ? String(row[4]).trim() : "";
            const savedFinanceId = row[5] ? String(row[5]).trim() : "";
            
            // အခြား Finance က တင်ထားတာ ဖြစ်လျှင်
            if (savedFinanceId !== String(financeId).trim()) {
              if (savedFinanceName && !otherFinances.includes(savedFinanceName)) {
                otherFinances.push(savedFinanceName);
              }
            }
          }
        }
      }

      // Duplicate တွေ့ရှိပါက Prompt တောင်းမည်
      if (count > 0 && !isConfirmMode) {
        let msg = "";
        if (otherFinances.length > 0) {
          msg = `ဒီ Service ID ကို Finance ${otherFinances.join(", ")} မှ တင်ထားပြီးသားဖြစ်ပြီး သင်သည် ဒီနေ့ (${count + 1}) ကြိမ်မြောက်တင်ခြင်းဖြစ်သောကြောင့် Duplicate Reason ထည့်ပေးပါ။`;
        } else {
          msg = `ဤService ID ကို ဒီနေ့တည်းမှာသင်ကိုယ်တိုင်တင်ပြီးသားဖြစ်နေသောကြောင့် သင်သည် (${count + 1}) ကြိမ်မြောက်တင်ခြင်းဖြစ်ပါသည်။ Duplicate Reason ထည့်ပါ။`;
        }
        return jsonResponse(res, 200, { status: "duplicate_prompt", message: msg });
      }

      // 2. Duplicate မရှိလျှင် သို့မဟုတ် Confirm လုပ်ပြီးလျှင် Data သိမ်းမည်
      const timestamp = new Date().toISOString();
      const appendValues = [[timestamp, date, serviceId, customerName, financeName, financeId, reason, duplicateReason || ""]];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID_KPI,
        range: `${SHEET_NAME_KPI}!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: appendValues }
      });

      return jsonResponse(res, 200, { status: "success", message: "အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။" });
    }

  } catch (error) {
    return jsonResponse(res, 500, { status: "error", message: error.message || String(error) });
  }
};
