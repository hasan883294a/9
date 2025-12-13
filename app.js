// app.js

let selectedFile = null;
const fileInput = document.getElementById("file-input");
const processBtn = document.getElementById("process-btn");
const statusEl = document.getElementById("status");

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) {
    selectedFile = null;
    processBtn.disabled = true;
    statusEl.textContent = "";
    return;
  }
  selectedFile = file;
  processBtn.disabled = false;
  statusEl.textContent = `فایل انتخاب شد: ${file.name}`;
});

processBtn.addEventListener("click", () => {
  if (!selectedFile) return;
  processBtn.disabled = true;
  statusEl.textContent = "در حال خواندن و پردازش فایل...";

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      if (workbook.SheetNames.length === 0) {
        throw new Error("فایلی بدون شیت معتبر.");
      }

      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
// تبدیل شیت به JSON
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (rows.length === 0) {
        throw new Error("شیت خالی است.");
      }

      // تلاش برای پیدا کردن نام ستون‌ها بر اساس عنوان فارسی
      const sampleRow = rows[0];
      const headers = Object.keys(sampleRow);

      const variantCol = headers.find((h) =>
        h.toString().includes("کد تنوع")
      );
      const titleCol = headers.find((h) =>
        h.toString().includes("عنوان تنوع")
      );
      const creditAmountCol = headers.find((h) =>
        h.toString().includes("مبلغ نهایی بستانکار")
      );

      if (!variantCol || !creditAmountCol) {
        throw new Error(
          "ستون‌های «کد تنوع» یا «مبلغ نهایی بستانکار (﷼)» در هدر فایل پیدا نشد."
        );
      }

      // گروه‌بندی بر اساس کد تنوع
      const summaryMap = {};

      for (const row of rows) {
        const code = (row[variantCol] || "").toString().trim();
        if (!code) continue; // ردیف بدون کد تنوع

        const amountRaw = row[creditAmountCol];
        const amount = parseMoney(amountRaw);

        if (!summaryMap[code]) {
          summaryMap[code] = {
            variantCode: code,
            title: titleCol ? (row[titleCol] || "").toString().trim() : "",
            count: 0,
            sum: 0
          };
        }

        summaryMap[code].count += 1;
        summaryMap[code].sum += amount;
      }

      // تبدیل map به آرایه برای اکسل
      const summaryArray = Object.values(summaryMap).map((item) => {
        const avg = item.count ? Math.round(item.sum / item.count) : 0;
        return {
          "کد تنوع": item.variantCode,
          "عنوان تنوع (نمونه)": item.title,
          "تعداد ردیف": item.count,
          "جمع مبلغ نهایی بستانکار (ریال)": item.sum,
          "میانگین مبلغ نهایی بستانکار (ریال)": avg
        };
      });

      if (summaryArray.length === 0) {
        throw new Error("هیچ ردیفی با «کد تنوع» معتبر پیدا نشد.");
      }

      // ساخت ورک‌بوک خروجی
      const outWb = XLSX.utils.book_new();
      const outSheet = XLSX.utils.json_to_sheet(summaryArray);
      XLSX.utils.book_append_sheet(
        outWb,
        outSheet,
        "خلاصه بر اساس کد تنوع"
      );

      const outFileName = "summary_by_variant.xlsx";
      XLSX.writeFile(outWb, outFileName);

      statusEl.innerHTML =
        <span class="ok">پردازش انجام شد. فایل خروجی با نام <b>${outFileName}</b> دانلود شد.</span>;
    } catch (err) {
      console.error(err);
      statusEl.innerHTML =
        <span class="error">خطا در پردازش فایل: ${err.message}</span>;
    } finally {
      processBtn.disabled = false;
    }
  };

  reader.onerror = () => {
    statusEl.innerHTML = <span class="error">خطا در خواندن فایل.</span>;
    processBtn.disabled = false;
  };

  reader.readAsArrayBuffer(selectedFile);
});

/**
 * تبدیل مقدار مبلغ (که ممکن است عدد، رشته با جداکننده هزار، یا اعداد فارسی باشد)
 * به عدد جاوااسکریپتی.
 */
function parseMoney(value) {
  if (typeof value === "number") return value;
  if (value == null) return 0;

  const str = String(value);
  const persianDigits = "۰۱۲۳۴۵۶۷۸۹";

  let digitsOnly = "";
  for (const ch of str) {
    const idx = persianDigits.indexOf(ch);
    if (idx >= 0) {
      digitsOnly += idx.toString();
    } else if (ch >= "0" && ch <= "9") {
      digitsOnly += ch;
    }
  }

  if (!digitsOnly) return 0;
  return parseInt(digitsOnly, 10);
}