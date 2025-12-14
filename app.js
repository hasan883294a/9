// app.js

window.addEventListener("DOMContentLoaded", function () {
  console.log("DOM loaded, initializing app.js");

  const fileInput = document.getElementById("file-input");
  const processBtn = document.getElementById("process-btn");
  const statusEl = document.getElementById("status");

  if (!fileInput  !processBtn  !statusEl) {
    console.error("Elements not found. Check IDs in index.html");
    return;
  }

  let selectedFile = null;

  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    console.log("fileInput change:", file);
    if (!file) {
      selectedFile = null;
      processBtn.disabled = true;
      statusEl.textContent = "فایلی انتخاب نشد.";
      return;
    }
    selectedFile = file;
    processBtn.disabled = false;
    statusEl.textContent = `فایل انتخاب شد: ${file.name}`;
  });

  processBtn.addEventListener("click", () => {
    console.log("process button clicked");
    if (!selectedFile) {
      statusEl.textContent = "ابتدا فایل اکسل را انتخاب کنید.";
      return;
    }

    processBtn.disabled = true;
    statusEl.textContent = "در حال خواندن و پردازش فایل...";

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        if (typeof XLSX === "undefined") {
          throw new Error("کتابخانه XLSX لود نشده است.");
        }

        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        if (workbook.SheetNames.length === 0) {
          throw new Error("فایلی بدون شیت معتبر.");
        }

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (rows.length === 0) {
          throw new Error("شیت خالی است.");
        }

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

        const summaryMap = {};

        for (const row of rows) {
          const code = (row[variantCol] || "").toString().trim();
          if (!code) continue;

          const amountRaw = row[creditAmountCol];
          const amount = parseMoney(amountRaw);

          if (!summaryMap[code]) {
            summaryMap[code] = {
              variantCode: code,
              title: titleCol ? (row[titleCol] || "").toString().trim() : "",
              count: 0,
              sum: 0,
            };
          }

          summaryMap[code].count += 1;
          summaryMap[code].sum += amount;
        }

        const summaryArray = Object.values(summaryMap).map((item) => {
          const avg = item.count ? Math.round(item.sum / item.count) : 0;
          return {
            "کد تنوع": item.variantCode,
            "عنوان تنوع (نمونه)": item.title,
            "تعداد ردیف": item.count,
            "جمع مبلغ نهایی بستانکار (ریال)": item.sum,
            "میانگین مبلغ نهایی بستانکار (ریال)": avg,
          };
        });

        if (summaryArray.length === 0) {
          throw new Error("هیچ ردیفی با «کد تنوع» معتبر پیدا نشد.");
        }

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
});
