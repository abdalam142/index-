<!DOCTYPE html>
<html lang="ar">
<head>
  <meta charset="UTF-8">
  <title>بحث في ملف Excel مع الباركود</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://unpkg.com/@zxing/library@latest"></script>
  <style>
    body { font-family: Arial, sans-serif; text-align: center; padding: 20px; direction: rtl; }
    #searchBar { width: 50%; padding: 10px; font-size: 16px; margin-top: 20px; }
    table { width: 80%; margin: 20px auto; border-collapse: collapse; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
    th { background-color: #f2f2f2; }
    #cameraContainer { display: inline-flex; align-items: center; gap: 10px; margin-top: 10px; }
    #cameraView { width: 120px; height: 80px; border: 1px solid #ccc; object-fit: cover; }
    #startCamera { padding: 6px 12px; cursor: pointer; }
  </style>
</head>
<body>
  <h2>بحث في ملف Excel مع الباركود</h2>
  <input type="file" id="upload" accept=".xlsx, .xls">
  <select id="sheetSelect"></select>
  <br>
  <input type="text" id="searchBar" placeholder="اكتب الاسم أو الباركود أو الوصف واضغط إنتر">
  <div id="cameraContainer">
    <button id="startCamera">تشغيل الكاميرا</button>
    <video id="cameraView" autoplay></video>
  </div>
  <table id="results"></table>

  <script>
    let excelData = {};
    let selectedSheet = "";

    document.getElementById('upload').addEventListener('change', handleFile, false);
    document.getElementById('sheetSelect').addEventListener('change', () => {
      selectedSheet = document.getElementById('sheetSelect').value;
    });
    document.getElementById('searchBar').addEventListener('keypress', function (e) {
      if (e.key === 'Enter') search();
    });

    function handleFile(e) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetSelect = document.getElementById('sheetSelect');
        sheetSelect.innerHTML = '';
        workbook.SheetNames.forEach(name => {
          let option = document.createElement('option');
          option.value = name;
          option.text = name;
          sheetSelect.appendChild(option);
        });
        selectedSheet = workbook.SheetNames[0];

        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          excelData[sheetName] = rows;
        });
      };
      reader.readAsArrayBuffer(file);
    }

    function search(queryText) {
      const query = (queryText || document.getElementById('searchBar').value).trim().toLowerCase();
      if (!query || !selectedSheet) return;

      const rows = excelData[selectedSheet];
      const resultsTable = document.getElementById('results');
      resultsTable.innerHTML = '';

      // Create header row A-D
      let headerRow = `<tr><th>A</th><th>B</th><th>C</th><th>D</th></tr>`;
      resultsTable.innerHTML = headerRow;

      rows.forEach((row, index) => {
        if (index === 0) return; // skip header
        let a = (row[0] || "").toString().toLowerCase();
        let b = (row[1] || "").toString();
        let c = (row[2] || "").toString().toLowerCase();
        let d = (row[3] || "").toString().toLowerCase();

        if (a.includes(query) || c.includes(query) || d.includes(query)) {
          let tr = `<tr>
            <td>${row[0] || ''}</td>
            <td>${row[1] || ''}</td>
            <td>${row[2] || ''}</td>
            <td>${query}</td>
          </tr>`;
          resultsTable.innerHTML += tr;
        }
      });
    }

    // ZXing Barcode Scanner
    document.getElementById('startCamera').addEventListener('click', () => {
      const codeReader = new ZXing.BrowserMultiFormatReader();
      const video = document.getElementById('cameraView');

      codeReader.decodeFromVideoDevice(null, video, (result, err) => {
        if (result) {
          const barcode = result.text;
          document.getElementById('searchBar').value = barcode;
          search(barcode);
        }
      });
    });
  </script>
</body>
</html>
