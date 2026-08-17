# 🚜 볼보 굴착기 DTC 고장진단 뷰어 & 분석 시스템

## 1. 시스템 구조 개요
* **프론트엔드 (GitHub Pages):** PWA 기반 반응형 뷰어 (`index.html`) 및 캡처/PDF 파서 (`parser.html`)
* **백엔드 (Google Apps Script):** Google Drive OCR 문서 판독 + 14개 필드 표(Table) 구조화 파서 + Google Sheet REST API
* **데이터베이스 (Google Sheets):** 14개 표준 진단 데이터 저장소

---

## 2. Google Apps Script 백엔드 소스코드

### [1] `appsscript.json` (매니페스트 설정)
```json
{
  "timeZone": "Asia/Seoul",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "serviceId": "drive",
        "version": "v3"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "[https://www.googleapis.com/auth/drive](https://www.googleapis.com/auth/drive)",
    "[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)",
    "[https://www.googleapis.com/auth/documents](https://www.googleapis.com/auth/documents)",
    "[https://www.googleapis.com/auth/script.external_request](https://www.googleapis.com/auth/script.external_request)"
  ]
}
```

### [2] `Code.gs` (통합 서버 엔진)
```javascript
function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(err) { data = e.parameter; }
    } else if (e.parameter) {
      data = e.parameter;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DTC데이터") || ss.getActiveSheet();
    initSheetHeaders(sheet);

    if (data.action === "parse_image" && data.imageBase64) {
      var imageBytes = Utilities.base64Decode(data.imageBase64);
      var imageBlob = Utilities.newBlob(imageBytes, data.mimeType || "image/png", "temp_capture");
      var tempFileId = "";

      if (Drive.Files && Drive.Files.create) {
        var fileMetadata = { name: "[OCR_TEMP]", mimeType: "application/vnd.google-apps.document" };
        var tempFile = Drive.Files.create(fileMetadata, imageBlob, { ocrLanguage: "ko" });
        tempFileId = tempFile.id;
      } else if (Drive.Files && Drive.Files.insert) {
        var fileMetadata = { title: "[OCR_TEMP]", mimeType: "application/vnd.google-apps.document" };
        var tempFile = Drive.Files.insert(fileMetadata, imageBlob, { ocr: true, ocrLanguage: "ko" });
        tempFileId = tempFile.id;
      }

      var doc = DocumentApp.openById(tempFileId);
      var parsed = parseFromDocTables(doc);
      var fullText = doc.getBody().getText();
      DriveApp.getFileById(tempFileId).setTrashed(true);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success", parsed: parsed, rawText: fullText
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var rowData = [
      String(data.dtc || ""), String(data.modelType || ""), String(data.factory || ""),
      String(data.serialStart || ""), String(data.serialEnd || ""), String(data.ecu || ""),
      String(data.compFull || ""), String(data.standard || ""), String(data.event || ""),
      String(data.symptom || ""), String(data.condition || ""), String(data.preDtc || ""),
      String(data.causes || ""), String(data.actions || "")
    ];

    var allData = sheet.getDataRange().getValues();
    var targetRowIndex = -1;

    if (data.mode === "overwrite" && data.dtc) {
      for (var r = 1; r < allData.length; r++) {
        if (String(allData[r][0]).trim().toUpperCase() === String(data.dtc).trim().toUpperCase()) {
          targetRowIndex = r + 1;
          break;
        }
      }
    }

    if (targetRowIndex > 0) {
      var range = sheet.getRange(targetRowIndex, 1, 1, rowData.length);
      range.setNumberFormat("@");
      range.setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
      sheet.getRange(sheet.getLastRow(), 1, 1, rowData.length).setNumberFormat("@");
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", result: "saved" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function parseFromDocTables(doc) {
  var body = doc.getBody();
  var fullText = body.getText();
  var result = {
    dtc: '', modelType: 'EW140E Volvo', factory: '', serialStart: '', serialEnd: '',
    ecu: '', compFull: '', standard: '', event: '', symptom: '없음',
    condition: '없음', preDtc: '없음', causes: '없음', actions: '없음'
  };

  var dtcMatch = fullText.match(/([P|U|B|C][0-9A-Z]{5,7}|PID\s*\d+|SE\d+-\d+)/);
  if (dtcMatch) result.dtc = dtcMatch[1].replace(/\s+/g, '');

  var clean = function(s) { return s ? s.replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim() : ''; };
  var cleanBullets = function(s) {
    if (!s) return '없음';
    var lines = s.split(/[\n¡○•●\t]+/).map(function(l) { return l.replace(/^[:\s\-O0]+/, '').trim(); }).filter(function(l) { return l && l !== '없음'; });
    return lines.length > 0 ? lines.join(' | ') : '없음';
  };

  var tables = body.getTables();
  if (tables && tables.length > 0) {
    for (var t = 0; t < tables.length; t++) {
      var table = tables[t];
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        var cells = row.getNumCells();
        if (cells === 4) {
          var c0 = clean(row.getCell(0).getText());
          if (c0.includes('Volvo') || c0.includes('EW') || c0.includes('EC')) {
            result.modelType = c0;
            result.factory = clean(row.getCell(1).getText());
            result.serialStart = clean(row.getCell(2).getText());
            result.serialEnd = clean(row.getCell(3).getText());
          }
        }
        if (cells >= 2) {
          var label = clean(row.getCell(0).getText());
          var valText = row.getCell(1).getText().trim();
          if (label.includes('컨트롤 유니트')) result.ecu = clean(valText);
          else if (label.includes('고장 유형') || label.includes('구성부품')) result.compFull = clean(valText);
          else if (label.includes('기준')) result.standard = clean(valText);
          else if (label.includes('고장 이벤트')) result.event = clean(valText);
          else if (label.includes('증상')) result.symptom = cleanBullets(valText);
          else if (label.includes('조건')) result.condition = cleanBullets(valText);
          else if (label.includes('선행조치')) {
            var codes = valText.match(/([P|U|B|C][0-9A-Z]{5,7})/g) || [];
            var validCodes = codes.filter(function(c) { return c !== result.dtc && !c.startsWith('DTC'); });
            result.preDtc = Array.from(new Set(validCodes)).join(' | ') || '없음';
          }
          else if (label.includes('가능한 이유')) result.causes = cleanBullets(valText);
          else if (label.includes('조치')) result.actions = cleanBullets(valText);
        }
      }
    }
  }

  if (!result.compFull && result.dtc) {
    var compMatch = fullText.match(new RegExp("고장\\s*유형[\\s\\S]*?(" + result.dtc + "[\\s\\S]*?)(?=기준|$)", "i"));
    if (compMatch) result.compFull = clean(compMatch[1]);
  }
  if (!result.event) {
    var evMatch = fullText.match(/고장\s*이벤트\s*([\s\S]*?)(?=증상|$)/);
    if (evMatch) result.event = clean(evMatch[1]);
  }
  return result;
}

function initSheetHeaders(sheet) {
  var headers = [
    "에러코드", "모델 종류", "생산 사업장", "일련 번호 시작", "일련 번호 종료",
    "컨트롤 유니트", "DTC 코드, 시스템/구성부품, 고장유형", "기준", "고장이벤트",
    "증상", "조건", "선행조치DTC", "가능한 이유", "조치"
  ];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#d9d9d9").setFontWeight("bold");
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DTC데이터") || ss.getSheets()[0];
    var values = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (!row || row.join('').trim() === '') continue;
      results.push({
        dtc: String(row[0]||''), modelType: String(row[1]||''), factory: String(row[2]||''),
        serialStart: String(row[3]||''), serialEnd: String(row[4]||''), ecu: String(row[5]||''),
        compFull: String(row[6]||''), standard: String(row[7]||''), event: String(row[8]||''),
        symptom: String(row[9]||''), condition: String(row[10]||''), preDtc: String(row[11]||''),
        causes: String(row[12]||''), actions: String(row[13]||'')
      });
    }
    return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

## 3. 초기 설정 및 배포 가이드
1. 구글 스프레드시트 생성 (`DTC데이터` 시트 확인)
2. **확장 프로그램 > Apps Script** 이동
3. 좌측 **⚙️ 프로젝트 설정** ➔ `'appsscript.json' 매니페스트 파일 표시` 체크
4. `appsscript.json` 및 `Code.gs` 소스코드 붙여넣기
5. 좌측 **서비스 (+)** ➔ **`Drive API (v3)`** 추가
6. 우측 상단 **배포 > 새 배포** (다음 사용자로 실행: `나`, 액세스: `모든 사용자`)
7. 생성된 **웹 앱 URL**을 `index.html`과 `parser.html`의 `API_URL` / `scriptUrl`에 입력
