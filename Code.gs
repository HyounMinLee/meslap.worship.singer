// Code.gs

// ==========================================
// 공통 설정
// ==========================================

// 📁 원본 악보 PPT들이 모여있는 '검색 대상 폴더 ID' (업로드 변환 결과도 여기에 저장됨)
var SEARCH_FOLDER_ID = "1sCQPVxJOQUsNfwb5kfu5UvwPlI3PPJG9";

// 🔤 병합 PPT에 적용할 수 있는 글꼴 (index.html 의 MERGE_FONTS 와 반드시 동일하게 유지)
var MERGE_FONT_WHITELIST = ["DX시인과나", "맑은 고딕"];
var MERGE_FONT_DEFAULT = "DX시인과나";

// 화면 표시 이름 -> 실제로 슬라이드에 기록할 글꼴 이름
// PowerPoint 에서 글꼴이 대체되면 이 값만 바꾸면 된다. (예: "맑은 고딕" -> "Malgun Gothic")
var MERGE_FONT_MAP = {
  "DX시인과나": "DX시인과나",
  "맑은 고딕": "맑은 고딕"
};

// 🔠 병합 PPT 글자 크기 허용 범위 (pt)
var FONT_SIZE_MIN = 10;
var FONT_SIZE_MAX = 40;

var MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
var MIME_PPT = "application/vnd.ms-powerpoint";
var MIME_GSLIDES = "application/vnd.google-apps.presentation";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    var searchFolderId = SEARCH_FOLDER_ID;

    // --- [액션 A] 폴더 내 파일 이름 단일 검색 ---
    if (action === "search") {
      var query = String(data.query || "").trim();
      if (!query) throw new Error("검색어가 비어 있습니다.");

      var folder = DriveApp.getFolderById(searchFolderId);
      var files = folder.searchFiles("title contains '" + escapeQuery_(query) + "' and mimeType = '" + MIME_GSLIDES + "'");
      var results = [];

      while (files.hasNext()) {
        var file = files.next();
        results.push({ name: file.getName(), id: file.getId() });
      }
      return jsonOut_({ status: "success", data: results });
    }

    // --- [액션 B] 텍스트 리스트 다중 자동 검색 (띄어쓰기 무시) ---
    else if (action === "batchSearch") {
      var queries = data.queries || [];
      var folderB = DriveApp.getFolderById(searchFolderId);
      var filesB = folderB.searchFiles("mimeType = '" + MIME_GSLIDES + "'");

      // 폴더 내의 모든 PPT 파일을 가져와 띄어쓰기를 없앤 비교용 이름을 만듦
      var fileList = [];
      while (filesB.hasNext()) {
        var f = filesB.next();
        var rawName = f.getName();
        // 파일 이름에서 띄어쓰기 제거 및 소문자 변환
        var noSpaceName = rawName.replace(/\s/g, '').toLowerCase();
        fileList.push({ name: rawName, id: f.getId(), noSpaceName: noSpaceName });
      }

      var results = [];
      for (var i = 0; i < queries.length; i++) {
        var q = String(queries[i]).replace(/\s/g, '').toLowerCase(); // 검색어도 띄어쓰기 제거
        var exact = null, partial = null;

        for (var j = 0; j < fileList.length; j++) {
          // 완전일치를 우선한다. ("은혜" 검색 시 "은혜아니면"이 먼저 잡히는 문제 방지)
          if (fileList[j].noSpaceName === q) { exact = fileList[j]; break; }
          if (partial === null && fileList[j].noSpaceName.indexOf(q) !== -1) partial = fileList[j];
        }

        var hit = exact || partial;
        if (hit) {
          results.push({ query: queries[i], name: hit.name, id: hit.id, found: true, matchType: exact ? "exact" : "partial" });
        } else {
          results.push({ query: queries[i], name: null, id: null, found: false });
        }
      }

      return jsonOut_({ status: "success", data: results });
    }

    // --- [액션 C] PPT 파일 업로드 → 구글 슬라이드 변환 ---
    else if (action === "upload") {
      var origName = String(data.fileName || "").trim();
      var base64 = data.fileData;
      var overwrite = (data.overwrite === true);

      if (!origName) throw new Error("파일 이름이 없습니다.");
      if (!base64) throw new Error("파일 데이터가 없습니다.");
      if (!/\.pptx?$/i.test(origName)) throw new Error("PPT 또는 PPTX 파일만 업로드할 수 있습니다.");

      var cleanTitle = origName.replace(/\.pptx?$/i, "");
      var sourceMime = /\.pptx$/i.test(origName) ? MIME_PPTX : MIME_PPT;

      var folderU = DriveApp.getFolderById(searchFolderId);

      // 같은 이름의 구글 슬라이드가 이미 있는지 확인
      var dupIds = [];
      var dupFiles = folderU.searchFiles("title = '" + escapeQuery_(cleanTitle) + "' and mimeType = '" + MIME_GSLIDES + "'");
      while (dupFiles.hasNext()) dupIds.push(dupFiles.next().getId());

      // 덮어쓰기 지시가 없으면 변환하지 않고 사용자에게 확인을 넘긴다.
      if (dupIds.length > 0 && !overwrite) {
        return jsonOut_({ status: "success", duplicate: true, name: cleanTitle, id: dupIds[0] });
      }

      var bytes = Utilities.base64Decode(base64);
      var blob = Utilities.newBlob(bytes, sourceMime, origName);

      var created = Drive.Files.create({
        name: cleanTitle,
        mimeType: MIME_GSLIDES,
        parents: [searchFolderId]
      }, blob);

      // 변환에 성공한 뒤에 기존 파일을 정리한다. (실패 시 기존 파일이 사라지지 않도록 순서 유지)
      var trashed = 0;
      if (overwrite) {
        for (var d = 0; d < dupIds.length; d++) {
          try { DriveApp.getFileById(dupIds[d]).setTrashed(true); trashed++; }
          catch (eTrash) { Logger.log("기존 파일 삭제 실패: " + eTrash.toString()); }
        }
      }

      return jsonOut_({
        status: "success",
        duplicate: false,
        name: cleanTitle,
        id: created.id,
        replaced: trashed
      });
    }

    // --- [액션 D] 파일 자동 병합 및 즉시 다운로드 ---
    else if (action === "merge") {
      var fileIds = data.fileIds;
      var mergedName = data.mergedName || "통합_악보_PPT";
      var fontLabel = normalizeFontName_(data.fontName);      // 화면 표시용 이름
      var fontName = MERGE_FONT_MAP[fontLabel] || fontLabel;  // 실제 기록할 이름
      var fontSize = normalizeFontSize_(data.fontSize);       // null 이면 원본 크기 유지

      if (!fileIds || fileIds.length === 0) throw new Error("병합할 파일 ID가 없습니다.");

      var newPresentation = SlidesApp.create(mergedName);
      var newPresentationId = newPresentation.getId();
      var defaultSlide = newPresentation.getSlides()[0];

      // 1) 슬라이드 병합
      var appended = 0;
      for (var m = 0; m < fileIds.length; m++) {
        try {
          var sourcePresentation = SlidesApp.openById(fileIds[m]);
          var slides = sourcePresentation.getSlides();
          for (var s = 0; s < slides.length; s++) {
            newPresentation.appendSlide(slides[s]);
            appended++;
          }
        } catch (errFile) {
          // 파일 하나가 실패해도 전체 병합은 계속 진행한다.
          Logger.log("병합 실패 (" + fileIds[m] + "): " + errFile.toString());
        }
      }

      if (appended === 0) {
        DriveApp.getFileById(newPresentationId).setTrashed(true);
        throw new Error("병합된 슬라이드가 없습니다. 원본 파일 접근 권한을 확인해 주세요.");
      }

      // 2) 기본 빈 슬라이드 제거 (글꼴 적용 대상에서 제외하기 위해 먼저 삭제)
      defaultSlide.remove();

      // 3) 글꼴 / 글자 크기 일괄 적용 (슬라이드 + 발표자 노트 + 레이아웃 + 마스터)
      var fontReport = applyFontToPresentation_(newPresentation, fontName, fontSize);

      // 4) 변경 내용 확정 (이 호출 전에 export 하면 변경 전 상태가 내보내진다)
      newPresentation.saveAndClose();

      // 5) PPTX 내보내기
      var exportUrl = "https://docs.google.com/presentation/d/" + newPresentationId + "/export/pptx";
      var options = { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
      var response = UrlFetchApp.fetch(exportUrl, options);
      var code = response.getResponseCode();

      // 응답 코드를 확인하지 않으면 오류 페이지(HTML)가 그대로 pptx 로 저장되어 손상 파일이 된다.
      if (code !== 200) {
        DriveApp.getFileById(newPresentationId).setTrashed(true);
        throw new Error("PPTX 내보내기 실패 (HTTP " + code + ")");
      }

      var base64Data = Utilities.base64Encode(response.getBlob().getBytes());

      DriveApp.getFileById(newPresentationId).setTrashed(true);

      return jsonOut_({
        status: "success",
        fileName: mergedName + ".pptx",
        fileData: base64Data,
        fontApplied: fontLabel,
        fontSizeApplied: fontSize,
        fontReport: fontReport
      });
    }

    // --- 알 수 없는 액션 ---
    else {
      throw new Error("알 수 없는 요청입니다: " + action);
    }

  } catch (error) {
    return jsonOut_({
      status: "error",
      message: error.toString(),
      stack: (error && error.stack) ? String(error.stack).split("\n")[0] : ""
    });
  }
}

function doGet(e) {
  // 배포 상태 확인용
  return jsonOut_({
    status: "success",
    message: "웹앱이 정상 배포되었습니다.",
    fonts: MERGE_FONT_WHITELIST,
    fontSizeRange: [FONT_SIZE_MIN, FONT_SIZE_MAX]
  });
}

function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Drive 검색 쿼리에 들어가는 문자열의 역슬래시와 작은따옴표를 이스케이프한다. */
function escapeQuery_(text) {
  return String(text).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

// ==========================================
// 🔤 글꼴 / 글자 크기 일괄 적용 모듈
// ==========================================

/** 화이트리스트에 없는 글꼴이 들어오면 기본 글꼴로 강제한다. */
function normalizeFontName_(fontName) {
  if (!fontName) return MERGE_FONT_DEFAULT;
  var name = String(fontName).trim();
  for (var i = 0; i < MERGE_FONT_WHITELIST.length; i++) {
    if (MERGE_FONT_WHITELIST[i] === name) return name;
  }
  return MERGE_FONT_DEFAULT;
}

/** 허용 범위를 벗어나거나 값이 없으면 null 을 반환한다. (null = 원본 크기 유지) */
function normalizeFontSize_(fontSize) {
  if (fontSize === null || fontSize === undefined || fontSize === "") return null;
  var n = parseInt(fontSize, 10);
  if (isNaN(n)) return null;
  if (n < FONT_SIZE_MIN || n > FONT_SIZE_MAX) return null;
  return n;
}

/**
 * 프레젠테이션 전체의 글꼴과 글자 크기를 변경한다.
 * 슬라이드 본문뿐 아니라 발표자 노트 / 레이아웃 / 마스터까지 처리해야
 * 자리표시자 상속 텍스트의 서식이 원본으로 남는 문제를 막을 수 있다.
 *
 * @param {SlidesApp.Presentation} presentation  열려 있는 프레젠테이션 객체
 * @param {string} fontName  실제로 기록할 글꼴 이름
 * @param {number|null} fontSize  적용할 글자 크기(pt). null 이면 크기를 바꾸지 않는다.
 * @return {Object} 적용 결과 리포트
 */
function applyFontToPresentation_(presentation, fontName, fontSize) {
  var report = {
    font: fontName,
    fontSize: fontSize,
    shapes: 0,    // 서식을 적용한 도형 수
    cells: 0,     // 서식을 적용한 표 셀 수
    groups: 0,    // 진입한 그룹 수
    skipped: 0,   // 텍스트 스타일 API가 없어 건너뛴 개체 수 (워드아트 / 이미지화된 텍스트 등)
    errors: []    // 개체 단위 예외 (전체 중단 없이 수집)
  };

  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    applyFontToElements_(slides[i].getPageElements(), fontName, fontSize, report, "슬라이드" + (i + 1));

    try {
      var notes = slides[i].getNotesPage();
      if (notes) applyFontToElements_(notes.getPageElements(), fontName, fontSize, report, "노트" + (i + 1));
    } catch (eNote) {
      report.errors.push("노트" + (i + 1) + ": " + eNote.message);
    }
  }

  var layouts = presentation.getLayouts();
  for (var j = 0; j < layouts.length; j++) {
    applyFontToElements_(layouts[j].getPageElements(), fontName, fontSize, report, "레이아웃" + (j + 1));
  }

  var masters = presentation.getMasters();
  for (var k = 0; k < masters.length; k++) {
    applyFontToElements_(masters[k].getPageElements(), fontName, fontSize, report, "마스터" + (k + 1));
  }

  // errors 가 지나치게 길어지면 응답이 커지므로 앞부분만 남긴다.
  if (report.errors.length > 30) {
    report.errorCount = report.errors.length;
    report.errors = report.errors.slice(0, 30);
  }
  return report;
}

/**
 * 페이지 요소를 재귀적으로 순회하며 서식을 적용한다.
 * 개체 하나가 예외를 던져도 나머지 개체는 계속 처리한다.
 */
function applyFontToElements_(elements, fontName, fontSize, report, where) {
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    var type;

    // getPageElementType() 자체가 예외를 던지면 이전 코드에서는 병합 전체가 중단됐다.
    try {
      type = element.getPageElementType();
    } catch (eType) {
      report.errors.push(where + ": 개체 종류 확인 실패 - " + eType.message);
      continue;
    }

    try {
      if (type === SlidesApp.PageElementType.SHAPE) {
        if (forceFont_(element.asShape().getText(), fontName, fontSize, report, where)) report.shapes++;

      } else if (type === SlidesApp.PageElementType.TABLE) {
        var table = element.asTable();
        for (var r = 0; r < table.getNumRows(); r++) {
          for (var c = 0; c < table.getNumColumns(); c++) {
            // 셀 단위로 try 를 걸어야 병합 셀 하나 때문에 표 전체가 건너뛰어지지 않는다.
            try {
              var cell = table.getCell(r, c);
              if (cell.getMergeState() === SlidesApp.CellMergeState.MERGED) continue;
              if (forceFont_(cell.getText(), fontName, fontSize, report, where)) report.cells++;
            } catch (eCell) {
              report.errors.push(where + " 표(" + r + "," + c + "): " + eCell.message);
            }
          }
        }

      } else if (type === SlidesApp.PageElementType.GROUP) {
        report.groups++;
        applyFontToElements_(element.asGroup().getChildren(), fontName, fontSize, report, where + ">그룹");

      } else {
        // WORD_ART / IMAGE / VIDEO / LINE / SHEETS_CHART 등은 텍스트 스타일 API가 없다.
        report.skipped++;
      }
    } catch (eEl) {
      report.errors.push(where + " [" + type + "]: " + eEl.message);
    }
  }
}

/**
 * TextRange 전체 + run 단위로 글꼴과 글자 크기를 강제 적용한다.
 * 문단(getParagraphs) 단계는 run 과 동일한 텍스트를 중복 처리하므로 사용하지 않는다.
 * (API 호출 수가 줄어 6분 실행 제한 초과 위험이 낮아진다)
 *
 * @return {boolean} 적용 대상 텍스트가 실제로 있었는지 여부
 */
function forceFont_(textRange, fontName, fontSize, report, where) {
  if (!textRange) return false;

  var str = "";
  try {
    str = textRange.asString();
  } catch (eRead) {
    report.errors.push(where + ": 텍스트 읽기 실패 - " + eRead.message);
    return false;
  }
  if (!str || str.replace(/[\s\u000b]/g, "") === "") return false;

  // 1단계: 텍스트 상자 전체 범위
  try {
    var styleAll = textRange.getTextStyle();
    styleAll.setFontFamily(fontName);
    if (fontSize !== null) styleAll.setFontSize(fontSize);
  } catch (eAll) {
    report.errors.push(where + ": 전체 서식 적용 실패 - " + eAll.message);
  }

  // 2단계: run 단위 (굵게/기울임 등으로 서식이 나뉜 조각까지 덮어쓰기)
  try {
    var runs = textRange.getRuns();
    for (var k = 0; k < runs.length; k++) {
      try {
        var styleRun = runs[k].getTextStyle();
        styleRun.setFontFamily(fontName);
        if (fontSize !== null) styleRun.setFontSize(fontSize);
      } catch (eRun) {
        report.errors.push(where + " run" + k + ": " + eRun.message);
      }
    }
  } catch (eRuns) {
    report.errors.push(where + ": run 목록 조회 실패 - " + eRuns.message);
  }

  return true;
}

// ==========================================
// 🔎 [점검용] 특정 파일에 어떤 글꼴/크기가 남아 있는지 확인
//    편집기에서 파일 ID를 넣고 직접 실행 후 실행 로그를 확인한다.
// ==========================================
function checkFontApplied() {
  var presentationId = "여기에_슬라이드_파일_ID_입력";

  var pres = SlidesApp.openById(presentationId);
  var slides = pres.getSlides();
  var found = {};

  for (var i = 0; i < slides.length; i++) {
    var els = slides[i].getPageElements();
    for (var j = 0; j < els.length; j++) {
      try {
        if (els[j].getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
        var tr = els[j].asShape().getText();
        if (!tr.asString().trim()) continue;
        var f = tr.getTextStyle().getFontFamily();
        var sz = tr.getTextStyle().getFontSize();
        var key = (f === null ? "(혼합 또는 상속)" : f) + " / " + (sz === null ? "(혼합)" : sz + "pt");
        if (!found[key]) found[key] = [];
        found[key].push("슬라이드" + (i + 1));
      } catch (e) { /* 무시 */ }
    }
  }
  Logger.log(JSON.stringify(found, null, 2));
  return found;
}

// ==========================================
// 🚀 폴더 내 모든 .pptx 파일을 구글 슬라이드로 일괄 변환하는 매크로
// ==========================================
function convertAllPptxToGoogleSlides() {
  // ★ PPT 파일들이 들어있는 폴더 ID를 여기에 입력하세요.
  var targetFolderId = SEARCH_FOLDER_ID;

  var folder = DriveApp.getFolderById(targetFolderId);

  Logger.log("🔍 폴더 내에서 이미 변환된 구글 프레젠테이션 파일 목록을 수집 중입니다...");

  // 1. 이미 변환되어 존재하는 구글 프레젠테이션 파일들의 이름을 수집 (속도 최적화)
  var existingSlides = {};
  var slideFiles = folder.searchFiles("mimeType = '" + MIME_GSLIDES + "'");
  while (slideFiles.hasNext()) {
    existingSlides[slideFiles.next().getName()] = true;
  }

  Logger.log("✅ 수집 완료! 일괄 변환을 시작합니다...");

  var files = folder.getFiles();
  var count = 0;
  var skipCount = 0;

  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    var fileName = file.getName();

    // MS 파워포인트 파일 감지
    if (
      mimeType === MIME_PPTX ||
      mimeType === MIME_PPT ||
      fileName.toLowerCase().endsWith(".pptx") ||
      fileName.toLowerCase().endsWith(".ppt")
    ) {

      var cleanTitle = fileName.replace(/\.pptx?$/i, "");

      // ★ 핵심 해결책: 수집해둔 목록에 이름이 이미 존재한다면, 0.1초 만에 건너뜁니다!
      if (existingSlides[cleanTitle]) {
        Logger.log("⏭️ 건너뜀 (이미 변환됨): " + cleanTitle);
        skipCount++;
        continue;
      }

      // 변환 로직 시작
      try {
        var fileBlob = file.getBlob();
        fileBlob.setContentType(MIME_PPTX);

        var resource = {
          name: cleanTitle,
          mimeType: MIME_GSLIDES,
          parents: [targetFolderId]
        };

        // 구글 프레젠테이션으로 변환 생성
        Drive.Files.create(resource, fileBlob);

        Logger.log("✅ 변환 성공: " + fileName + " -> " + cleanTitle);
        count++;

        // 방금 변환한 파일도 목록에 추가하여 중복 방지
        existingSlides[cleanTitle] = true;

      } catch (err) {
        Logger.log("❌ 변환 실패 (" + fileName + "): " + err.toString());
      }
    }
  }

  Logger.log("🎉 작업 종료! (새로 변환됨: " + count + "개 / 건너뜀: " + skipCount + "개)");
}
