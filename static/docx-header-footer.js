/* docx-header-footer.js
 * Постобработка DOCX: верхний колонтитул в 2 строки (по левому краю),
 * нижний — две части в одной строке (лево + право) с живыми полями PAGE/NUMPAGES.
 * Лениво подгружает JSZip при первом вызове.
 * Публичный API: window.DocxHeaderFooter.addHeaderFooter(blob, options)
 */
(function (global) {
  'use strict';

  var JSZIP_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
  var jsZipPromise = null;

  // Ширина области табуляции в twips (1 inch = 1440 twips).
  // 9360 ≈ 6.5" — стандарт для Letter/A4 с полями 1".
  // Если поля другие — см. подсказку в конце статьи.
  var TAB_POS = 9360;

  // --- Ленивая загрузка JSZip ------------------------------------------------
  function loadJSZip() {
    if (global.JSZip) return Promise.resolve(global.JSZip);
    if (jsZipPromise) return jsZipPromise;

    jsZipPromise = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = JSZIP_CDN;
      s.async = true;
      s.onload = function () {
        global.JSZip ? resolve(global.JSZip)
                     : reject(new Error('JSZip загрузился, но не определён'));
      };
      s.onerror = function () { reject(new Error('Не удалось загрузить JSZip')); };
      document.head.appendChild(s);
    });

    return jsZipPromise;
  }

  // --- Утилиты ---------------------------------------------------------------
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function randId(prefix) {
    return prefix + Math.random().toString(36).slice(2, 8);
  }

  function rPr(opts) {
    opts = opts || {};
    return '<w:rPr>' +
      (opts.bold ? '<w:b/>' : '') +
      (opts.italic ? '<w:i/>' : '') +
      '<w:color w:val="' + (opts.color || '666666') + '"/>' +
      '<w:sz w:val="' + (opts.size || 18) + '"/>' +
      '<w:szCs w:val="' + (opts.size || 18) + '"/>' +
    '</w:rPr>';
  }

  // --- Строительные блоки runs ----------------------------------------------
  function run(text, opts) {
    return '<w:r>' + rPr(opts) +
      '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }

  function runBr(opts) {
    return '<w:r>' + rPr(opts) + '<w:br/></w:r>';
  }

  function runTab(opts) {
    return '<w:r>' + rPr(opts) + '<w:tab/></w:r>';
  }

  // Run с полем Word (PAGE, NUMPAGES и т.п.)
  // Между begin и end Word сам подставит значение. Значение «1»
  // в separate-части — это то, что увидит пользователь, если поле
  // не обновится (например, в старом viewer'е).
  function fieldRuns(instr, opts) {
    var r = rPr(opts);
    return '' +
      '<w:r>' + r + '<w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r>' + r + '<w:instrText xml:space="preserve"> ' + instr + ' </w:instrText></w:r>' +
      '<w:r>' + r + '<w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r>' + r + '<w:t>1</w:t></w:r>' +
      '<w:r>' + r + '<w:fldChar w:fldCharType="end"/></w:r>';
  }

  // --- Верхний колонтитул: 2 строки, по левому краю -------------------------
  function buildHeaderXml(lines) {
    var blocks = (lines || []).map(function (l) {
      return (typeof l === 'string') ? { text: l } : l;
    });

    var runsXml = '';
    blocks.forEach(function (b, i) {
      if (i > 0) runsXml += runBr(b);
      runsXml += run(b.text, b);
    });

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      '       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<w:p>' +
          '<w:pPr>' +
            '<w:jc w:val="left"/>' +
            '<w:pBdr>' +
              '<w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/>' +
            '</w:pBdr>' +
          '</w:pPr>' +
          runsXml +
        '</w:p>' +
      '</w:hdr>';
  }

  // --- Нижний колонтитул: слева текст, справа «стр. X из Y» -----------------
  function buildFooterXml(leftOpts) {
    var left = leftOpts || {};
    var rightStyle = { color: '888888', size: 16 };

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      '       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<w:p>' +
          '<w:pPr>' +
            // Правый таб-стоп у правого края полосы набора
            '<w:tabs>' +
              '<w:tab w:val="right" w:pos="' + TAB_POS + '"/>' +
            '</w:tabs>' +
            '<w:pBdr>' +
              '<w:top w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/>' +
            '</w:pBdr>' +
          '</w:pPr>' +
          // Левая часть — статический текст
          run(left.text || '', left) +
          // Табуляция до правого края
          runTab(left) +
          // Правая часть — «стр. X из Y» с живыми полями
          run('стр. ', rightStyle) +
          fieldRuns('PAGE', rightStyle) +
          run(' из ', rightStyle) +
          fieldRuns('NUMPAGES', rightStyle) +
        '</w:p>' +
      '</w:ftr>';
  }

  // --- Вставка headerReference/footerReference в sectPr ----------------------
  function injectSectPrRefs(documentXml, hdrRid, ftrRid) {
    if (documentXml.indexOf('<w:headerReference') !== -1) {
      return documentXml;
    }
    var idx = documentXml.lastIndexOf('<w:sectPr');
    if (idx === -1) return documentXml;

    var close = documentXml.indexOf('>', idx);
    if (close === -1) return documentXml;

    var refs =
      '<w:headerReference w:type="default" r:id="' + hdrRid + '"/>' +
      '<w:footerReference w:type="default" r:id="' + ftrRid + '"/>';

    return documentXml.slice(0, close + 1) + refs + documentXml.slice(close + 1);
  }

  // --- Основная функция ------------------------------------------------------
  async function addHeaderFooter(blob, options) {
    var opts = options || {};

    var headerLines = opts.headerLines || [];
    var footerLeft  = opts.footerLeft  || {};

    var JSZip = await loadJSZip();
    var zip = await JSZip.loadAsync(blob);

    var hdrRid = randId('rIdHdr');
    var ftrRid = randId('rIdFtr');

    // 1. Кладём XML колонтитулов
    zip.file('word/header1.xml', buildHeaderXml(headerLines));
    zip.file('word/footer1.xml', buildFooterXml(footerLeft));

    // 2. [Content_Types].xml
    var ctPath = '[Content_Types].xml';
    var ct = await zip.file(ctPath).async('string');
    if (ct.indexOf('/word/header1.xml') === -1) {
      ct = ct.replace(
        '</Types>',
        '<Override PartName="/word/header1.xml" ' +
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/footer1.xml" ' +
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '</Types>'
      );
    }
    zip.file(ctPath, ct);

    // 3. word/_rels/document.xml.rels
    var relsPath = 'word/_rels/document.xml.rels';
    var rels = await zip.file(relsPath).async('string');
    if (rels.indexOf('Target="header1.xml"') === -1) {
      rels = rels.replace(
        '</Relationships>',
        '<Relationship Id="' + hdrRid + '" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ' +
          'Target="header1.xml"/>' +
        '<Relationship Id="' + ftrRid + '" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" ' +
          'Target="footer1.xml"/>' +
        '</Relationships>'
      );
    }
    zip.file(relsPath, rels);

    // 4. word/document.xml
    var docPath = 'word/document.xml';
    var docXml = await zip.file(docPath).async('string');
    docXml = injectSectPrRefs(docXml, hdrRid, ftrRid);
    zip.file(docPath, docXml);

    // 5. Новый blob
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  // --- Экспорт ---------------------------------------------------------------
  global.DocxHeaderFooter = {
    addHeaderFooter: addHeaderFooter,
    loadJSZip: loadJSZip
  };
})(window);
