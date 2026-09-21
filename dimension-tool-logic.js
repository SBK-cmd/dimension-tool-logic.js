/*
  dimension-tool-logic.js  (remote-hosted ExtendScript logic)
  --------------------------------------------------------------
  DIMENSION_TOOL_VERSION = "1.3.0"

  Fetched fresh from the web by the Dimension Line Tool panel every
  time you click "สร้างเส้นบอกขนาด", then run inside Illustrator.
  Editing this file (or the GitHub file it is copied into) updates
  the tool immediately for anyone using the panel — no reinstall.

  Exposes createDimensionLines(paramsJSON), called right after this
  script is evaluated.
*/
var DIMENSION_TOOL_VERSION = "1.3.0";

// ---------- Thai strings used on the artboard / console (unicode-escaped for safety) ----------
var STR_ERR_NO_DOC = "กรุณาเปิดไฟล์งานก่อนใช้งานนี้";
var STR_ERR_NO_SEL = "กรุณาเลือกชิ้นงาน (object) ในอาร์ตบอร์ดก่อน แล้วกดปุ่มอีกครั้ง";
var STR_OK_DONE = "สร้างเส้นบอกขนาดเรียบร้อยแล้ว";
var STR_ERR_FONT = "ไม่พบฟอนต์ที่เลือก จะใช้ฟอนต์เริ่มต้นของโปรแกรมแทน";
var STR_LAYER_NAME = "เส้นบอกขนาด (Dimensions)";
var STR_GROUP_NAME = "ชุดเส้นบอกขนาด";

function createDimensionLines(paramsJSON) {
    try {
        if (app.documents.length === 0) {
            return "ERR:" + STR_ERR_NO_DOC;
        }
        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            return "ERR:" + STR_ERR_NO_SEL;
        }

        var p;
        try {
            p = eval("(" + paramsJSON + ")");
        } catch (parseErr) {
            return "ERR:invalid params (" + parseErr + ")";
        }

        var unit = p.unit || "cm";
        var gapMM = isFinite(p.gapMM) ? p.gapMM : 5;
        var r = clamp255(p.r), g = clamp255(p.g), b = clamp255(p.b);
        var fontName = p.font || "NotoSansThai-Regular";
        var fontSizeDim = (isFinite(p.fontSizeDim) && p.fontSizeDim > 0) ? p.fontSizeDim :
            ((isFinite(p.fontSize) && p.fontSize > 0) ? p.fontSize : 24);
        var fontSizeQty = (isFinite(p.fontSizeQty) && p.fontSizeQty > 0) ? p.fontSizeQty :
            ((isFinite(p.fontSize) && p.fontSize > 0) ? p.fontSize : 24);
        var bold = !!p.bold;
        var rotateVertical = !!p.rotateVertical;
        var prefixText = p.prefix !== undefined ? p.prefix : "";
        var qtyText = p.qty !== undefined ? String(p.qty) : "";
        var suffixText = p.suffix !== undefined ? p.suffix : "";

        // ---------- unit helpers ----------
        function ptToUnit(pt, u) {
            var inches = pt / 72;
            if (u === "mm") return inches * 25.4;
            if (u === "cm") return inches * 2.54;
            return inches;
        }
        function mmToPt(mm) { return (mm / 25.4) * 72; }
        function unitLabel(u) {
            if (u === "mm") return "mm";
            if (u === "cm") return "cm";
            return "in";
        }
        function formatNumber(n) {
            var rounded = Math.round(n * 100) / 100;
            var str = rounded.toFixed(2);
            str = str.replace(/0+$/, "").replace(/\.$/, "");
            return str;
        }
        function clamp255(v) {
            v = parseInt(v, 10);
            if (isNaN(v)) v = 0;
            return Math.min(255, Math.max(0, v));
        }

        // ---------- bounding box of current selection ----------
        var sel = doc.selection;
        var left = null, top = null, right = null, bottom = null;
        for (var i = 0; i < sel.length; i++) {
            var bb = sel[i].visibleBounds; // [left, top, right, bottom]
            if (left === null) {
                left = bb[0]; top = bb[1]; right = bb[2]; bottom = bb[3];
            } else {
                left = Math.min(left, bb[0]);
                top = Math.max(top, bb[1]);
                right = Math.max(right,
