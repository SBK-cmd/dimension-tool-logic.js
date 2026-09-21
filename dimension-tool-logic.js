/*
  dimension-tool-logic.js  (remote-hosted ExtendScript logic)
  --------------------------------------------------------------
  DIMENSION_TOOL_VERSION = "1.1.0"

  Fetched fresh from the web by the Dimension Line Tool panel every
  time you click "สร้างเส้นบอกขนาด", then run inside Illustrator.
  Editing this file (republishing this artifact) updates the tool
  for everyone using the panel — no reinstall needed.

  Exposes createDimensionLines(paramsJSON), called right after this
  script is evaluated.
*/
var DIMENSION_TOOL_VERSION = "1.1.0";

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
        var fontSize = (isFinite(p.fontSize) && p.fontSize > 0) ? p.fontSize : 24;
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
                right = Math.max(right, bb[2]);
                bottom = Math.min(bottom, bb[3]);
            }
        }
        var widthPt = right - left;
        var heightPt = top - bottom;
        var widthVal = formatNumber(ptToUnit(widthPt, unit)) + " " + unitLabel(unit);
        var heightVal = formatNumber(ptToUnit(heightPt, unit)) + " " + unitLabel(unit);

        var gapPt = mmToPt(gapMM);
        var capLen = mmToPt(4);     // length of the small perpendicular end-cap tick
        var textGap = mmToPt(2);

        // ---------- dimension layer ----------
        var dimLayer;
        try {
            dimLayer = doc.layers.getByName(STR_LAYER_NAME);
        } catch (e) {
            dimLayer = doc.layers.add();
            dimLayer.name = STR_LAYER_NAME;
        }
        dimLayer.locked = false;
        dimLayer.visible = true;

        var lineColor = new RGBColor();
        lineColor.red = r; lineColor.green = g; lineColor.blue = b;

        function setStroke(item) {
            item.stroked = true;
            item.strokeColor = lineColor;
            item.strokeWidth = 1;
            item.filled = false;
        }
        function setFillShape(item) {
            item.filled = true;
            item.fillColor = lineColor;
            item.stroked = false;
        }
        function addLine(x1, y1, x2, y2) {
            var path = dimLayer.pathItems.add();
            path.setEntirePath([[x1, y1], [x2, y2]]);
            setStroke(path);
            return path;
        }
        function addCap(x, y, dirX, dirY) {
            // dirX/dirY: unit vector along the dimension line at this end.
            // Draws a short tick perpendicular to the line, centered on (x, y).
            var perpX = -dirY, perpY = dirX;
            var half = capLen / 2;
            return addLine(x - perpX * half, y - perpY * half, x + perpX * half, y + perpY * half);
        }
        function addText(content, x, y, justification, rotateDeg) {
            var t = dimLayer.textFrames.add();
            t.contents = content;
            t.textRange.characterAttributes.size = fontSize;
            try {
                var f = app.textFonts.getByName(fontName);
                t.textRange.characterAttributes.textFont = f;
            } catch (e) {
                $.writeln(STR_ERR_FONT + " (" + fontName + ")");
            }
            if (bold) {
                try {
                    var famName = t.textRange.characterAttributes.textFont.family.name;
                    var boldFont = app.textFonts.getByName(famName + "-Bold");
                    t.textRange.characterAttributes.textFont = boldFont;
                } catch (e2) { /* ignore if bold face not found */ }
            }
            t.textRange.characterAttributes.fillColor = lineColor;
            t.contents = content;
            t.paragraphs[0].justification = justification;
            t.position = [x, y];
            if (rotateDeg) t.rotate(rotateDeg);
            return t;
        }

        var createdItems = [];

        // ===== WIDTH dimension (horizontal, above artwork) =====
        var wY = top + gapPt;
        createdItems.push(addLine(left, wY, right, wY));   // main horizontal dim line
        createdItems.push(addCap(left, wY, 1, 0));          // end-cap tick, left
        createdItems.push(addCap(right, wY, 1, 0));         // end-cap tick, right
        createdItems.push(addText(widthVal, (left + right) / 2, wY + textGap, Justification.CENTER, 0));

        // ===== HEIGHT dimension (vertical, right of artwork) =====
        var hX = right + gapPt;
        createdItems.push(addLine(hX, top, hX, bottom));    // main vertical dim line
        createdItems.push(addCap(hX, top, 0, 1));           // end-cap tick, top
        createdItems.push(addCap(hX, bottom, 0, 1));        // end-cap tick, bottom
        createdItems.push(addText(heightVal, hX + textGap, (top + bottom) / 2, Justification.LEFT,
            rotateVertical ? 90 : 0));

        // ===== QUANTITY text (centered below artwork) =====
        var qtyContent = (prefixText ? prefixText + " " : "") + qtyText + (suffixText ? " " + suffixText : "");
        var qY = bottom - gapPt - mmToPt(3);
        createdItems.push(addText(qtyContent, (left + right) / 2, qY, Justification.CENTER, 0));

        // ---------- group everything ----------
        var grp = dimLayer.groupItems.add();
        grp.name = STR_GROUP_NAME;
        for (var gi = createdItems.length - 1; gi >= 0; gi--) {
            createdItems[gi].move(grp, ElementPlacement.PLACEATBEGINNING);
        }

        doc.selection = null;
        grp.selected = true;
        app.redraw();

        return "OK:" + STR_OK_DONE + " (" + widthVal + " x " + heightVal + ") [v" + DIMENSION_TOOL_VERSION + "]";
    } catch (err) {
        return "ERR:" + err.toString();
    }
}
