import PizZip from 'pizzip';

/**
 * 核心修改邏輯：解析 docx 並替換答案 + 標紅
 * @param {ArrayBuffer} data - 上傳的 docx 檔案內容
 * @returns {Promise<Blob>} - 修改後的 docx Blob
 */
export async function modifyDocx(data) {
  const zip = new PizZip(data);
  const docXml = zip.file("word/document.xml").asText();
  
  // 使用 DOMParser 解析 XML
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXml, "text/xml");
  const paragraphs = xmlDoc.getElementsByTagName("w:p");

  // 解答資料庫
  const ans1 = ["O", "O", "X", "O", "O"]; // 是非題
  const ans2 = ["2", "3", "3", "2", "1"]; // 選擇題 (1: [2], 2: [3], 3: [3], 4: [2], 5: [1])

  let currentSection = 0; // 0: None, 1: 是非題, 2: 選擇題

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const text = p.textContent.trim();

    // 更強大的章節偵測
    if (text.includes("是非題") || text.includes("一.")) {
      currentSection = 1;
      console.log("Switched to Section 1 (是非題)");
    } else if (text.includes("選擇題") || text.includes("二.")) {
      currentSection = 2;
      console.log("Switched to Section 2 (選擇題)");
    }

    if (currentSection === 0) continue;

    // 匹配題號模式: 允許前方有空白，[ ] 括號 (包括全形)，以及 1-20 數字
    // Regex 改良：支援多種情況，並使用全域搜尋 /g 處理同段落多題
    const regex = /[\[［](.*?)[\]］]\s*(\d{1,2})\s*[.、\)]?/g;
    let match;
    
    while ((match = regex.exec(text)) !== null) {
      const qIndex = parseInt(match[2]);
      const bracketMatch = match[0]; // 整個匹配的字串，例如 "[O] 1."
      const matchIndex = match.index; // 在段落文字中的起始位置
      
      // 確保索引在範圍內
      if (qIndex > 0 && qIndex <= (currentSection === 1 ? ans1.length : ans2.length)) {
        const correctAns = currentSection === 1 ? ans1[qIndex - 1] : ans2[qIndex - 1];
        console.log(`Processing Section ${currentSection} Q${qIndex}: Found at ${matchIndex}, text: '${bracketMatch}', target: ${correctAns}`);
        applyAnswerToParagraph(p, correctAns, xmlDoc, matchIndex, bracketMatch);
      }
    }
  }

  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(xmlDoc);
  zip.file("word/document.xml", newXml);

  return zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

function applyAnswerToParagraph(para, answer, xmlDoc, matchStartIndex, fullMatchText) {
  const runs = para.getElementsByTagName("w:r");
  
  let currentOffset = 0;
  let firstBracketRun = null;
  let lastBracketRun = null;
  let bracketInParaOffset = -1;

  // 第一步：精確找出對應 matchStartIndex 的括號 Run
  for (let i = 0; i < runs.length; i++) {
    const r = runs[i];
    const tNodes = r.getElementsByTagName("w:t");
    if (tNodes.length === 0) continue;
    const t = tNodes[0];
    const runText = t.textContent;
    const runStart = currentOffset;
    const runEnd = currentOffset + runText.length;

    // 檢查 matchStartIndex 是否在這個 Run 或之後
    if (!firstBracketRun) {
      // 找到 "[" 的位置。在 fullMatchText 中第一個字就是 [ 或 ［
      const bracketOpenPosInPara = matchStartIndex; 
      if (bracketOpenPosInPara >= runStart && bracketOpenPosInPara < runEnd) {
        firstBracketRun = r;
      }
    }

    if (firstBracketRun && !lastBracketRun) {
      // 找到 "]" 的位置。在 fullMatchText 中找到第一個 ] 或 ］
      const closeMatch = fullMatchText.match(/[\]］]/);
      if (closeMatch) {
        const bracketClosePosInPara = matchStartIndex + closeMatch.index;
        if (bracketClosePosInPara >= runStart && bracketClosePosInPara < runEnd) {
          lastBracketRun = r;
        }
      }
    }

    currentOffset += runText.length;
    if (lastBracketRun) break;
  }

  if (firstBracketRun && lastBracketRun) {
    // 檢查現有內容
    // 從 fullMatchText 中擷取
    const openM = fullMatchText.match(/[\[［]/);
    const closeM = fullMatchText.match(/[\]］]/);
    const existingContent = (openM && closeM) ? fullMatchText.substring(openM.index + 1, closeM.index).trim() : "";
    
    const needsCorrection = (existingContent !== answer);
    console.log(`  Target Match Index: ${matchStartIndex}, Existing: '${existingContent}', Target: '${answer}', Needs Correction: ${needsCorrection}`);

    let foundFirst = false;
    let foundLast = false;

    for (let i = 0; i < runs.length; i++) {
      const r = runs[i];
      if (r === firstBracketRun) foundFirst = true;

      if (foundFirst && !foundLast) {
        if (needsCorrection) {
          let rPr = r.getElementsByTagName("w:rPr")[0];
          if (!rPr) {
            rPr = xmlDoc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:rPr");
            r.insertBefore(rPr, r.firstChild);
          }
          let color = rPr.getElementsByTagName("w:color")[0];
          if (!color) {
            color = xmlDoc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:color");
            rPr.appendChild(color);
          }
          color.setAttribute("w:val", "FF0000");
          color.removeAttribute("w:themeColor"); // 移除可能的主題色覆蓋
          color.removeAttribute("w:themeShade");
          let b = rPr.getElementsByTagName("w:b")[0];
          if (!b) {
            b = xmlDoc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:b");
            rPr.appendChild(b);
          }
        }

        const t = r.getElementsByTagName("w:t")[0];
        if (t && needsCorrection) {
          if (r === firstBracketRun) {
            const currentT = t.textContent;
            // 直接找 Run 裡的第一個 [ (因為已經透過 offset 鎖定 Run 了)
            const openMatch = currentT.match(/[\[［]/);
            const openIdx = openMatch ? openMatch.index : -1;
            const prefix = openIdx !== -1 ? currentT.substring(0, openIdx) : "";
            
            if (firstBracketRun === lastBracketRun) {
              const closeMatch = currentT.match(/[\]］]/);
              const closeIdx = closeMatch ? closeMatch.index : -1;
              const suffix = closeIdx !== -1 ? currentT.substring(closeIdx + 1) : "";
              t.textContent = prefix + `[${answer}]` + suffix;
            } else {
              t.textContent = prefix + `[${answer}]`;
            }
          } else if (r === lastBracketRun) {
            const currentT = t.textContent;
            const closeMatch = currentT.match(/[\]］]/);
            const closeIdx = closeMatch ? closeMatch.index : -1;
            t.textContent = closeIdx !== -1 ? currentT.substring(closeIdx + 1) : "";
          } else {
            t.textContent = "";
          }
        }
      }
      if (r === lastBracketRun) foundLast = true;
    }
  }
}
