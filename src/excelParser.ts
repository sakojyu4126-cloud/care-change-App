import * as XLSX from 'xlsx';

export interface ParsedMasterData {
  offices: Array<{ officeName: string; careManagers: string[] }>;
  users: string[];
  serviceCodes: string[];
}

/**
 * Checks if a string looks like binary garbled data (e.g. ZIP/XLSX binary header or unprintable chars)
 */
function isGarbageBinaryText(text: string): boolean {
  if (!text) return true;
  if (text.includes("PK\x03\x04") || text.includes("xl/workbook") || text.includes("rels/.rels") || text.includes("[Content_Types].xml")) {
    return true;
  }
  // Check ratio of replacement characters or non-printable ASCII
  let nonPrintableCount = 0;
  for (let i = 0; i < Math.min(text.length, 500); i++) {
    const code = text.charCodeAt(i);
    if ((code < 32 && code !== 9 && code !== 10 && code !== 13) || code === 65533) {
      nonPrintableCount++;
    }
  }
  return nonPrintableCount > 5;
}

/**
 * Helper to clean up any string from binary artifacts or excessive control characters
 */
export function isCleanJapaneseNameOrOffice(str: string): boolean {
  if (!str || str.length < 2 || str.length > 80) return false;
  if (isGarbageBinaryText(str)) return false;
  if (/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFD]/.test(str)) return false;
  if (str.includes("PK\x03\x04") || str.includes("xl/") || str.includes("xml") || str.includes("docProps") || str.includes("theme/")) return false;
  return true;
}

export function parseExcelOrTextMaster(rawText: string): ParsedMasterData {
  if (!rawText || !rawText.trim()) {
    return { offices: [], users: [], serviceCodes: [] };
  }

  let textToParse = rawText;

  // 1. If user pasted binary XLSX string directly, try parsing with XLSX library
  if (isGarbageBinaryText(rawText)) {
    try {
      const workbook = XLSX.read(rawText, { type: 'binary' });
      const sheetTexts: string[] = [];
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (sheet) {
          const csv = XLSX.utils.sheet_to_csv(sheet, { FS: '\t' });
          if (csv && csv.trim()) sheetTexts.push(csv);
        }
      });
      if (sheetTexts.length > 0) {
        textToParse = sheetTexts.join("\n");
      }
    } catch {
      // If binary parse failed, fallback to line-by-line filtering
    }
  }

  const officeMap = new Map<string, Set<string>>();
  const usersSet = new Set<string>();
  const serviceCodesSet = new Set<string>();

  // Helper to add office & manager
  const addOfficeManager = (officeName: string, managerName?: string) => {
    const cleanOffice = officeName.trim()
      .replace(/^[\s\t,;:"'「」【】]+/g, '')
      .replace(/[\s\t,;:"'「」【】]+$/g, '');
    if (!cleanOffice || !isCleanJapaneseNameOrOffice(cleanOffice)) return;
    
    if (!officeMap.has(cleanOffice)) {
      officeMap.set(cleanOffice, new Set<string>());
    }
    
    if (managerName) {
      const cleanManager = managerName.trim()
        .replace(/(様|殿|様分|マネージャー|ケアマネ|担当)+$/g, '')
        .replace(/^[\s\t,;:"'「」【】]+/g, '')
        .trim();
      if (cleanManager && isCleanJapaneseNameOrOffice(cleanManager)) {
        officeMap.get(cleanOffice)!.add(cleanManager);
      }
    }
  };

  // Helper to clean user name
  const addUserName = (name: string) => {
    let clean = name.trim()
      .replace(/(様|殿|様分|\(様\)|（様）)$/g, '')
      .replace(/^(利用者|氏名|患者|顧客)[:：\s]*/g, '')
      .replace(/[\(（][^\)）]*[\)）]/g, '') // remove brackets like (要介護1)
      .trim();
    if (clean && isCleanJapaneseNameOrOffice(clean) && !clean.includes("事業所") && !clean.includes("センター") && !clean.includes("コード")) {
      usersSet.add(clean);
    }
  };

  // Helper to add service code
  const addServiceCode = (code: string) => {
    let clean = code.trim().replace(/^(サービスコード|サービス)[:：\s]*/g, '').trim();
    if (clean && clean.length >= 2 && !isGarbageBinaryText(clean)) {
      serviceCodesSet.add(clean);
    }
  };

  const rawLines = textToParse.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  // Filter out any line that is binary garbage
  const lines = rawLines.filter(line => !isGarbageBinaryText(line));

  // Determine headers if first line looks like header row
  let officeColIdx = -1;
  let cmColIdx = -1;
  let userColIdx = -1;
  let codeColIdx = -1;

  if (lines.length > 0) {
    const firstLineCols = lines[0].split(/\t|,/).map(c => c.trim().toLowerCase());
    firstLineCols.forEach((col, idx) => {
      if (col.includes("事業所") || col.includes("居宅") || col.includes("オフィス")) officeColIdx = idx;
      else if (col.includes("マネ") || col.includes("担当") || col.includes("cm") || col.includes("ケア")) cmColIdx = idx;
      else if (col.includes("利用者") || col.includes("氏名") || col.includes("患者") || col.includes("名前")) userColIdx = idx;
      else if (col.includes("サービス") || col.includes("コード") || col.includes("内容")) codeColIdx = idx;
    });
  }

  const hasHeaderMatch = officeColIdx !== -1 || cmColIdx !== -1 || userColIdx !== -1 || codeColIdx !== -1;
  const startRowIdx = hasHeaderMatch ? 1 : 0;

  for (let i = startRowIdx; i < lines.length; i++) {
    const line = lines[i];

    // Check key-value explicit prefixes like 利用者：中島太郎 or サービスコード：身体1
    if (line.includes("利用者") || line.includes("サービス") || line.includes("事業所") || line.includes("ケアマネ")) {
      const userMatch = line.match(/(?:利用者|氏名|患者)[:：\s]*([^\n,、\t]+)/g);
      if (userMatch) {
        userMatch.forEach(m => {
          const name = m.replace(/(?:利用者|氏名|患者)[:：\s]*/, '');
          name.split(/[,、\s]+/).forEach(addUserName);
        });
      }

      const serviceMatch = line.match(/(?:サービス|コード)[:：\s]*([^\n,、\t]+)/g);
      if (serviceMatch) {
        serviceMatch.forEach(m => {
          const code = m.replace(/(?:サービス|コード)[:：\s]*/, '');
          code.split(/[,、\s]+/).forEach(addServiceCode);
        });
      }
    }

    // Tab or Comma split (Excel TSV/CSV)
    const cols = line.split(/\t|,/).map(c => c.trim()).filter(Boolean);

    if (cols.length >= 2) {
      if (hasHeaderMatch) {
        const off = officeColIdx !== -1 ? cols[officeColIdx] : "";
        const cm = cmColIdx !== -1 ? cols[cmColIdx] : "";
        const usr = userColIdx !== -1 ? cols[userColIdx] : "";
        const code = codeColIdx !== -1 ? cols[codeColIdx] : "";

        if (off) addOfficeManager(off, cm);
        if (usr) addUserName(usr);
        if (code) addServiceCode(code);
      } else {
        // Fallback TSV heuristic
        const col0 = cols[0];
        const col1 = cols[1];

        if (col0.includes("事業所") || col0.includes("センター") || col0.includes("プラン") || col0.includes("クリニック") || col0.includes("支援") || col0.includes("ステーション") || col0.includes("居宅") || col0.includes("大津") || col0.includes("山科")) {
          addOfficeManager(col0, col1);
        } else if (col0.match(/^(身体|生活|予防|身\d|生\d)/) || col1.match(/^(身体|生活|予防|身\d|生\d)/)) {
          if (col0.match(/^(身体|生活|予防|身\d|生\d)/)) addServiceCode(col0);
          if (col1.match(/^(身体|生活|予防|身\d|生\d)/)) addServiceCode(col1);
        } else {
          // Default 2-column TSV assumption: Office \t CareManager
          addOfficeManager(col0, col1);
        }

        if (cols.length >= 3) {
          const col2 = cols[2];
          if (col2.match(/^(身体|生活|予防)/)) {
            addServiceCode(col2);
          } else {
            addUserName(col2);
          }
        }
      }
    } else if (cols.length === 1) {
      const single = cols[0];
      if (single.includes("事業所") || single.includes("センター") || single.includes("プラン") || single.includes("居宅")) {
        addOfficeManager(single);
      } else if (single.match(/^(身体|生活|予防|身\d|生\d)/)) {
        addServiceCode(single);
      } else if (!single.includes("例") && !single.includes("マスター") && !single.includes("【") && !single.includes("解析")) {
        addUserName(single);
      }
    }
  }

  const officesArray = Array.from(officeMap.entries()).map(([officeName, managersSet]) => ({
    officeName,
    careManagers: Array.from(managersSet)
  }));

  return {
    offices: officesArray,
    users: Array.from(usersSet),
    serviceCodes: Array.from(serviceCodesSet)
  };
}

