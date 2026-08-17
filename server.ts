import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

import { parseExcelOrTextMaster, isCleanJapaneseNameOrOffice } from "./src/excelParser";

dotenv.config();

const isProd = process.env.NODE_ENV === "production";
const PORT = 3000;

async function startServer() {
  const app = express();
  app.use(express.json({ limit: "10mb" }));

  // API Routes
  app.post("/api/import-ai", async (req, res) => {
    try {
      const { text } = req.body;
      if (!text || typeof text !== "string") {
        return res.status(400).json({ error: "テキストを入力してください。" });
      }

      // 1. Direct local TSV/Excel rule-based parse
      const localResult = parseExcelOrTextMaster(text);

      let aiResult: any = { offices: [], users: [], serviceCodes: [] };
      const apiKey = process.env.GEMINI_API_KEY;

      if (apiKey) {
        try {
          const ai = new GoogleGenAI({
            apiKey,
            httpOptions: {
              headers: {
                "User-Agent": "aistudio-build",
              },
            },
          });

          const prompt = `以下のテキストは、介護事業所の業務システム、Excel、スプレッドシート等から直接コピー＆ペーストされたマスターデータ（居宅介護支援事業所、担当ケアマネージャー、利用者、または介護サービスコードなど）です。
Tab（タブ）区切り、カンマ区切り、改行のみ、あるいは表形式（TSV/CSV）の不規則なフォーマットである可能性があります。
このデータを高度に解析・整形し、構造化されたJSONデータ（offices, users, serviceCodes）を抽出してください。

【解析対象テキスト】:
${text}

【最重要ルール：文字表記の100%完全保持】
- 事業所名、ケアマネ名、利用者名、サービスコードの漢字・ひらがな・カタカナ・英数字は、元のテキストの文字・表記を100%正確に維持してください。
- 絶対に勝手に「ひらがな（読み）」に変換したり（例：「大津」を「おおず」に変換等）、文字を削除・改変したり（例：「ケアプランセンターことは」を「ケアプランセンターとは」に改変等）しないでください。

【高度な抽出ルール】:
1. "offices":
   - 居宅介護支援事業所名 (officeName) と、その事業所に所属するケアマネージャー名 (careManagers) のリスト。
   - Excel等の行データで「事業所名 [Tab] ケアマネ名」や「事業所名, ケアマネ名」のような関係性がある場合、正確に紐付けてください。
   - もし関係性が明確でないテキストの箇条書きや単一列コピーの場合でも、事業所名（例：〜支援センター、〜プランニング）とケアマネージャーの氏名を可能な限り推測して紐付けるか、個別に抽出して officeName と careManagers に振り分けてください。
   - 重複する事業所名は1つにマージし、careManagers を配列にまとめてください。

2. "users":
   - 利用者様の氏名の配列。
   - 「様」「殿」「様分」「(様)」などの敬称、または「ID: 12345」「(要介護1)」などの余分な記号や属性値、内訳情報はすべて除外（クレンジング）し、「氏名（例：志波 啓子）」のみのクリーンな状態にしてください。
   - 明らかな事業所名やケアマネ名、サービスコードは含めないでください。

3. "serviceCodes":
   - 介護サービスコード、またはサービス内容の名称（例：「身体1」「生活2」「身体2生活1」「身2生1」「予防訪問介護」など）。
   - Excel のセルからコピーされた不要な文字（数字のみの羅列など関係のない値）は除外し、介護サービスの種類を明確に示す文字列のみにしてください。

4. その他の注意点:
   - テキスト内に存在しないカテゴリ、あるいは抽出できない場合は、null ではなく「空の配列 []」を返してください。`;

          const response = await ai.models.generateContent({
            model: "gemini-3.5-flash",
            contents: prompt,
            config: {
              responseMimeType: "application/json",
              responseSchema: {
                type: Type.OBJECT,
                properties: {
                  offices: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        officeName: { type: Type.STRING, description: "居宅介護支援事業所名" },
                        careManagers: {
                          type: Type.ARRAY,
                          items: { type: Type.STRING },
                          description: "所属する担当ケアマネージャー名の配列"
                        }
                      },
                      required: ["officeName", "careManagers"]
                    }
                  },
                  users: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                    description: "利用者名の配列（「様」などは含めない）"
                  },
                  serviceCodes: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                    description: "介護サービスコード/種類の配列（例：身体1、生活3）"
                  }
                },
                required: ["offices", "users", "serviceCodes"]
              }
            }
          });

          aiResult = JSON.parse(response.text || "{}");
        } catch (aiErr) {
          console.warn("Gemini AI API call failed, using local TSV parser fallback:", aiErr);
        }
      }

      // Merge localResult and aiResult
      const mergedOfficeMap = new Map<string, Set<string>>();

      const addOfficeToMap = (off: { officeName: string; careManagers: string[] }) => {
        if (!off || !off.officeName) return;
        const name = off.officeName.trim();
        if (!name || !isCleanJapaneseNameOrOffice(name)) return;
        if (!mergedOfficeMap.has(name)) {
          mergedOfficeMap.set(name, new Set());
        }
        (off.careManagers || []).forEach(cm => {
          if (cm && isCleanJapaneseNameOrOffice(cm)) mergedOfficeMap.get(name)!.add(cm.trim());
        });
      };

      (localResult.offices || []).forEach(addOfficeToMap);
      (aiResult.offices || []).forEach(addOfficeToMap);

      const mergedOffices = Array.from(mergedOfficeMap.entries()).map(([officeName, managersSet]) => ({
        officeName,
        careManagers: Array.from(managersSet)
      }));

      const mergedUsers = Array.from(new Set([
        ...(localResult.users || []),
        ...(aiResult.users || [])
      ])).filter(u => u && isCleanJapaneseNameOrOffice(u));

      const mergedServiceCodes = Array.from(new Set([
        ...(localResult.serviceCodes || []),
        ...(aiResult.serviceCodes || [])
      ])).filter(c => c && isCleanJapaneseNameOrOffice(c));

      res.json({
        offices: mergedOffices,
        users: mergedUsers,
        serviceCodes: mergedServiceCodes
      });
    } catch (error: any) {
      console.error("Import API Error:", error);
      res.status(500).json({ error: error.message || "解析中にエラーが発生しました。" });
    }
  });

  // Vite Integration
  if (!isProd) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running at http://localhost:${PORT}`);
  });
}

startServer();
