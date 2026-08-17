import { useState, useEffect, useRef, useMemo } from "react";
import {
  Plus,
  Trash,
  Printer,
  Search,
  Sparkles,
  RefreshCw,
  FileText,
  Settings,
  Calendar,
  User,
  Home,
  ArrowRight,
  Edit2,
  Edit3,
  Copy,
  Save,
  CheckCircle,
  AlertCircle,
  Upload,
  Check,
  X
} from "lucide-react";
import * as XLSX from 'xlsx';
import { parseExcelOrTextMaster, isCleanJapaneseNameOrOffice } from "./excelParser";
import {
  CareManagerOffice,
  CareUser,
  SenderInfo,
  ServiceChangeItem,
  NoticeRecord
} from "./types";
import {
  DEFAULT_OFFICES,
  DEFAULT_USERS,
  DEFAULT_SERVICE_CODES,
  DEFAULT_SENDER_INFO
} from "./defaultData";

export default function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState<"create" | "history" | "master">("create");

  // Master Data State
  const [offices, setOffices] = useState<CareManagerOffice[]>([]);
  const [users, setUsers] = useState<CareUser[]>([]);
  const [serviceCodes, setServiceCodes] = useState<string[]>([]);
  const [senderInfo, setSenderInfo] = useState<SenderInfo>(DEFAULT_SENDER_INFO);

  // History State
  const [history, setHistory] = useState<NoticeRecord[]>([]);

  // AI Import State
  const [aiInputText, setAiInputText] = useState("");
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [aiSuccessMsg, setAiSuccessMsg] = useState("");
  const [aiErrorMsg, setAiErrorMsg] = useState("");

  // Create Form State
  const [editingId, setEditingId] = useState<string | null>(null);
  const [reportDate, setReportDate] = useState("");
  const [startMonth, setStartMonth] = useState("");
  const [startDay, setStartDay] = useState("");
  const [selectedOffice, setSelectedOffice] = useState("");
  const [selectedCareManager, setSelectedCareManager] = useState("");
  const [userNameInput, setUserNameInput] = useState("");
  const [showUserSuggest, setShowUserSuggest] = useState(false);
  const [generalDayOfWeek, setGeneralDayOfWeek] = useState("毎日");
  const [changes, setChanges] = useState<ServiceChangeItem[]>([]);
  
  // Notice Preview text overrides
  const [preGreeting, setPreGreeting] = useState("");
  const [postComment, setPostComment] = useState("");
  const [previewSender, setPreviewSender] = useState<SenderInfo>(DEFAULT_SENDER_INFO);

  // Manual Master Edit States
  const [newOfficeName, setNewOfficeName] = useState("");
  const [newCareManagerName, setNewCareManagerName] = useState("");
  const [selectedOfficeForManager, setSelectedOfficeForManager] = useState("");
  const [newUserName, setNewUserName] = useState("");
  const [newServiceCode, setNewServiceCode] = useState("");

  // Inline Master Editing States
  const [editingOfficeName, setEditingOfficeName] = useState<string | null>(null);
  const [editingOfficeNewVal, setEditingOfficeNewVal] = useState<string>("");
  const [editingCMKey, setEditingCMKey] = useState<string | null>(null);
  const [editingCMNewVal, setEditingCMNewVal] = useState<string>("");

  // Search Filter State
  const [searchQuery, setSearchQuery] = useState("");

  // Live TSV Excel Parse Memo for AI Import Preview
  const liveParsed = useMemo(() => {
    return parseExcelOrTextMaster(aiInputText);
  }, [aiInputText]);

  // Auto-resize comment textarea ref
  const commentTextareaRef = useRef<HTMLTextAreaElement>(null);

  // Load Initial Data from localStorage or defaults
  useEffect(() => {
    let storedOffices: string | null = null;
    let storedUsers: string | null = null;
    let storedServiceCodes: string | null = null;
    let storedSenderInfo: string | null = null;
    let storedHistory: string | null = null;

    try {
      storedOffices = localStorage.getItem("hvc_offices");
      storedUsers = localStorage.getItem("hvc_users");
      storedServiceCodes = localStorage.getItem("hvc_service_codes");
      storedSenderInfo = localStorage.getItem("hvc_sender_info");
      storedHistory = localStorage.getItem("hvc_history");
    } catch (e) {
      console.error("LocalStorage access blocked or errored:", e);
    }

    try {
      if (storedOffices) {
        const parsed = JSON.parse(storedOffices);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const officeMap = new Map<string, Set<string>>();
          parsed.forEach((o: any) => {
            if (!o || !o.officeName) return;
            let name = String(o.officeName).trim();
            if (!name || !isCleanJapaneseNameOrOffice(name)) return;
            
            // Auto-correct legacy or mangled office names
            if (name === "りんく大津" || name === "りんくおおず" || name === "おおず") {
              name = "大津居宅介護支援事業所";
            } else if (name === "ケアプランセンターとは") {
              name = "ケアプランセンターことは";
            }

            if (!officeMap.has(name)) {
              officeMap.set(name, new Set());
            }
            if (Array.isArray(o.careManagers)) {
              o.careManagers.forEach((cm: string) => {
                if (cm && isCleanJapaneseNameOrOffice(cm)) {
                  officeMap.get(name)!.add(String(cm).trim());
                }
              });
            }
          });

          const cleaned = Array.from(officeMap.entries()).map(([officeName, cms]) => ({
            officeName,
            careManagers: Array.from(cms)
          }));

          setOffices(cleaned.length > 0 ? cleaned : DEFAULT_OFFICES);
          localStorage.setItem("hvc_offices", JSON.stringify(cleaned.length > 0 ? cleaned : DEFAULT_OFFICES));
        } else {
          setOffices(DEFAULT_OFFICES);
        }
      } else {
        setOffices(DEFAULT_OFFICES);
        localStorage.setItem("hvc_offices", JSON.stringify(DEFAULT_OFFICES));
      }
    } catch (e) {
      setOffices(DEFAULT_OFFICES);
    }

    try {
      if (storedUsers) {
        const parsed = JSON.parse(storedUsers);
        if (Array.isArray(parsed)) {
          const cleaned = parsed.filter((u: any) => {
            const name = typeof u === "string" ? u : u?.name;
            return name && isCleanJapaneseNameOrOffice(name);
          });
          setUsers(cleaned.length > 0 ? cleaned : DEFAULT_USERS);
          localStorage.setItem("hvc_users", JSON.stringify(cleaned.length > 0 ? cleaned : DEFAULT_USERS));
        } else setUsers(DEFAULT_USERS);
      } else {
        setUsers(DEFAULT_USERS);
        localStorage.setItem("hvc_users", JSON.stringify(DEFAULT_USERS));
      }
    } catch (e) {
      setUsers(DEFAULT_USERS);
    }

    try {
      if (storedServiceCodes) {
        const parsed = JSON.parse(storedServiceCodes);
        if (Array.isArray(parsed)) setServiceCodes(parsed);
        else setServiceCodes(DEFAULT_SERVICE_CODES);
      } else {
        setServiceCodes(DEFAULT_SERVICE_CODES);
        localStorage.setItem("hvc_service_codes", JSON.stringify(DEFAULT_SERVICE_CODES));
      }
    } catch (e) {
      setServiceCodes(DEFAULT_SERVICE_CODES);
    }

    try {
      if (storedSenderInfo) {
        const parsed = JSON.parse(storedSenderInfo);
        if (parsed && typeof parsed === "object") {
          setSenderInfo(parsed);
          setPreviewSender(parsed);
        } else {
          setSenderInfo(DEFAULT_SENDER_INFO);
          setPreviewSender(DEFAULT_SENDER_INFO);
        }
      } else {
        setSenderInfo(DEFAULT_SENDER_INFO);
        setPreviewSender(DEFAULT_SENDER_INFO);
        localStorage.setItem("hvc_sender_info", JSON.stringify(DEFAULT_SENDER_INFO));
      }
    } catch (e) {
      setSenderInfo(DEFAULT_SENDER_INFO);
      setPreviewSender(DEFAULT_SENDER_INFO);
    }

    try {
      if (storedHistory) {
        const parsed = JSON.parse(storedHistory);
        if (Array.isArray(parsed)) setHistory(parsed);
      }
    } catch (e) {
      console.error("Error parsing history storage:", e);
    }

    // Calculate current Japanese Imperial Date (2026 is Reiwa 8)
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const date = today.getDate();
    // 2026 -> R8
    const reiwaYear = year - 2018;
    const dayOfWeekStr = ["日", "月", "火", "水", "木", "金", "土"][today.getDay()];
    setReportDate(`R${reiwaYear} 年 ${month} 月 ${date} 日 (${dayOfWeekStr})`);

    // Default start date next month or tomorrow
    const nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
    setStartMonth(String(nextWeek.getMonth() + 1));
    setStartDay(String(nextWeek.getDate()));
  }, []);

  // Sync Master Data to LocalStorage when changed
  const saveOfficesToStorage = (updated: CareManagerOffice[]) => {
    setOffices(updated);
    localStorage.setItem("hvc_offices", JSON.stringify(updated));
  };

  const saveUsersToStorage = (updated: CareUser[]) => {
    setUsers(updated);
    localStorage.setItem("hvc_users", JSON.stringify(updated));
  };

  const saveServiceCodesToStorage = (updated: string[]) => {
    setServiceCodes(updated);
    localStorage.setItem("hvc_service_codes", JSON.stringify(updated));
  };

  const saveSenderInfoToStorage = (updated: SenderInfo) => {
    setSenderInfo(updated);
    localStorage.setItem("hvc_sender_info", JSON.stringify(updated));
  };

  const saveHistoryToStorage = (updated: NoticeRecord[]) => {
    setHistory(updated);
    localStorage.setItem("hvc_history", JSON.stringify(updated));
  };

  // Initialize a fresh change row
  const createFreshChangeRow = (dayOfWeek = "毎日"): ServiceChangeItem => {
    return {
      id: Math.random().toString(36).substring(2, 9),
      dayOfWeek: dayOfWeek,
      dayOfWeekAfter: dayOfWeek,
      before: {
        startTime: "",
        endTime: "",
        serviceCode: "",
        customService: ""
      },
      after: {
        startTime: "",
        endTime: "",
        serviceCode: "",
        customService: "",
        remarks: ""
      }
    };
  };

  // Setup initial changes state if empty
  useEffect(() => {
    if (changes.length === 0) {
      setChanges([createFreshChangeRow(generalDayOfWeek)]);
    }
  }, []);

  // Update change items if generalDayOfWeek changes, only for those still matching default
  const handleGeneralDayOfWeekChange = (val: string) => {
    setGeneralDayOfWeek(val);
    setChanges(prev =>
      prev.map(item => {
        const updated = { ...item };
        if (item.dayOfWeek === generalDayOfWeek || !item.dayOfWeek) {
          updated.dayOfWeek = val;
        }
        if (item.dayOfWeekAfter === generalDayOfWeek || !item.dayOfWeekAfter) {
          updated.dayOfWeekAfter = val;
        }
        return updated;
      })
    );
  };

  // Handle office selection -> Auto select the first care manager
  const handleOfficeChange = (officeName: string) => {
    setSelectedOffice(officeName);
    const office = offices.find(o => o.officeName === officeName);
    if (office && office.careManagers.length > 0) {
      setSelectedCareManager(office.careManagers[0]);
    } else {
      setSelectedCareManager("");
    }
  };

  // Auto-generate greeting based on start date
  useEffect(() => {
    const greetingText = `平素はお世話になり有難うございます。
下記利用者様のサービス内容の変更を　${startMonth || "〇"}月　${startDay || "〇"}日以降に実施させていただきます。
お手数をおかけしますが、提供票の変更・修正を宜しくお願いいたします。`;
    setPreGreeting(greetingText);
  }, [startMonth, startDay]);

  // Adjust comment textarea height dynamically
  useEffect(() => {
    if (commentTextareaRef.current) {
      commentTextareaRef.current.style.height = "auto";
      commentTextareaRef.current.style.height = `${commentTextareaRef.current.scrollHeight}px`;
    }
  }, [postComment]);

  // Handle prediction of users
  const filteredUsersSuggest = useMemo(() => {
    if (!userNameInput.trim()) return [];
    return users.filter(u =>
      u.name.toLowerCase().includes(userNameInput.toLowerCase())
    );
  }, [userNameInput, users]);

  // Add a change row
  const addChangeRow = () => {
    setChanges(prev => [...prev, createFreshChangeRow(generalDayOfWeek)]);
  };

  // Remove a change row
  const removeChangeRow = (id: string) => {
    if (changes.length <= 1) {
      alert("少なくとも1つの変更箇所が必要です。");
      return;
    }
    setChanges(prev => prev.filter(c => c.id !== id));
  };

  // Update a field inside a specific change row
  const updateChangeRowField = (
    rowId: string,
    target: "before" | "after",
    field: string,
    value: string
  ) => {
    setChanges(prev =>
      prev.map(row => {
        if (row.id === rowId) {
          const currentTarget = row[target] || { startTime: "", endTime: "", serviceCode: "", customService: "", remarks: "" };
          return {
            ...row,
            [target]: {
              ...currentTarget,
              [field]: value
            }
          };
        }
        return row;
      })
    );
  };

  const updateChangeRowDayOfWeek = (rowId: string, value: string) => {
    setChanges(prev =>
      prev.map(row => (row.id === rowId ? { ...row, dayOfWeek: value } : row))
    );
  };

  const updateChangeRowDayOfWeekAfter = (rowId: string, value: string) => {
    setChanges(prev =>
      prev.map(row => (row.id === rowId ? { ...row, dayOfWeekAfter: value } : row))
    );
  };

  // Action: Save report to history
  const handleSaveReport = () => {
    if (!selectedOffice) {
      alert("居宅介護支援事業所を選択してください。");
      return;
    }
    if (!userNameInput.trim()) {
      alert("利用者名を入力してください。");
      return;
    }

    const newRecord: NoticeRecord = {
      id: editingId || Math.random().toString(36).substring(2, 9),
      createdAt: new Date().toISOString(),
      reportDate,
      startDate: `${startMonth}月 ${startDay}日`,
      officeName: selectedOffice,
      careManagerName: selectedCareManager,
      userName: userNameInput,
      senderInfo: previewSender,
      preGreeting,
      postComment,
      changes
    };

    let updatedHistory: NoticeRecord[] = [];
    if (editingId) {
      updatedHistory = history.map(h => (h.id === editingId ? newRecord : h));
      alert("報告書の更新を保存しました！");
    } else {
      updatedHistory = [newRecord, ...history];
      alert("新しい報告書を保存しました！「送信履歴・検索」タブから確認できます。");
    }

    saveHistoryToStorage(updatedHistory);
    // If the user name was not in master, optionally add it!
    if (!users.some(u => u.name.trim() === userNameInput.trim())) {
      const updatedUsers = [...users, { name: userNameInput.trim() }];
      saveUsersToStorage(updatedUsers);
    }
  };

  // Load report from history to edit
  const handleLoadRecord = (record: NoticeRecord) => {
    setEditingId(record.id);
    setReportDate(record.reportDate);
    
    // Parse startDate e.g. "6月 1日"
    const match = record.startDate.match(/(\d+)月\s*(\d+)日/);
    if (match) {
      setStartMonth(match[1]);
      setStartDay(match[2]);
    }

    setSelectedOffice(record.officeName);
    setSelectedCareManager(record.careManagerName);
    setUserNameInput(record.userName);
    setChanges(record.changes);
    setPreGreeting(record.preGreeting);
    setPostComment(record.postComment);
    if (record.senderInfo) {
      setPreviewSender(record.senderInfo);
    }
    setActiveTab("create");
  };

  // Copy report to use as a template for a new one
  const handleCopyRecord = (record: NoticeRecord) => {
    setEditingId(null); // Create new
    const today = new Date();
    const reiwaYear = today.getFullYear() - 2018;
    const dayOfWeekStr = ["日", "月", "火", "水", "木", "金", "土"][today.getDay()];
    setReportDate(`R${reiwaYear} 年 ${today.getMonth() + 1} 月 ${today.getDate()} 日 (${dayOfWeekStr})`);
    
    const match = record.startDate.match(/(\d+)月\s*(\d+)日/);
    if (match) {
      setStartMonth(match[1]);
      setStartDay(match[2]);
    }

    setSelectedOffice(record.officeName);
    setSelectedCareManager(record.careManagerName);
    setUserNameInput(record.userName);
    setChanges(record.changes.map(c => ({ ...c, id: Math.random().toString(36).substring(2, 9) })));
    setPreGreeting(record.preGreeting);
    setPostComment(record.postComment);
    if (record.senderInfo) {
      setPreviewSender(record.senderInfo);
    }
    setActiveTab("create");
    alert("履歴データをテンプレートとしてコピーしました。編集して保存してください。");
  };

  // Delete record
  const handleDeleteRecord = (id: string) => {
    if (confirm("この履歴を完全に削除してもよろしいですか？")) {
      const updated = history.filter(h => h.id !== id);
      saveHistoryToStorage(updated);
    }
  };

  // Clear/Reset Form
  const handleResetForm = () => {
    setEditingId(null);
    const today = new Date();
    const dayOfWeekStr = ["日", "月", "火", "水", "木", "金", "土"][today.getDay()];
    setReportDate(`R${today.getFullYear() - 2018} 年 ${today.getMonth() + 1} 月 ${today.getDate()} 日 (${dayOfWeekStr})`);
    setSelectedOffice(offices[0]?.officeName || "");
    setSelectedCareManager(offices[0]?.careManagers[0] || "");
    setUserNameInput("");
    setChanges([createFreshChangeRow(generalDayOfWeek)]);
    setPostComment("");
    setPreviewSender(senderInfo);
    // Refresh start date tomorrow
    const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
    setStartMonth(String(tomorrow.getMonth() + 1));
    setStartDay(String(tomorrow.getDate()));
  };

  // Action: Print / PDF output
  const handlePrint = () => {
    window.print();
  };

  // AI Import update handler
  const handleAiImportUpdate = async () => {
    if (!aiInputText.trim()) {
      setAiErrorMsg("解析するテキストを貼り付けてください。");
      return;
    }

    setIsAiLoading(true);
    setAiSuccessMsg("");
    setAiErrorMsg("");

    try {
      const response = await fetch("/api/import-ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: aiInputText })
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || "インポート解析中にエラーが発生しました。");
      }

      const data = await response.json();

      let addedOfficesCount = 0;
      let addedUsersCount = 0;
      let addedCodesCount = 0;

      // 1. Merge Offices
      const mergedOffices = [...offices];
      if (data.offices && Array.isArray(data.offices)) {
        data.offices.forEach((newOff: any) => {
          if (!newOff.officeName) return;
          const existing = mergedOffices.find(
            o => o.officeName.trim() === newOff.officeName.trim()
          );
          if (existing) {
            // merge care managers without duplicates
            const managersSet = new Set([
              ...existing.careManagers,
              ...(newOff.careManagers || [])
            ]);
            existing.careManagers = Array.from(managersSet).filter(Boolean);
          } else {
            mergedOffices.push({
              officeName: newOff.officeName.trim(),
              careManagers: (newOff.careManagers || []).filter(Boolean)
            });
            addedOfficesCount++;
          }
        });
      }

      // 2. Merge Users
      const mergedUsers = [...users];
      if (data.users && Array.isArray(data.users)) {
        data.users.forEach((userName: any) => {
          if (typeof userName !== "string") return;
          const cleanedName = userName.trim();
          if (cleanedName && !mergedUsers.some(u => u.name === cleanedName)) {
            mergedUsers.push({ name: cleanedName });
            addedUsersCount++;
          }
        });
      }

      // 3. Merge Service Codes
      const mergedCodes = [...serviceCodes];
      if (data.serviceCodes && Array.isArray(data.serviceCodes)) {
        data.serviceCodes.forEach((code: any) => {
          if (typeof code !== "string") return;
          const cleanedCode = code.trim();
          if (cleanedCode && !mergedCodes.includes(cleanedCode)) {
            mergedCodes.push(cleanedCode);
            addedCodesCount++;
          }
        });
      }

      // Save to states and localStorage
      saveOfficesToStorage(mergedOffices);
      saveUsersToStorage(mergedUsers);
      saveServiceCodesToStorage(mergedCodes);

      setAiSuccessMsg(
        `解析が完了しました！
・居宅事業所: ${addedOfficesCount}件を追加/更新
・利用者名: ${addedUsersCount}名を追加
・サービスコード: ${addedCodesCount}種類を追加`
      );
      setAiInputText("");
    } catch (err: any) {
      console.error(err);
      setAiErrorMsg(err.message || "テキストの解析に失敗しました。");
    } finally {
      setIsAiLoading(false);
    }
  };

  // Handle Excel (.xlsx, .xls) and text (.csv, .tsv, .txt) files
  const handleExcelOrTextFile = (file: File) => {
    const fileName = file.name.toLowerCase();
    const isExcel = fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || file.type.includes("sheet") || file.type.includes("excel");
    const reader = new FileReader();

    if (isExcel) {
      reader.onload = (evt) => {
        try {
          const buffer = evt.target?.result as ArrayBuffer;
          const workbook = XLSX.read(buffer, { type: 'array' });
          const sheetTexts: string[] = [];
          workbook.SheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            if (sheet) {
              const csv = XLSX.utils.sheet_to_csv(sheet, { FS: '\t' });
              if (csv && csv.trim()) sheetTexts.push(csv);
            }
          });
          if (sheetTexts.length > 0) {
            const text = sheetTexts.join("\n");
            setAiInputText(text);
            setAiSuccessMsg(`Excelファイル「${file.name}」を正常に読み込みました。下の「① Excelデータから即時反映」をクリックしてマスターに追加してください。`);
            setAiErrorMsg("");
          } else {
            setAiErrorMsg("Excelファイルにデータが見つかりませんでした。");
          }
        } catch (err) {
          console.error("XLSX read error:", err);
          setAiErrorMsg("Excelファイルの解析に失敗しました。ファイルが破損または保護されていないか確認してください。");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (evt) => {
        const content = evt.target?.result as string;
        if (content) {
          setAiInputText(content);
          setAiSuccessMsg(`ファイル「${file.name}」を読み込みました。`);
          setAiErrorMsg("");
        }
      };
      reader.readAsText(file);
    }
  };

  const handleAiInputClear = () => {
    setAiInputText("");
    setAiSuccessMsg("");
    setAiErrorMsg("");
  };

  // Direct TSV/Excel Rule-Based Import (Local, Instant)
  const handleDirectExcelImport = () => {
    if (!aiInputText.trim()) {
      setAiErrorMsg("インポートするテキストまたはファイル内容を貼り付けてください。");
      return;
    }

    const parsed = parseExcelOrTextMaster(aiInputText);

    let addedOfficesCount = 0;
    let addedUsersCount = 0;
    let addedCodesCount = 0;

    // Merge offices
    const mergedOffices = [...offices];
    parsed.offices.forEach((newOff) => {
      if (!newOff.officeName) return;
      const existing = mergedOffices.find(
        o => o.officeName.trim() === newOff.officeName.trim()
      );
      if (existing) {
        const managersSet = new Set([
          ...existing.careManagers,
          ...(newOff.careManagers || [])
        ]);
        existing.careManagers = Array.from(managersSet).filter(Boolean);
      } else {
        mergedOffices.push({
          officeName: newOff.officeName.trim(),
          careManagers: (newOff.careManagers || []).filter(Boolean)
        });
        addedOfficesCount++;
      }
    });

    // Merge users
    const mergedUsers = [...users];
    parsed.users.forEach((userName) => {
      const cleanedName = userName.trim();
      if (cleanedName && !mergedUsers.some(u => u.name === cleanedName)) {
        mergedUsers.push({ name: cleanedName });
        addedUsersCount++;
      }
    });

    // Merge service codes
    const mergedCodes = [...serviceCodes];
    parsed.serviceCodes.forEach((code) => {
      const cleanedCode = code.trim();
      if (cleanedCode && !mergedCodes.includes(cleanedCode)) {
        mergedCodes.push(cleanedCode);
        addedCodesCount++;
      }
    });

    saveOfficesToStorage(mergedOffices);
    saveUsersToStorage(mergedUsers);
    saveServiceCodesToStorage(mergedCodes);

    setAiSuccessMsg(
      `Excelデータの即時インポートが完了しました！\n・居宅事業所: ${addedOfficesCount}件を追加/更新\n・利用者名: ${addedUsersCount}名を追加\n・サービスコード: ${addedCodesCount}種類を追加`
    );
    setAiErrorMsg("");
  };

  // Office Inline Renaming
  const handleStartEditOffice = (oldName: string) => {
    setEditingOfficeName(oldName);
    setEditingOfficeNewVal(oldName);
  };

  const handleSaveEditOffice = (oldName: string) => {
    const trimmed = editingOfficeNewVal.trim();
    if (!trimmed) {
      setEditingOfficeName(null);
      return;
    }
    if (trimmed !== oldName && offices.some(o => o.officeName === trimmed)) {
      alert("すでに同じ名前の事業所が存在します。");
      return;
    }
    const updated = offices.map(o => {
      if (o.officeName === oldName) {
        return { ...o, officeName: trimmed };
      }
      return o;
    });
    saveOfficesToStorage(updated);
    if (selectedOffice === oldName) {
      setSelectedOffice(trimmed);
    }
    if (selectedOfficeForManager === oldName) {
      setSelectedOfficeForManager(trimmed);
    }
    setEditingOfficeName(null);
  };

  // Care Manager Inline Renaming
  const handleStartEditCareManager = (officeName: string, oldCM: string) => {
    setEditingCMKey(`${officeName}::${oldCM}`);
    setEditingCMNewVal(oldCM);
  };

  const handleSaveEditCareManager = (officeName: string, oldCM: string) => {
    const trimmed = editingCMNewVal.trim();
    if (!trimmed) {
      setEditingCMKey(null);
      return;
    }
    const updated = offices.map(o => {
      if (o.officeName === officeName) {
        const managers = o.careManagers.map(m => (m === oldCM ? trimmed : m));
        return { ...o, careManagers: managers };
      }
      return o;
    });
    saveOfficesToStorage(updated);
    if (selectedCareManager === oldCM) {
      setSelectedCareManager(trimmed);
    }
    setEditingCMKey(null);
  };

  // Manual Master Add Handlers
  const handleAddOffice = () => {
    if (!newOfficeName.trim()) return;
    if (offices.some(o => o.officeName === newOfficeName.trim())) {
      alert("すでに同じ名前の事業所が登録されています。");
      return;
    }
    const updated = [...offices, { officeName: newOfficeName.trim(), careManagers: [] }];
    saveOfficesToStorage(updated);
    setNewOfficeName("");
  };

  const handleAddCareManager = () => {
    if (!selectedOfficeForManager || !newCareManagerName.trim()) return;
    const updated = offices.map(o => {
      if (o.officeName === selectedOfficeForManager) {
        if (o.careManagers.includes(newCareManagerName.trim())) return o;
        return {
          ...o,
          careManagers: [...o.careManagers, newCareManagerName.trim()]
        };
      }
      return o;
    });
    saveOfficesToStorage(updated);
    setNewCareManagerName("");
  };

  const handleDeleteCareManager = (officeName: string, managerName: string) => {
    const updated = offices.map(o => {
      if (o.officeName === officeName) {
        return {
          ...o,
          careManagers: o.careManagers.filter(m => m !== managerName)
        };
      }
      return o;
    });
    saveOfficesToStorage(updated);
  };

  const handleDeleteOffice = (officeName: string) => {
    if (confirm(`「${officeName}」を完全に削除してもよろしいですか？`)) {
      const updated = offices.filter(o => o.officeName !== officeName);
      saveOfficesToStorage(updated);
    }
  };

  const handleAddUserManual = () => {
    if (!newUserName.trim()) return;
    if (users.some(u => u.name === newUserName.trim())) {
      alert("すでに登録されている利用者名です。");
      return;
    }
    const updated = [...users, { name: newUserName.trim() }];
    saveUsersToStorage(updated);
    setNewUserName("");
  };

  const handleDeleteUser = (name: string) => {
    const updated = users.filter(u => u.name !== name);
    saveUsersToStorage(updated);
  };

  const handleAddServiceCodeManual = () => {
    if (!newServiceCode.trim()) return;
    if (serviceCodes.includes(newServiceCode.trim())) {
      alert("すでに登録されているサービスコードです。");
      return;
    }
    const updated = [...serviceCodes, newServiceCode.trim()];
    saveServiceCodesToStorage(updated);
    setNewServiceCode("");
  };

  const handleDeleteServiceCode = (code: string) => {
    const updated = serviceCodes.filter(c => c !== code);
    saveServiceCodesToStorage(updated);
  };

  const handleUpdateSenderProfile = () => {
    saveSenderInfoToStorage(previewSender);
    alert("事業所の標準発信者情報を更新しました！今後の新規作成に適用されます。");
  };

  const handleResetMasterData = () => {
    if (confirm("マスターデータを初期状態（標準の事業所・ケアマネ・利用者・サービスコード）に復元・リセットしますか？\n（誤った表記や旧ダミーデータが削除され、標準状態に戻ります）")) {
      saveOfficesToStorage(DEFAULT_OFFICES);
      saveUsersToStorage(DEFAULT_USERS);
      saveServiceCodesToStorage(DEFAULT_SERVICE_CODES);
      alert("マスターデータを標準初期状態に復元しました。");
    }
  };

  // Search Logic
  const filteredHistory = useMemo(() => {
    if (!searchQuery.trim()) return history;
    const query = searchQuery.toLowerCase();
    return history.filter(
      h =>
        h.userName.toLowerCase().includes(query) ||
        h.officeName.toLowerCase().includes(query) ||
        h.careManagerName.toLowerCase().includes(query) ||
        h.startDate.toLowerCase().includes(query) ||
        h.reportDate.toLowerCase().includes(query)
    );
  }, [history, searchQuery]);

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800">
      {/* Top Navbar */}
      <header className="no-print bg-white border-b border-slate-200 sticky top-0 z-50 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between h-16 items-center">
            <div className="flex items-center space-x-3">
              <div className="p-2 bg-teal-600 text-white rounded-lg">
                <FileText className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-lg font-bold text-slate-950 leading-none">
                  訪問介護サービス提供内容変更通知
                </h1>
                <p className="text-xs text-slate-500 mt-1">桃の郷東山 - ケアマネ連絡報告システム</p>
              </div>
            </div>
            
            {/* Nav Links */}
            <nav className="flex space-x-1">
              <button
                onClick={() => setActiveTab("create")}
                className={`flex items-center space-x-1 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                  activeTab === "create"
                    ? "bg-teal-50 text-teal-700"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <Plus className="w-4 h-4" />
                <span>新規作成 / 編集</span>
              </button>
              <button
                onClick={() => setActiveTab("history")}
                className={`flex items-center space-x-1 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                  activeTab === "history"
                    ? "bg-teal-50 text-teal-700"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <Search className="w-4 h-4" />
                <span>送信履歴・検索</span>
              </button>
              <button
                onClick={() => setActiveTab("master")}
                className={`flex items-center space-x-1 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                  activeTab === "master"
                    ? "bg-teal-50 text-teal-700"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <Settings className="w-4 h-4" />
                <span>マスタ設定 & AIインポート</span>
              </button>
            </nav>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {/* ==================== CREATE TAB ==================== */}
        {activeTab === "create" && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            
            {/* Input Form Column */}
            <div className="lg:col-span-5 space-y-6 no-print">
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-5">
                <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                  <h2 className="text-md font-bold text-slate-900 flex items-center space-x-2">
                    <Edit2 className="w-4 h-4 text-teal-600" />
                    <span>変更内容 入力フォーム</span>
                  </h2>
                  {editingId ? (
                    <span className="text-xs bg-amber-50 text-amber-700 border border-amber-200 px-2 py-0.5 rounded-full font-medium">
                      履歴編集中
                    </span>
                  ) : (
                    <span className="text-xs bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded-full font-medium">
                      新規作成
                    </span>
                  )}
                </div>

                {/* Report Date */}
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    報告年月日 (和暦)
                  </label>
                  <input
                    type="text"
                    value={reportDate}
                    onChange={(e) => setReportDate(e.target.value)}
                    className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-hidden"
                    placeholder="例: R8 年 5 月 28 日"
                  />
                </div>

                {/* Care Manager Office Select */}
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    居宅介護支援事業所名
                  </label>
                  <select
                    value={selectedOffice}
                    onChange={(e) => handleOfficeChange(e.target.value)}
                    className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 bg-white focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-hidden"
                  >
                    <option value="">-- 選択してください --</option>
                    {offices.map((o) => (
                      <option key={o.officeName} value={o.officeName}>
                        {o.officeName}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Care Manager Name Select */}
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    担当ケアマネージャー様
                  </label>
                  <select
                    value={selectedCareManager}
                    onChange={(e) => setSelectedCareManager(e.target.value)}
                    className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 bg-white focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-hidden"
                    disabled={!selectedOffice}
                  >
                    <option value="">-- 選択してください --</option>
                    {selectedOffice &&
                      offices
                        .find((o) => o.officeName === selectedOffice)
                        ?.careManagers.map((cm) => (
                          <option key={cm} value={cm}>
                            {cm}
                          </option>
                        ))}
                  </select>
                  {!selectedOffice && (
                    <p className="text-[10px] text-slate-400 mt-1">※居宅介護支援事業所を選択すると候補が表示されます。</p>
                  )}
                </div>

                {/* User Name (Register Prediction list) */}
                <div className="relative">
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    利用者様氏名
                  </label>
                  <div className="relative">
                    <input
                      type="text"
                      value={userNameInput}
                      onChange={(e) => {
                        setUserNameInput(e.target.value);
                        setShowUserSuggest(true);
                      }}
                      onFocus={() => setShowUserSuggest(true)}
                      onBlur={() => setTimeout(() => setShowUserSuggest(false), 200)}
                      autoComplete="off"
                      className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-hidden"
                      placeholder="利用者名を入力（登録リストから予測）"
                    />
                    <User className="absolute right-3 top-2.5 w-4 h-4 text-slate-400" />
                  </div>
                  {/* Autosuggest prediction box */}
                  {showUserSuggest && filteredUsersSuggest.length > 0 && (
                    <div className="absolute z-10 w-full bg-white border border-slate-200 mt-1 rounded-lg shadow-lg max-h-40 overflow-y-auto">
                      {filteredUsersSuggest.map((u) => (
                        <button
                          key={u.name}
                          onMouseDown={() => {
                            setUserNameInput(u.name);
                            setShowUserSuggest(false);
                          }}
                          className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 border-b border-slate-100 last:border-0"
                        >
                          {u.name}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                {/* Start Date & Day of Week */}
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      実施開始 〇月 〇日
                    </label>
                    <div className="flex items-center space-x-1">
                      <input
                        type="number"
                        value={startMonth}
                        onChange={(e) => setStartMonth(e.target.value)}
                        className="w-full text-center text-sm border border-slate-300 rounded-lg py-2 outline-hidden focus:ring-2 focus:ring-teal-500/20"
                        placeholder="月"
                        min="1"
                        max="12"
                      />
                      <span className="text-xs text-slate-500">月</span>
                      <input
                        type="number"
                        value={startDay}
                        onChange={(e) => setStartDay(e.target.value)}
                        className="w-full text-center text-sm border border-slate-300 rounded-lg py-2 outline-hidden focus:ring-2 focus:ring-teal-500/20"
                        placeholder="日"
                        min="1"
                        max="31"
                      />
                      <span className="text-xs text-slate-500">日</span>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      標準の曜日
                    </label>
                    <input
                      type="text"
                      value={generalDayOfWeek}
                      onChange={(e) => handleGeneralDayOfWeekChange(e.target.value)}
                      className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 outline-hidden focus:ring-2 focus:ring-teal-500/20"
                      placeholder="例: 毎日、月・水"
                    />
                  </div>
                </div>

              </div>

              {/* Dynamic Service Changes Section */}
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-4">
                <div className="flex justify-between items-center border-b border-slate-100 pb-2">
                  <h3 className="text-sm font-bold text-slate-900">
                    予定変更・サービス内容 (行数自動追加)
                  </h3>
                  <button
                    onClick={addChangeRow}
                    className="flex items-center space-x-1 text-xs text-teal-600 hover:text-teal-700 font-semibold"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>変更箇所を追加</span>
                  </button>
                </div>

                {changes.map((row, index) => (
                  <div key={row.id} className="p-4 bg-slate-50 rounded-lg border border-slate-100 space-y-3 relative">
                    <div className="flex justify-between items-center">
                      <span className="text-xs font-bold text-slate-500 bg-slate-200/60 px-2 py-0.5 rounded-sm">
                        変更パターン {index + 1}
                      </span>
                      {changes.length > 1 && (
                        <button
                          onClick={() => removeChangeRow(row.id)}
                          className="text-slate-400 hover:text-rose-600 transition-colors"
                          title="この行を削除"
                        >
                          <Trash className="w-4 h-4" />
                        </button>
                      )}
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-500 mb-0.5">特記事項・備考 (変更後用)</label>
                      <input
                        type="text"
                        value={row.after.remarks || ""}
                        onChange={(e) => updateChangeRowField(row.id, "after", "remarks", e.target.value)}
                        className="w-full text-xs border border-slate-300 rounded-md p-1 outline-hidden"
                        placeholder="例: 時間変更のみ。内容は同一。"
                      />
                    </div>

                    {/* Before & After Times */}
                    <div className="grid grid-cols-2 gap-4">
                      {/* Before (変更前) */}
                      <div className="bg-amber-50/40 p-2 rounded border border-amber-100 space-y-1.5">
                        <div className="flex justify-between items-center">
                          <span className="text-[10px] font-bold text-amber-700 bg-amber-100/50 px-1.5 py-0.2 rounded-xs">
                            変更前
                          </span>
                          <button
                            type="button"
                            onClick={() => {
                              updateChangeRowField(row.id, "before", "startTime", "");
                              updateChangeRowField(row.id, "before", "endTime", "");
                            }}
                            className="text-[9px] text-amber-600 hover:text-amber-800 hover:underline font-semibold"
                            title="開始・終了時刻を空欄にして追加支援として設定します"
                          >
                            空欄にする（追加支援用）
                          </button>
                        </div>

                        <div className="flex space-x-1 items-center">
                          <span className="text-[9px] text-amber-800 font-bold w-7">曜日:</span>
                          <input
                            type="text"
                            value={row.dayOfWeek}
                            onChange={(e) => updateChangeRowDayOfWeek(row.id, e.target.value)}
                            className="w-full text-xs border border-slate-300 rounded p-1 outline-hidden bg-white"
                            placeholder="例: 月・木"
                          />
                        </div>

                        <div className="flex space-x-1 items-center">
                          <span className="text-[9px] text-amber-800 font-bold w-7">時間:</span>
                          <input
                            type="text"
                            value={row.before?.startTime || ""}
                            onChange={(e) => updateChangeRowField(row.id, "before", "startTime", e.target.value)}
                            className="w-full text-xs border border-slate-300 rounded p-1 text-center outline-hidden bg-white focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
                            placeholder="例: 07:00"
                          />
                          <span className="text-slate-400 text-xs self-center">~</span>
                          <input
                            type="text"
                            value={row.before?.endTime || ""}
                            onChange={(e) => updateChangeRowField(row.id, "before", "endTime", e.target.value)}
                            className="w-full text-xs border border-slate-300 rounded p-1 text-center outline-hidden bg-white focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
                            placeholder="例: 07:20"
                          />
                        </div>
                        
                        <div className="pl-8">
                          <label className="block text-[9px] font-bold text-amber-800 mb-0.5">サービス内容 (直接入力可)</label>
                          <input
                            type="text"
                            list={`form-dl-before-${row.id}`}
                            value={row.before?.serviceCode || row.before?.customService || ""}
                            onChange={(e) => {
                              updateChangeRowField(row.id, "before", "serviceCode", e.target.value);
                              updateChangeRowField(row.id, "before", "customService", e.target.value);
                            }}
                            className="w-full text-xs border border-slate-300 rounded p-1.5 bg-white outline-hidden focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
                            placeholder="例: 身体01、身体1、生活2など"
                          />
                          <datalist id={`form-dl-before-${row.id}`}>
                            {serviceCodes.map((code) => (
                              <option key={code} value={code} />
                            ))}
                          </datalist>
                        </div>
                      </div>

                      {/* After (変更後) */}
                      <div className="bg-emerald-50/40 p-2 rounded border border-emerald-100 space-y-1.5">
                        <span className="text-[10px] font-bold text-emerald-700 bg-emerald-100/50 px-1.5 py-0.2 rounded-xs inline-block">
                          変更後
                        </span>

                        <div className="flex space-x-1 items-center">
                          <span className="text-[9px] text-emerald-800 font-bold w-7">曜日:</span>
                          <input
                            type="text"
                            value={row.dayOfWeekAfter !== undefined ? row.dayOfWeekAfter : row.dayOfWeek}
                            onChange={(e) => updateChangeRowDayOfWeekAfter(row.id, e.target.value)}
                            className="w-full text-xs border border-slate-300 rounded p-1 outline-hidden bg-white focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                            placeholder="例: 月・木"
                          />
                        </div>

                        <div className="flex space-x-1 items-center">
                          <span className="text-[9px] text-emerald-800 font-bold w-7">時間:</span>
                          <input
                            type="text"
                            value={row.after?.startTime || ""}
                            onChange={(e) => updateChangeRowField(row.id, "after", "startTime", e.target.value)}
                            className="w-full text-xs border border-slate-300 rounded p-1 text-center outline-hidden bg-white focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                            placeholder="例: 07:20"
                          />
                          <span className="text-slate-400 text-xs self-center">~</span>
                          <input
                            type="text"
                            value={row.after?.endTime || ""}
                            onChange={(e) => updateChangeRowField(row.id, "after", "endTime", e.target.value)}
                            className="w-full text-xs border border-slate-300 rounded p-1 text-center outline-hidden bg-white focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                            placeholder="例: 07:40"
                          />
                        </div>

                        <div className="pl-8">
                          <label className="block text-[9px] font-bold text-emerald-800 mb-0.5">サービス内容 (直接入力可)</label>
                          <input
                            type="text"
                            list={`form-dl-after-${row.id}`}
                            value={row.after?.serviceCode || row.after?.customService || ""}
                            onChange={(e) => {
                              updateChangeRowField(row.id, "after", "serviceCode", e.target.value);
                              updateChangeRowField(row.id, "after", "customService", e.target.value);
                            }}
                            className="w-full text-xs border border-slate-300 rounded p-1.5 bg-white outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                            placeholder="例: 身体01、身体1、生活2など"
                          />
                          <datalist id={`form-dl-after-${row.id}`}>
                            {serviceCodes.map((code) => (
                              <option key={code} value={code} />
                            ))}
                          </datalist>
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Extra Comments & Notes (Manual hand-written) */}
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-3">
                <label className="block text-xs font-semibold text-slate-700">
                  連絡事項 / 状況ご報告 (手入力・文字数で自動行高さ調整)
                </label>
                <textarea
                  value={postComment}
                  onChange={(e) => setPostComment(e.target.value)}
                  className="w-full text-sm border border-slate-300 rounded-lg p-3 min-h-[100px] outline-hidden focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500"
                  placeholder="例: デイサービスにて一時体調を崩されました。緊急対応の上、18時前に桃の郷へ戻られました。その後のご様子は安定されております。"
                />
              </div>

              {/* Reset & Save Toolbar */}
              <div className="flex space-x-3">
                <button
                  onClick={handleResetForm}
                  className="flex-1 py-3 border border-slate-300 text-slate-700 rounded-xl text-sm font-semibold hover:bg-slate-50 transition-colors bg-white shadow-xs"
                >
                  フォームをクリア
                </button>
                <button
                  onClick={handleSaveReport}
                  className="flex-1 py-3 bg-teal-600 text-white rounded-xl text-sm font-semibold hover:bg-teal-700 transition-colors shadow-sm flex items-center justify-center space-x-2"
                >
                  <Save className="w-4 h-4" />
                  <span>この内容を履歴保存</span>
                </button>
              </div>
            </div>

            {/* LIVE EXCEL-STYLE PREVIEW COLUMN */}
            <div className="lg:col-span-7 print-container">
              
              {/* Instructions on Preview Mode */}
              <div className="no-print bg-amber-50 text-amber-800 p-4 rounded-xl border border-amber-100 flex items-start space-x-3 mb-4">
                <AlertCircle className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
                <div className="text-xs space-y-1">
                  <p className="font-semibold">【直接プレビュー編集機能】</p>
                  <p>右側のエクセル形式プレビューに表示される<strong>定型文（平素はお世話になり...）、日付、発信者情報、コメント等</strong>は、印刷前の最終微調整としてその場で直接手書き修正可能です。</p>
                </div>
              </div>

              {/* Notice Sheet A4 Paper Wrapper */}
              <div className="bg-white p-8 rounded-xl border border-slate-300 shadow-md print-container min-h-[842px] max-w-4xl mx-auto flex flex-col justify-between" id="report-sheet">
                
                {/* Main Content of Sheet */}
                <div className="space-y-4">
                  {/* Print Date Row */}
                  <div className="flex justify-end items-start text-xs font-normal font-hgpgothic text-slate-500">
                    {/* Interactive Report Date (Editable) */}
                    <div className="flex items-center space-x-1">
                      <span>報告日: </span>
                      <input
                        type="text"
                        value={reportDate}
                        onChange={(e) => setReportDate(e.target.value)}
                        className="text-right border-0 font-normal text-slate-800 bg-transparent outline-hidden px-1 py-0.5 w-48 font-meiryoui text-[11px]"
                        placeholder=""
                      />
                    </div>
                  </div>

                  {/* Title (Underline removed as requested) */}
                  <div className="text-center pb-2">
                    <h1 className="text-sm md:text-[15px] font-normal tracking-widest text-slate-900 inline-block px-12 font-hgpgothic">
                      訪問介護サービス提供内容変更のお知らせ
                    </h1>
                  </div>

                  {/* Destinations & Sender Info side-by-side */}
                  <div className="grid grid-cols-12 gap-4 items-stretch font-hgpgothic font-normal">
                    
                    {/* Destinations Block (Left) - Unboxed header, boxed details like requested */}
                    <div className="col-span-6 flex flex-col justify-start">
                      <div className="text-[10px] text-slate-500 font-normal mb-1">送信先：</div>
                      <table className="w-full border border-slate-900 border-collapse font-hgpgothic font-normal text-slate-900 text-[12px]">
                        <tbody>
                          <tr className="border-b border-slate-900">
                            <td className="w-[45%] border-r border-slate-900 p-1.5 text-center bg-slate-50/20 whitespace-nowrap text-[10px] text-slate-700 font-hgpgothic font-normal">
                              居宅介護支援事業所名
                            </td>
                            <td className="w-[55%] p-1.5 text-center whitespace-nowrap overflow-hidden text-ellipsis font-hgpgothic font-normal text-[11px]">
                              {selectedOffice || "（事業所を選択）"}
                            </td>
                          </tr>
                          <tr>
                            <td className="border-r border-slate-900 p-1.5 text-center bg-slate-50/20 whitespace-nowrap text-[10px] text-slate-700 font-hgpgothic font-normal">
                              担当ケアマネージャー様
                            </td>
                            <td className="p-1.5 whitespace-nowrap font-hgpgothic font-normal text-[11px]">
                              <div className="flex justify-between items-center px-4 w-full">
                                <span className="flex-1 text-center overflow-hidden text-ellipsis">
                                  {selectedCareManager || "（ケアマネを選択）"}
                                </span>
                                <span className="text-slate-900 flex-shrink-0 select-none">様</span>
                              </div>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Sender Info Block (Right - Editable) */}
                    <div className="col-span-6 border border-slate-900 p-2 text-[10px] leading-relaxed space-y-0.5 relative group font-hgpgothic font-normal flex flex-col justify-between">
                      <div className="text-right font-normal text-[9px] text-teal-600 no-print absolute right-2 top-1.5 opacity-0 group-hover:opacity-100 transition-opacity">
                        直接編集可能
                      </div>
                      <div className="flex items-center">
                        <span className="text-slate-500 w-14 text-[9px]">発信者:</span>
                        <span className="text-slate-900 text-[11px] font-hgpgothic font-normal">{previewSender.stationName}</span>
                      </div>
                      <div className="text-slate-700 flex items-center">
                        <span className="w-14 text-[9px] text-slate-500">郵便番号:</span>
                        <span className="font-meiryoui text-[10px]">〒 </span>
                        <input
                          type="text"
                          value={previewSender.postalCode}
                          onChange={(e) => setPreviewSender({ ...previewSender, postalCode: e.target.value })}
                          className="border-0 outline-hidden w-20 bg-transparent font-meiryoui ml-1 text-[10px]"
                          placeholder=""
                        />
                      </div>
                      <div className="text-slate-700 flex items-center">
                        <span className="w-14 text-[9px] text-slate-500">住所:</span>
                        <input
                          type="text"
                          value={previewSender.address}
                          onChange={(e) => setPreviewSender({ ...previewSender, address: e.target.value })}
                          className="border-0 outline-hidden flex-1 bg-transparent text-[10px]"
                          placeholder=""
                        />
                      </div>
                      <div className="text-slate-700 flex items-center">
                        <span className="w-14 text-[9px] text-slate-500">TEL:</span>
                        <input
                          type="text"
                          value={previewSender.tel}
                          onChange={(e) => setPreviewSender({ ...previewSender, tel: e.target.value })}
                          className="border-0 outline-hidden flex-1 bg-transparent font-meiryoui text-[10px]"
                          placeholder=""
                        />
                      </div>
                      <div className="text-slate-700 flex items-center">
                        <span className="w-14 text-[9px] text-slate-500">FAX:</span>
                        <input
                          type="text"
                          value={previewSender.fax}
                          onChange={(e) => setPreviewSender({ ...previewSender, fax: e.target.value })}
                          className="border-0 outline-hidden flex-1 bg-transparent font-meiryoui text-[10px]"
                          placeholder=""
                        />
                      </div>
                      <div className="text-slate-700 pt-0.5 flex justify-between items-center">
                        <span className="flex items-center">
                          <span className="text-[9px] text-slate-500 mr-1">担当者:</span>
                          <input
                            type="text"
                            value={previewSender.picName}
                            onChange={(e) => setPreviewSender({ ...previewSender, picName: e.target.value })}
                            className="border-0 outline-hidden w-20 bg-transparent font-hgpgothic font-normal text-slate-900 text-[10px]"
                            placeholder=""
                          />
                        </span>
                        <button
                          onClick={handleUpdateSenderProfile}
                          className="no-print text-[9px] bg-teal-50 text-teal-700 hover:bg-teal-100 border border-teal-200 px-1 py-0.2 rounded-sm transition-colors font-hgpgothic font-normal"
                          title="これを標準テンプレートとして保存"
                        >
                          標準保存
                        </button>
                      </div>
                    </div>

                  </div>

                  {/* Intro Greeting Paragraph (Highly editable live textarea) */}
                  <div className="relative group">
                    <div className="text-[9px] text-teal-600 absolute right-2 -top-4 opacity-0 group-hover:opacity-100 transition-opacity no-print font-bold">
                      挨拶定型文は以下を直接編集できます
                    </div>
                    <textarea
                      value={preGreeting}
                      onChange={(e) => setPreGreeting(e.target.value)}
                      className="w-full text-xs text-slate-800 font-medium font-udkyokasho leading-relaxed border-0 p-2 bg-transparent rounded-sm transition-colors outline-hidden resize-y min-h-[60px]"
                      placeholder=""
                    />
                  </div>

                  {/* User Name Block (Dotted underline removed) */}
                  <div className="w-[30%] border border-slate-900 px-2 py-1 bg-slate-50/20 font-hgpgothic font-normal text-slate-800">
                    <div className="flex items-center whitespace-nowrap overflow-hidden">
                      <span className="text-[10px] text-slate-500 mr-1 flex-shrink-0">利用者名：</span>
                      <input
                        type="text"
                        value={userNameInput}
                        onChange={(e) => setUserNameInput(e.target.value)}
                        autoComplete="off"
                        className="text-[13px] text-slate-900 font-hgpgothic font-normal border-0 outline-hidden bg-transparent w-full px-0.5"
                        placeholder=""
                      />
                      <span className="text-slate-900 ml-1 flex-shrink-0">様</span>
                    </div>
                  </div>

                  {/* Visual Headers "変更前" / "変更後" */}
                  <div className="flex justify-between items-center text-xs font-normal pt-1 font-hgpgothic text-slate-800">
                    <div className="w-[30%] bg-slate-100 text-center py-1 border border-slate-900 rounded-xs tracking-wider">
                      変更前
                    </div>
                    <div className="w-[70%] bg-emerald-50/20 text-center py-1 border border-slate-900 rounded-xs text-emerald-900 tracking-wider">
                      {changes.some(c => !(c.before?.startTime || "").trim() && !(c.before?.endTime || "").trim()) 
                        ? "変更後 (追加支援)" 
                        : "変更後 (重要なお知らせ)"}
                    </div>
                  </div>

                  {/* Dynamic Side-By-Side Care Sheets List */}
                  <div className="space-y-4">
                    {changes.map((item, index) => {
                      const isBeforeEmpty = !(item.before?.startTime || "").trim() && !(item.before?.endTime || "").trim();
                      return (
                        <div key={item.id} className="flex justify-between items-stretch print-avoid-break">
                          
                          {/* Before Block (Width adjusted to 30% for narrower display) */}
                          <div className="w-[30%] border border-slate-900 text-xs flex flex-col justify-between font-hgpgothic font-normal">
                            <div className="grid grid-cols-12 border-b border-slate-900 bg-slate-100 text-[9px] text-slate-700">
                              <div className="col-span-3 border-r border-slate-900 font-normal p-1 text-center">
                                曜日
                              </div>
                              <div className="col-span-9 font-normal p-1 text-center">
                                サービス内容
                              </div>
                            </div>
                            
                            <div className="grid grid-cols-12 flex-1">
                              {/* Day (Minimized padding and space) */}
                              <div className="col-span-3 border-r border-slate-900 flex items-center justify-center p-0.5 text-center text-[10px] text-slate-800 font-hgpgothic font-normal break-words">
                                <input
                                  type="text"
                                  value={item.dayOfWeek || ""}
                                  onChange={(e) => updateChangeRowDayOfWeek(item.id, e.target.value)}
                                  className="w-full text-center bg-transparent border-0 focus:bg-amber-50 outline-hidden p-0.5 text-[10px] font-hgpgothic"
                                  placeholder=""
                                />
                              </div>
                              {/* Service Content */}
                              <div className="col-span-9 p-1 space-y-0.5 flex flex-col justify-center text-center">
                                {isBeforeEmpty ? (
                                  <div className="text-[10px] text-slate-400 font-hgpgothic font-normal">
                                    （なし）
                                  </div>
                                ) : (
                                  <>
                                    <div className="text-[10px] font-normal tracking-tight font-meiryoui text-slate-800 flex items-center justify-center space-x-0.5">
                                      <input
                                        type="text"
                                        value={item.before?.startTime || ""}
                                        onChange={(e) => updateChangeRowField(item.id, "before", "startTime", e.target.value)}
                                        className="w-11 text-center bg-transparent border-0 outline-hidden p-0 text-[10px] font-meiryoui"
                                        placeholder=""
                                      />
                                      <span>〜</span>
                                      <input
                                        type="text"
                                        value={item.before?.endTime || ""}
                                        onChange={(e) => updateChangeRowField(item.id, "before", "endTime", e.target.value)}
                                        className="w-11 text-center bg-transparent border-0 outline-hidden p-0 text-[10px] font-meiryoui"
                                        placeholder=""
                                      />
                                    </div>
                                    <div className="text-[9px] text-slate-700 bg-slate-50/50 border border-slate-200 py-0.2 rounded-xs px-0.5 mx-auto max-w-[95%]">
                                      <input
                                        type="text"
                                        list={`sheet-dl-before-${item.id}`}
                                        value={item.before?.serviceCode || item.before?.customService || ""}
                                        onChange={(e) => {
                                          updateChangeRowField(item.id, "before", "serviceCode", e.target.value);
                                          updateChangeRowField(item.id, "before", "customService", e.target.value);
                                        }}
                                        className="w-full text-center bg-transparent border-0 outline-hidden p-0 text-[9px] font-hgpgothic"
                                        placeholder=""
                                      />
                                      <datalist id={`sheet-dl-before-${item.id}`}>
                                        {serviceCodes.map((code) => (
                                          <option key={code} value={code} />
                                        ))}
                                      </datalist>
                                    </div>
                                  </>
                                )}
                              </div>
                            </div>
                          </div>

                          {/* After Block (Width increased to 70% for maximum readability on changes) */}
                          <div className="w-[70%] border border-slate-900 text-xs flex flex-col justify-between font-hgpgothic font-normal">
                            <div className="grid grid-cols-12 border-b border-slate-900 bg-slate-100 text-[9px] text-slate-700">
                              <div className="col-span-2 border-r border-slate-900 font-normal p-1 text-center">
                                曜日
                              </div>
                              <div className="col-span-3 border-r border-slate-900 font-normal p-1 text-center">
                                {isBeforeEmpty ? "追加支援内容" : "サービス内容 (変更後)"}
                              </div>
                              <div className="col-span-7 font-normal p-1 text-center">
                                備考・特記事項
                              </div>
                            </div>

                            <div className="grid grid-cols-12 flex-1">
                              {/* Day (Minimized padding and space) */}
                              <div className="col-span-2 border-r border-slate-900 flex items-center justify-center p-0.5 text-center text-[10px] text-slate-800 font-hgpgothic font-normal break-words">
                                <input
                                  type="text"
                                  value={item.dayOfWeekAfter !== undefined ? item.dayOfWeekAfter : (item.dayOfWeek || "")}
                                  onChange={(e) => updateChangeRowDayOfWeekAfter(item.id, e.target.value)}
                                  className="w-full text-center bg-transparent border-0 focus:bg-emerald-50 outline-hidden p-0.5 text-[10px] font-hgpgothic"
                                  placeholder=""
                                />
                              </div>
                              {/* Service Content */}
                              <div className="col-span-3 border-r border-slate-900 p-1 space-y-0.5 flex flex-col justify-center text-center font-hgpgothic font-normal">
                                <div className="text-[11px] font-normal font-meiryoui text-emerald-950 flex items-center justify-center space-x-0.5">
                                  <input
                                    type="text"
                                    value={item.after?.startTime || ""}
                                    onChange={(e) => updateChangeRowField(item.id, "after", "startTime", e.target.value)}
                                    className="w-12 text-center bg-transparent border-0 outline-hidden p-0 text-[11px] font-meiryoui text-emerald-950 font-medium"
                                    placeholder=""
                                  />
                                  <span>〜</span>
                                  <input
                                    type="text"
                                    value={item.after?.endTime || ""}
                                    onChange={(e) => updateChangeRowField(item.id, "after", "endTime", e.target.value)}
                                    className="w-12 text-center bg-transparent border-0 outline-hidden p-0 text-[11px] font-meiryoui text-emerald-950 font-medium"
                                    placeholder=""
                                  />
                                </div>
                                <div className="text-[10px] font-normal text-emerald-900 bg-emerald-50/50 border border-emerald-100 py-0.2 rounded-xs px-0.5 mx-auto max-w-[95%]">
                                  <input
                                    type="text"
                                    list={`sheet-dl-after-${item.id}`}
                                    value={item.after?.serviceCode || item.after?.customService || ""}
                                    onChange={(e) => {
                                      updateChangeRowField(item.id, "after", "serviceCode", e.target.value);
                                      updateChangeRowField(item.id, "after", "customService", e.target.value);
                                    }}
                                    className="w-full text-center bg-transparent border-0 outline-hidden p-0 text-[10px] font-hgpgothic text-emerald-900 font-medium"
                                    placeholder=""
                                  />
                                  <datalist id={`sheet-dl-after-${item.id}`}>
                                    {serviceCodes.map((code) => (
                                      <option key={code} value={code} />
                                    ))}
                                  </datalist>
                                </div>
                              </div>
                              {/* Remarks */}
                              <div className="col-span-7 p-1.5 flex items-center text-slate-800 leading-normal text-[10px] font-normal font-udkyokasho">
                                <input
                                  type="text"
                                  value={item.after?.remarks || ""}
                                  onChange={(e) => updateChangeRowField(item.id, "after", "remarks", e.target.value)}
                                  className="w-full bg-transparent border-0 outline-hidden p-0.5 text-[10px] font-udkyokasho text-slate-800"
                                  placeholder=""
                                />
                              </div>
                            </div>
                          </div>

                        </div>
                      );
                    })}
                  </div>

                  {/* Hand-written Instruction Comment Block */}
                  <div className="border-2 border-slate-900 p-4 text-xs font-sans leading-relaxed print-avoid-break">
                    <div className="font-bold mb-2 text-slate-900 text-[11px]">連絡事項（詳細ご報告）:</div>
                    <div className="relative group">
                      <div className="text-[9px] text-teal-600 absolute right-1 -top-4 opacity-0 group-hover:opacity-100 transition-opacity no-print font-bold">
                        以下を直接編集して最終微調整できます
                      </div>
                      <textarea
                        ref={commentTextareaRef}
                        value={postComment}
                        onChange={(e) => setPostComment(e.target.value)}
                        className="w-full text-xs text-slate-800 font-medium font-udkyokasho leading-relaxed bg-slate-50/55 hover:bg-white border-0 focus:bg-white rounded-sm p-2 outline-hidden resize-none overflow-hidden"
                        placeholder=""
                      />
                    </div>
                  </div>

                </div>

              </div>

              {/* Notice Action Toolbar */}
              <div className="no-print mt-4 max-w-4xl mx-auto flex justify-end space-x-3">
                <button
                  onClick={handlePrint}
                  className="px-6 py-3 bg-teal-600 text-white rounded-xl text-sm font-bold hover:bg-teal-700 transition-colors shadow-md flex items-center space-x-2"
                >
                  <Printer className="w-4 h-4" />
                  <span>PDF保存 ・ 印刷出力</span>
                </button>
              </div>

            </div>

          </div>
        )}

        {/* ==================== HISTORY SEARCH TAB ==================== */}
        {activeTab === "history" && (
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-6 space-y-6">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <h2 className="text-md font-bold text-slate-900">送信報告書 履歴一覧・検索</h2>
                <p className="text-xs text-slate-500 mt-1">
                  過去に作成・保存した報告書を検索、再編集、または再印刷できます。
                </p>
              </div>

              {/* Search Bar */}
              <div className="relative w-full md:w-80">
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full text-sm border border-slate-300 rounded-lg pl-9 pr-4 py-2 focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-hidden bg-slate-50"
                  placeholder="利用者名・ケアマネ・日付等で検索..."
                />
                <Search className="absolute left-3 top-2.5 w-4 h-4 text-slate-400" />
              </div>
            </div>

            {filteredHistory.length === 0 ? (
              <div className="text-center py-16 text-slate-400 space-y-2">
                <FileText className="w-12 h-12 mx-auto stroke-[1.5]" />
                <p className="text-sm">該当する保存履歴が見つかりません。</p>
                <button
                  onClick={() => setActiveTab("create")}
                  className="text-xs text-teal-600 hover:underline font-semibold"
                >
                  新しい報告書を作成する
                </button>
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 text-slate-500 font-semibold border-b border-slate-200">
                      <th className="p-3.5">報告年月日 (和暦)</th>
                      <th className="p-3.5">利用者様氏名</th>
                      <th className="p-3.5">居宅介護支援事業所 / ケアマネ</th>
                      <th className="p-3.5">実施開始日</th>
                      <th className="p-3.5 text-center">変更行数</th>
                      <th className="p-3.5 text-right">操作</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredHistory.map((item) => (
                      <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                        <td className="p-3.5 font-medium text-slate-900">{item.reportDate}</td>
                        <td className="p-3.5 font-semibold text-teal-800">{item.userName} 様</td>
                        <td className="p-3.5">
                          <div className="text-xs font-semibold text-slate-800">{item.officeName}</div>
                          <div className="text-[11px] text-slate-500 mt-0.5">
                            担当CM: {item.careManagerName} 様
                          </div>
                        </td>
                        <td className="p-3.5">
                          <span className="bg-teal-50 text-teal-700 text-xs px-2 py-0.5 rounded font-semibold">
                            {item.startDate}〜
                          </span>
                        </td>
                        <td className="p-3.5 text-center font-mono font-bold text-slate-600">
                          {item.changes.length} 行
                        </td>
                        <td className="p-3.5 text-right space-x-1.5">
                          <button
                            onClick={() => handleLoadRecord(item)}
                            className="inline-flex items-center space-x-1 text-xs bg-slate-100 text-slate-700 hover:bg-slate-200 px-2.5 py-1 rounded font-semibold transition-colors"
                            title="この報告書を編集する"
                          >
                            <Edit2 className="w-3 h-3" />
                            <span>編集</span>
                          </button>
                          <button
                            onClick={() => handleCopyRecord(item)}
                            className="inline-flex items-center space-x-1 text-xs bg-teal-50 text-teal-700 hover:bg-teal-100 px-2.5 py-1 rounded font-semibold transition-colors"
                            title="これをテンプレートとして新規コピー"
                          >
                            <Copy className="w-3 h-3" />
                            <span>コピー作成</span>
                          </button>
                          <button
                            onClick={() => handleDeleteRecord(item.id)}
                            className="inline-flex items-center text-xs bg-rose-50 text-rose-700 hover:bg-rose-100 p-1 rounded transition-colors"
                            title="履歴から削除"
                          >
                            <Trash className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* ==================== MASTER TAB & AI IMPORT ==================== */}
        {activeTab === "master" && (
          <div className="space-y-8">
            
            {/* AI & Excel Data Import Panel */}
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3 flex-wrap gap-2">
                <div className="flex items-center space-x-2 text-md font-bold text-slate-900">
                  <Sparkles className="w-5 h-5 text-teal-600" />
                  <span>Excel・スプレッドシート・AIデータインポート</span>
                </div>

                <div className="flex items-center space-x-2 flex-wrap gap-2">
                  {/* Paste from Clipboard Button */}
                  <button
                    type="button"
                    onClick={async () => {
                      try {
                        const text = await navigator.clipboard.readText();
                        if (text && text.trim()) {
                          setAiInputText(text);
                          setAiSuccessMsg("クリップボードからデータを貼り付けました！下の「① 即時追加」または「② AI解析追加」をクリックしてください。");
                          setAiErrorMsg("");
                        } else {
                          setAiErrorMsg("クリップボードにテキストデータが見つかりませんでした。先にExcelからセル範囲をコピー（Ctrl+C）してください。");
                        }
                      } catch (err) {
                        setAiErrorMsg("下のテキスト枠をクリックし、直接 Ctrl+V （ペースト）して貼り付けてください。");
                      }
                    }}
                    className="inline-flex items-center space-x-1 text-xs bg-teal-50 hover:bg-teal-100 text-teal-800 px-3 py-1.5 rounded-lg font-semibold cursor-pointer border border-teal-300 transition-colors"
                    title="クリップボードのコピーデータを自動で枠内に貼り付けます"
                  >
                    <Copy className="w-3.5 h-3.5 text-teal-600" />
                    <span>クリップボードから貼り付け (ペースト)</span>
                  </button>

                  {/* File Upload Button */}
                  <label className="inline-flex items-center space-x-1 text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-lg font-semibold cursor-pointer border border-slate-300 transition-colors">
                    <Upload className="w-3.5 h-3.5 text-slate-500" />
                    <span>Excel/CSV/テキストファイルを読み込む (.xlsx, .csv, .txt)</span>
                    <input
                      type="file"
                      accept=".csv,.tsv,.txt,.xlsx,.xls"
                      className="hidden"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) handleExcelOrTextFile(file);
                      }}
                    />
                  </label>
                </div>
              </div>

              <p className="text-xs text-slate-500 leading-relaxed">
                Excelの表（コピー＆ペースト）やExcelファイル(.xlsx)、テキストファイルから、事業所、ケアマネージャー、利用者、サービスコードをまとめて一括追加できます。
                <br />
                <strong>表をそのままコピー（Ctrl+C）して「クリップボードから貼り付け」ボタンを押すか、下の枠内に Ctrl+V で直接貼り付けてください。</strong>
              </p>

              {/* Textarea for Import */}
              <div className="space-y-3">
                <textarea
                  value={aiInputText}
                  onChange={(e) => setAiInputText(e.target.value)}
                  onPaste={(e) => {
                    const text = e.clipboardData.getData("text/plain") || e.clipboardData.getData("text");
                    if (text) {
                      setAiErrorMsg("");
                    }
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    const file = e.dataTransfer.files?.[0];
                    if (file) {
                      handleExcelOrTextFile(file);
                    } else {
                      const text = e.dataTransfer.getData("text");
                      if (text) setAiInputText(text);
                    }
                  }}
                  onDragOver={(e) => e.preventDefault()}
                  className="w-full h-40 text-xs font-mono border border-slate-300 rounded-lg p-3 outline-hidden focus:ring-2 focus:ring-teal-500/20 bg-slate-50"
                  placeholder="【Excelからセルをコピーしてここに貼り付けてください】
例（TSV/CSV表形式）:
大津居宅介護支援事業所	土井 益実
桃の郷居宅介護支援事業所	長谷川 順一
山科ケアプランセンター	吉川 雅美

例（箇条書き・リスト）:
利用者：中島 富美子, 鈴木 喜代子, 田中 義雄
サービスコード：身体01, 身体1, 身体2, 身体1生活1, 生活2..."
                />

                {/* Real-time TSV Live Detection Badge */}
                {aiInputText.trim() && (
                  <div className="bg-teal-50/80 border border-teal-200 rounded-lg p-3 text-xs text-teal-900 space-y-1.5">
                    <div className="font-bold flex items-center space-x-1 text-teal-800">
                      <CheckCircle className="w-4 h-4 text-teal-600" />
                      <span>【リアルタイム解析検出】貼り付けテキストから以下を認識しました:</span>
                    </div>
                    <div className="flex flex-wrap gap-2 text-[11px]">
                      <span className="bg-white px-2.5 py-1 rounded border border-teal-200 shadow-2xs">
                        居宅事業所: <strong className="text-teal-700">{liveParsed.offices.length}</strong> 件
                      </span>
                      <span className="bg-white px-2.5 py-1 rounded border border-teal-200 shadow-2xs">
                        ケアマネ名: <strong className="text-teal-700">{liveParsed.offices.reduce((acc, o) => acc + o.careManagers.length, 0)}</strong> 名
                      </span>
                      <span className="bg-white px-2.5 py-1 rounded border border-teal-200 shadow-2xs">
                        利用者名: <strong className="text-teal-700">{liveParsed.users.length}</strong> 名
                      </span>
                      <span className="bg-white px-2.5 py-1 rounded border border-teal-200 shadow-2xs">
                        サービスコード: <strong className="text-teal-700">{liveParsed.serviceCodes.length}</strong> 種類
                      </span>
                    </div>
                  </div>
                )}
                
                {/* Control buttons */}
                <div className="flex flex-wrap gap-2 justify-end">
                  <button
                    onClick={handleAiInputClear}
                    className="px-3.5 py-2 border border-slate-300 text-slate-700 rounded-lg text-xs font-semibold hover:bg-slate-50 transition-colors"
                  >
                    クリア
                  </button>

                  {/* Button 1: Direct Local Instant Import */}
                  <button
                    onClick={handleDirectExcelImport}
                    className="px-4 py-2 bg-emerald-700 text-white rounded-lg text-xs font-bold hover:bg-emerald-800 transition-colors flex items-center space-x-1.5 shadow-xs"
                    title="AI通信を使わずに貼り付けた表から直接マスターに追加します"
                  >
                    <CheckCircle className="w-3.5 h-3.5" />
                    <span>① Excelデータから即時反映 (高速)</span>
                  </button>

                  {/* Button 2: AI Enhanced Import */}
                  <button
                    onClick={handleAiImportUpdate}
                    disabled={isAiLoading}
                    className="px-4 py-2 bg-teal-600 text-white rounded-lg text-xs font-bold hover:bg-teal-700 transition-colors disabled:opacity-50 flex items-center space-x-1.5 shadow-xs"
                    title="AIを使って高度に表記揺れを推測・補正してマスターを更新します"
                  >
                    {isAiLoading ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <Sparkles className="w-3.5 h-3.5" />
                    )}
                    <span>{isAiLoading ? "AI解析中..." : "② AI高度解析でマスター更新"}</span>
                  </button>
                </div>
              </div>

              {/* Status Banner */}
              {aiSuccessMsg && (
                <div className="p-4 bg-emerald-50 text-emerald-800 text-xs border border-emerald-100 rounded-lg flex items-start space-x-2">
                  <CheckCircle className="w-5 h-5 text-emerald-600 flex-shrink-0 mt-0.5" />
                  <pre className="font-sans whitespace-pre-wrap">{aiSuccessMsg}</pre>
                </div>
              )}
              {aiErrorMsg && (
                <div className="p-4 bg-rose-50 text-rose-800 text-xs border border-rose-100 rounded-lg flex items-start space-x-2">
                  <AlertCircle className="w-5 h-5 text-rose-600 flex-shrink-0 mt-0.5" />
                  <p>{aiErrorMsg}</p>
                </div>
              )}
            </div>

            {/* Manual Master Data Viewers & Tweakers */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              
              {/* Offices & Care Managers List */}
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-4">
                <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                  <h3 className="text-sm font-bold text-slate-900">
                    居宅介護支援事業所 ＆ ケアマネージャー 登録リスト
                  </h3>
                  <button
                    onClick={handleResetMasterData}
                    className="text-[11px] text-slate-500 hover:text-rose-600 bg-slate-100 hover:bg-rose-50 px-2.5 py-1 rounded border border-slate-200 transition-colors flex items-center space-x-1"
                    title="古いダミーデータや表記揺れをクリアし、マスターデータを標準初期状態に戻します"
                  >
                    <RefreshCw className="w-3 h-3 text-slate-400 hover:text-rose-500" />
                    <span>マスター初期状態に戻す</span>
                  </button>
                </div>
                
                {/* Add Office form */}
                <div className="flex space-x-2">
                  <input
                    type="text"
                    value={newOfficeName}
                    onChange={(e) => setNewOfficeName(e.target.value)}
                    className="flex-1 text-xs border border-slate-300 rounded-lg px-2.5 py-1.5 outline-hidden"
                    placeholder="新しい居宅事業所名を入力 (例: 大津居宅介護支援事業所)"
                  />
                  <button
                    onClick={handleAddOffice}
                    className="bg-slate-800 text-white text-xs px-3 py-1.5 rounded-lg hover:bg-slate-900 transition-colors font-bold"
                  >
                    追加
                  </button>
                </div>

                {/* Add Care Manager form */}
                <div className="bg-slate-50 p-3 rounded-lg border border-slate-100 space-y-2">
                  <p className="text-[10px] font-bold text-slate-500">事業所に所属マネージャーを追加</p>
                  <div className="grid grid-cols-2 gap-2">
                    <select
                      value={selectedOfficeForManager}
                      onChange={(e) => setSelectedOfficeForManager(e.target.value)}
                      className="text-xs border border-slate-300 rounded p-1 bg-white"
                    >
                      <option value="">-- 事業所選択 --</option>
                      {offices.map((o) => (
                        <option key={o.officeName} value={o.officeName}>
                          {o.officeName}
                        </option>
                      ))}
                    </select>
                    <input
                      type="text"
                      value={newCareManagerName}
                      onChange={(e) => setNewCareManagerName(e.target.value)}
                      className="text-xs border border-slate-300 rounded p-1 bg-white"
                      placeholder="ケアマネ名"
                    />
                  </div>
                  <div className="flex justify-end">
                    <button
                      onClick={handleAddCareManager}
                      className="bg-teal-600 text-white text-xs px-2.5 py-1 rounded font-semibold hover:bg-teal-700 transition-colors"
                    >
                      ケアマネ追加
                    </button>
                  </div>
                </div>

                {/* Listings with Inline Editing */}
                <div className="max-h-72 overflow-y-auto space-y-2 border border-slate-100 rounded-lg p-2 bg-slate-50/50">
                  {offices.map((o) => (
                    <div key={o.officeName} className="p-2.5 bg-white rounded border border-slate-200 flex flex-col justify-between space-y-2">
                      <div className="flex justify-between items-center border-b border-slate-100 pb-1">
                        {editingOfficeName === o.officeName ? (
                          <div className="flex items-center space-x-1 flex-1 mr-2">
                            <input
                              type="text"
                              value={editingOfficeNewVal}
                              onChange={(e) => setEditingOfficeNewVal(e.target.value)}
                              className="text-xs border border-teal-500 rounded px-1.5 py-0.5 flex-1 bg-teal-50/30"
                              autoFocus
                            />
                            <button
                              onClick={() => handleSaveEditOffice(o.officeName)}
                              className="text-teal-700 hover:text-teal-900 text-xs p-1"
                              title="保存"
                            >
                              <Check className="w-3.5 h-3.5" />
                            </button>
                            <button
                              onClick={() => setEditingOfficeName(null)}
                              className="text-slate-400 hover:text-slate-600 text-xs p-1"
                              title="キャンセル"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center space-x-2">
                            <span className="font-bold text-xs text-slate-900">{o.officeName}</span>
                            <button
                              onClick={() => handleStartEditOffice(o.officeName)}
                              className="text-slate-400 hover:text-teal-600 text-[10px]"
                              title="事業所名を名称変更・修正"
                            >
                              <Edit3 className="w-3 h-3" />
                            </button>
                          </div>
                        )}

                        <button
                          onClick={() => handleDeleteOffice(o.officeName)}
                          className="text-rose-500 hover:text-rose-700 text-[10px] font-bold"
                        >
                          削除
                        </button>
                      </div>
                      
                      <div className="flex flex-wrap gap-1.5">
                        {o.careManagers.map((m) => {
                          const isEditingThisCM = editingCMKey === `${o.officeName}::${m}`;
                          return (
                            <span
                              key={m}
                              className="bg-slate-100 text-slate-700 text-[10px] px-2 py-0.5 rounded-sm inline-flex items-center space-x-1 border border-slate-200"
                            >
                              {isEditingThisCM ? (
                                <div className="flex items-center space-x-1">
                                  <input
                                    type="text"
                                    value={editingCMNewVal}
                                    onChange={(e) => setEditingCMNewVal(e.target.value)}
                                    className="text-[10px] border border-teal-500 rounded px-1 w-20 bg-white"
                                    autoFocus
                                  />
                                  <button
                                    onClick={() => handleSaveEditCareManager(o.officeName, m)}
                                    className="text-teal-700 hover:text-teal-900"
                                  >
                                    <Check className="w-3 h-3" />
                                  </button>
                                  <button
                                    onClick={() => setEditingCMKey(null)}
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <X className="w-3 h-3" />
                                  </button>
                                </div>
                              ) : (
                                <>
                                  <span>{m}</span>
                                  <button
                                    onClick={() => handleStartEditCareManager(o.officeName, m)}
                                    className="text-slate-400 hover:text-teal-600 ml-0.5"
                                    title="ケアマネ名を修正"
                                  >
                                    <Edit3 className="w-2.5 h-2.5" />
                                  </button>
                                  <button
                                    onClick={() => handleDeleteCareManager(o.officeName, m)}
                                    className="text-slate-400 hover:text-rose-500 font-bold ml-0.5"
                                    title="削除"
                                  >
                                    ×
                                  </button>
                                </>
                              )}
                            </span>
                          );
                        })}
                        {o.careManagers.length === 0 && (
                          <span className="text-[10px] text-slate-400">（担当ケアマネ未登録）</span>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Users & Service Codes List */}
              <div className="space-y-6">
                
                {/* Users Registry */}
                <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-4">
                  <h3 className="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">
                    利用者 登録リスト
                  </h3>
                  
                  <div className="flex space-x-2">
                    <input
                      type="text"
                      value={newUserName}
                      onChange={(e) => setNewUserName(e.target.value)}
                      className="flex-1 text-xs border border-slate-300 rounded-lg px-2.5 py-1.5 outline-hidden"
                      placeholder="新しい利用者名を入力 (敬称不要)"
                    />
                    <button
                      onClick={handleAddUserManual}
                      className="bg-slate-800 text-white text-xs px-3 py-1.5 rounded-lg hover:bg-slate-900 transition-colors font-bold"
                    >
                      追加
                    </button>
                  </div>

                  <div className="max-h-36 overflow-y-auto flex flex-wrap gap-1.5 p-2 bg-slate-50 rounded-lg border border-slate-100">
                    {users.map((u) => (
                      <span
                        key={u.name}
                        className="bg-white border border-slate-200 text-slate-800 text-xs px-2.5 py-1 rounded inline-flex items-center space-x-1"
                      >
                        <span>{u.name}</span>
                        <button
                          onClick={() => handleDeleteUser(u.name)}
                          className="text-slate-400 hover:text-rose-600 font-bold ml-1.5"
                        >
                          ×
                        </button>
                      </span>
                    ))}
                  </div>
                </div>

                {/* Service Codes Registry */}
                <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs space-y-4">
                  <h3 className="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">
                    介護サービスコード 登録リスト
                  </h3>
                  
                  <div className="flex space-x-2">
                    <input
                      type="text"
                      value={newServiceCode}
                      onChange={(e) => setNewServiceCode(e.target.value)}
                      className="flex-1 text-xs border border-slate-300 rounded-lg px-2.5 py-1.5 outline-hidden"
                      placeholder="例: 身体1生活3, 身体01..."
                    />
                    <button
                      onClick={handleAddServiceCodeManual}
                      className="bg-slate-800 text-white text-xs px-3 py-1.5 rounded-lg hover:bg-slate-900 transition-colors font-bold"
                    >
                      追加
                    </button>
                  </div>

                  <div className="max-h-36 overflow-y-auto flex flex-wrap gap-1.5 p-2 bg-slate-50 rounded-lg border border-slate-100">
                    {serviceCodes.map((code) => (
                      <span
                        key={code}
                        className="bg-white border border-slate-200 text-slate-800 text-xs px-2.5 py-1 rounded inline-flex items-center space-x-1"
                      >
                        <span>{code}</span>
                        <button
                          onClick={() => handleDeleteServiceCode(code)}
                          className="text-slate-400 hover:text-rose-600 font-bold ml-1.5"
                        >
                          ×
                        </button>
                      </span>
                    ))}
                  </div>
                </div>

              </div>

            </div>

          </div>
        )}

      </main>
    </div>
  );
}
