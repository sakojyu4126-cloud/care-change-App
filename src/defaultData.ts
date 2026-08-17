import { CareManagerOffice, CareUser, SenderInfo } from "./types";

export const DEFAULT_OFFICES: CareManagerOffice[] = [
  {
    officeName: "大津居宅介護支援事業所",
    careManagers: ["土井 益実", "西村 敏行"]
  },
  {
    officeName: "桃の郷居宅介護支援事業所",
    careManagers: ["長谷川 順一", "山田 和子"]
  },
  {
    officeName: "山科ケアプランセンター",
    careManagers: ["吉川 雅美", "中村 健司"]
  }
];

export const DEFAULT_USERS: CareUser[] = [
  { name: "中島 富美子" },
  { name: "佐藤 健治" },
  { name: "鈴木 喜代子" },
  { name: "田中 義雄" }
];

export const DEFAULT_SERVICE_CODES = [
  "身体01",
  "身体1",
  "身体2",
  "身体1生活1",
  "生活2",
  "生活3",
  "身体1生活2",
  "身体2生活1",
  "身体3"
];

export const DEFAULT_SENDER_INFO: SenderInfo = {
  postalCode: "607-8022",
  address: "京都市山科区四ノ宮小金塚1-134",
  tel: "(075) 574-7171",
  fax: "(075) 574-7979",
  stationName: "ヘルパーステーション桃の郷京都東山",
  picName: "長島"
};
