export interface CareManagerOffice {
  officeName: string;
  careManagers: string[];
}

export interface CareUser {
  name: string;
}

export interface SenderInfo {
  postalCode: string;
  address: string;
  tel: string;
  fax: string;
  stationName: string;
  picName: string;
}

export interface ServiceTimeContent {
  startTime: string;
  endTime: string;
  serviceCode: string;
  customService?: string;
  remarks?: string; // only used or displayed for "after"
}

export interface ServiceChangeItem {
  id: string;
  dayOfWeek: string;
  dayOfWeekAfter?: string;
  before: ServiceTimeContent;
  after: ServiceTimeContent;
}

export interface NoticeRecord {
  id: string;
  createdAt: string; // ISO string
  reportDate: string; // e.g. "R8年 5月 28日" or "2026年5月28日"
  startDate: string; // e.g. "6月 1日"
  officeName: string;
  careManagerName: string;
  userName: string;
  senderInfo: SenderInfo;
  preGreeting: string;
  postComment: string;
  changes: ServiceChangeItem[];
}
