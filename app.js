const STORAGE_KEY = "daol-fi-strategy-dashboard-v2";
const OLD_STORAGE_KEY = "daol-fi-strategy-dashboard-v1";
const OPERATING_WEEK_CUTOFF = "2026-W24";
const FIRST_MEETING_WEEK = "2026-W24";

// Microsoft Entra/SharePoint connection settings for the first shared-storage test.
// Fill clientId, tenantId, and sharePointHost before using Login/Test.
// Use "auto" redirectUri for localhost and deployed HTTPS URLs. Register every
// actual URL used here as a SPA redirect URI in Microsoft Entra.
const MICROSOFT_CONFIG = {
  clientId: "543edd8e-355d-43fc-8f9b-b5b89bb19dff",
  tenantId: "bbeeca21-dbbf-4217-9ad3-90bf14255c7a",
  redirectUri: "auto",
  sharePointHost: "daolfund.sharepoint.com",
  sitePath: "/sites/msteams_4b367d",
  testListName: "FI_WeeklyOpinions"
};

const GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All"];
const SUMMARY_MEMBER_ID = "__weeklySummary";
const SUMMARY_MEMBER = { id: SUMMARY_MEMBER_ID, name: "종합의견", team: "Weekly consensus", active: true, sortOrder: 9999 };

const DEFAULT_SETTINGS = {
  pnlUnit: "bp",
  evaluationLagWeeks: 1
};

const ARCHIVE_CALCULATION_VERSION = 1;
const WEEKLY_SUMMARY_LIST_NAME = "FI_WeeklySummaries";
const WEEKLY_SUMMARY_SCHEMA = [
  ["appId", "text"],
  ["week", "text"],
  ["rateDuration", "number"],
  ["curveDuration", "number"],
  ["creditDuration", "number"],
  ["ktb3yRangeLow", "number"],
  ["ktb3yRangeHigh", "number"],
  ["curveSpreadRangeLow", "number"],
  ["curveSpreadRangeHigh", "number"],
  ["creditSpreadRangeLow", "number"],
  ["creditSpreadRangeHigh", "number"],
  ["confidence", "number"],
  ["rateConfidence", "number"],
  ["curveConfidence", "number"],
  ["creditConfidence", "number"],
  ["overallSummaryText", "multiline"],
  ["rateSummaryText", "multiline"],
  ["curveSummaryText", "multiline"],
  ["creditSummaryText", "multiline"],
  ["status", "text"],
  ["draftSource", "text"],
  ["completedAt", "text"],
  ["createdAt", "text"],
  ["updatedAt", "text"]
];
const ARCHIVE_LIST_NAMES = {
  archives: "FI_WeeklyArchives",
  members: "FI_WeeklyArchiveMembers",
  opinionStrategies: "FI_WeeklyArchiveOpinionStrategies",
  markets: "FI_WeeklyArchiveMarkets",
  results: "FI_WeeklyArchiveResults",
  rawJson: "FI_WeeklyArchiveRawJson"
};

const ARCHIVE_LIST_SCHEMAS = {
  [ARCHIVE_LIST_NAMES.archives]: [
    ["appId", "text"],
    ["archiveId", "text"],
    ["week", "text"],
    ["version", "number"],
    ["status", "text"],
    ["archivedAt", "text"],
    ["archivedBy", "text"],
    ["memo", "multiline"],
    ["calculationVersion", "number"],
    ["sourceHash", "text"]
  ],
  [ARCHIVE_LIST_NAMES.members]: [
    ["appId", "text"],
    ["archiveId", "text"],
    ["week", "text"],
    ["version", "number"],
    ["memberId", "text"],
    ["memberName", "text"],
    ["team", "text"],
    ["active", "boolean"],
    ["sortOrder", "number"]
  ],
  [ARCHIVE_LIST_NAMES.opinionStrategies]: [
    ["appId", "text"],
    ["archiveId", "text"],
    ["week", "text"],
    ["version", "number"],
    ["sourceOpinionId", "text"],
    ["memberId", "text"],
    ["memberName", "text"],
    ["team", "text"],
    ["strategyKey", "text"],
    ["duration", "number"],
    ["rangeLow", "number"],
    ["rangeHigh", "number"],
    ["confidence", "number"],
    ["rationaleText", "multiline"],
    ["rationaleTags", "multiline"],
    ["sourceCreatedAt", "text"],
    ["sourceUpdatedAt", "text"]
  ],
  [ARCHIVE_LIST_NAMES.markets]: [
    ["appId", "text"],
    ["archiveId", "text"],
    ["week", "text"],
    ["version", "number"],
    ["marketRole", "text"],
    ["marketWeek", "text"],
    ["asOf", "text"],
    ["baseDate", "text"],
    ["ktb3y", "number"],
    ["ktb10y", "number"],
    ["curveSpread", "number"],
    ["msb2y", "number"],
    ["creditAAm2y", "number"],
    ["creditSpread", "number"],
    ["memo", "multiline"]
  ],
  [ARCHIVE_LIST_NAMES.results]: [
    ["appId", "text"],
    ["archiveId", "text"],
    ["week", "text"],
    ["version", "number"],
    ["sourceOpinionId", "text"],
    ["memberId", "text"],
    ["memberName", "text"],
    ["strategyKey", "text"],
    ["evaluatedWeek", "text"],
    ["marketDelta", "number"],
    ["pnl", "number"],
    ["ytdPnl", "number"],
    ["rangeHit", "boolean"]
  ],
  [ARCHIVE_LIST_NAMES.rawJson]: [
    ["appId", "text"],
    ["archiveId", "text"],
    ["week", "text"],
    ["version", "number"],
    ["json", "multiline"]
  ]
};

const STRATEGIES = [
  {
    key: "rate",
    title: "금리 방향성",
    durationKey: "rateDuration",
    marketKey: "ktb3y",
    rangeLowKey: "ktb3yRangeLow",
    rangeHighKey: "ktb3yRangeHigh",
    multiplier: 2.85,
    unit: "%",
    rangeUnit: "%"
  },
  {
    key: "curve",
    title: "국고 3Y-10Y 커브",
    durationKey: "curveDuration",
    marketKey: "curveSpread",
    rangeLowKey: "curveSpreadRangeLow",
    rangeHighKey: "curveSpreadRangeHigh",
    multiplier: 8.0,
    unit: "bp",
    rangeUnit: "bp"
  },
  {
    key: "credit",
    title: "여전 AA- 2Y - 통안 2Y",
    durationKey: "creditDuration",
    marketKey: "creditSpread",
    rangeLowKey: "creditSpreadRangeLow",
    rangeHighKey: "creditSpreadRangeHigh",
    multiplier: 1.85,
    unit: "bp",
    rangeUnit: "bp"
  }
];

const sampleData = {
  schemaVersion: 2,
  members: [
    { id: "m-001", name: "구성원 A", team: "채권운용본부", active: true, sortOrder: 1 },
    { id: "m-002", name: "구성원 B", team: "채권운용본부", active: true, sortOrder: 2 },
    { id: "m-003", name: "구성원 C", team: "리서치", active: true, sortOrder: 3 },
    { id: "m-004", name: "구성원 D", team: "채권운용본부", active: true, sortOrder: 4 }
  ],
  weeklyOpinions: [],
  marketData: [],
  dailyRates: [],
  marketWeekMappings: [],
  marketMappingHistory: [],
  uploadedRateFiles: [],
  weeklySummaries: [],
  settings: { ...DEFAULT_SETTINGS },
  performanceSnapshots: [],
  weeklyArchives: [],
  weeklyArchiveMembers: [],
  weeklyArchiveOpinionStrategies: [],
  weeklyArchiveMarkets: [],
  weeklyArchiveResults: [],
  weeklyArchiveRawJson: []
};
let state = loadState();
let selectedWeek = latestWeek();
let selectedPerformanceMetric = "rate";
let selectedTopRankingMetric = "rate";
let selectedArchiveId = "";
let selectedOpinionWeek = nextOpinionWeek();
let selectedWeeklySummaryWeek = selectedWeek;
const strategyChartModes = { rate: "duration", curve: "duration", credit: "duration" };

const $ = selector => document.querySelector(selector);

const els = {
  navButtons: document.querySelectorAll(".nav-button"),
  views: document.querySelectorAll(".view"),
  pageTitle: $("#pageTitle"),
  weekSelect: $("#weekSelect"),
  currentWeekButton: $("#currentWeekButton"),
  authToolbar: $(".auth-toolbar"),
  msSignInButton: $("#msSignInButton"),
  graphTestButton: $("#graphTestButton"),
  sharePointLoadButton: $("#sharePointLoadButton"),
  msAuthStatus: $("#msAuthStatus"),
  msAccountBadge: $("#msAccountBadge"),
  lastUpdatedBadge: $("#lastUpdatedBadge"),
  exportJsonButton: $("#exportJsonButton"),
  importJsonInput: $("#importJsonInput"),
  exportCsvButton: $("#exportCsvButton"),
  kpiParticipants: $("#kpiParticipants"),
  kpiActiveMembers: $("#kpiActiveMembers"),
  kpiRateDuration: $("#kpiRateDuration"),
  kpiCurveDuration: $("#kpiCurveDuration"),
  kpiCreditDuration: $("#kpiCreditDuration"),
  topRanking: $("#topRanking"),
  topRankingMetricSelect: $("#topRankingMetricSelect"),
  fullRankingTable: $("#fullRankingTable"),
  weekConsensus: $("#weekConsensus"),
  marketMovementList: $("#marketMovementList"),
  opinionForm: $("#opinionForm"),
  opinionWeek: $("#opinionWeek"),
  opinionMember: $("#opinionMember"),
  loadOpinionButton: $("#loadOpinionButton"),
  opinionLookupStatus: $("#opinionLookupStatus"),
  clearOpinionButton: $("#clearOpinionButton"),
  weeklySummaryForm: $("#weeklySummaryForm"),
  weeklySummaryWeek: $("#weeklySummaryWeek"),
  weeklySummaryStatus: $("#weeklySummaryStatus"),
  weeklySummaryDraftButton: $("#weeklySummaryDraftButton"),
  clearWeeklySummaryButton: $("#clearWeeklySummaryButton"),
  durationCharts: $("#durationCharts"),
  weeklyPerformanceTable: $("#weeklyPerformanceTable"),
  weeklyPerformanceChart: $("#weeklyPerformanceChart"),
  performanceMemberTable: $("#performanceMemberTable"),
  performanceDetailTable: $("#performanceDetailTable"),
  performanceMetricSelect: $("#performanceMetricSelect"),
  performanceChart: $("#performanceChart"),
  archiveSummaryList: $("#archiveSummaryList"),
  archiveSelectedDetail: $("#archiveSelectedDetail"),
  archiveWeekFilter: $("#archiveWeekFilter"),
  archiveVersionFilter: $("#archiveVersionFilter"),
  archiveMarketSnapshot: $("#archiveMarketSnapshot"),
  archiveOpinionTable: $("#archiveOpinionTable"),
  archiveResultTable: $("#archiveResultTable"),
  archiveMemo: $("#archiveMemo"),
  archiveWeekButton: $("#archiveWeekButton"),
  archiveStatus: $("#archiveStatus"),
  marketForm: $("#marketForm"),
  marketWeek: $("#marketWeek"),
  marketBaseDate: $("#marketBaseDate"),
  marketDatePreview: $("#marketDatePreview"),
  rateDateOptions: $("#rateDateOptions"),
  rateFileInput: $("#rateFileInput"),
  rateUploadStatus: $("#rateUploadStatus"),
  rateCoverage: $("#rateCoverage"),
  clearMarketButton: $("#clearMarketButton"),
  latestMarketList: $("#latestMarketList"),
  memberForm: $("#memberForm"),
  memberTable: $("#memberTable"),
  clearMemberButton: $("#clearMemberButton")
};

let msalClient = null;
let msAccount = null;

function relocateMarketDataIntoSettings() {
  const marketView = document.querySelector("#marketView");
  const settingsView = document.querySelector("#settingsView");
  if (!marketView || !settingsView || settingsView.querySelector("#marketSettingsPanel")) return;
  const panel = document.createElement("div");
  panel.id = "marketSettingsPanel";
  panel.className = "settings-market-panel";
  while (marketView.firstChild) panel.appendChild(marketView.firstChild);
  settingsView.appendChild(panel);
  marketView.remove();
}

relocateMarketDataIntoSettings();

function makeMarket(week, asOf, ktb3y, ktb10y, msb2y, creditAA2y, memo = "") {
  return {
    week,
    asOf,
    ktb3y,
    ktb10y,
    msb2y,
    creditAA2y,
    curveSpread: round((ktb10y - ktb3y) * 100, 1),
    creditSpread: round((creditAA2y - msb2y) * 100, 1),
    memo
  };
}

function loadState() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      return normalizeState(JSON.parse(stored));
    } catch {
      return structuredClone(sampleData);
    }
  }
  localStorage.removeItem(OLD_STORAGE_KEY);
  const initialState = normalizeState(sampleData);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(initialState));
  return initialState;
}

function normalizeState(data) {
  return {
    schemaVersion: 2,
    members: Array.isArray(data.members) ? data.members : [],
    weeklyOpinions: filterOperatingWeeks(Array.isArray(data.weeklyOpinions) ? data.weeklyOpinions : []),
    marketData: filterOperatingWeeks(Array.isArray(data.marketData) ? data.marketData.map(normalizeMarket) : []),
    dailyRates: Array.isArray(data.dailyRates) ? data.dailyRates.map(normalizeDailyRate).filter(Boolean) : [],
    marketWeekMappings: filterOperatingWeeks(Array.isArray(data.marketWeekMappings)
      ? data.marketWeekMappings.map(normalizeWeekMapping).filter(Boolean)
      : legacyMarketMappings(data.marketData)),
    marketMappingHistory: filterOperatingWeeks(Array.isArray(data.marketMappingHistory) ? data.marketMappingHistory : []),
    uploadedRateFiles: Array.isArray(data.uploadedRateFiles) ? data.uploadedRateFiles : [],
    weeklySummaries: filterOperatingWeeks(Array.isArray(data.weeklySummaries) ? data.weeklySummaries : []),
    settings: normalizeSettings(data.settings),
    performanceSnapshots: Array.isArray(data.performanceSnapshots) ? data.performanceSnapshots : [],
    weeklyArchives: filterOperatingWeeks(Array.isArray(data.weeklyArchives) ? data.weeklyArchives : []),
    weeklyArchiveMembers: filterOperatingWeeks(Array.isArray(data.weeklyArchiveMembers) ? data.weeklyArchiveMembers : []),
    weeklyArchiveOpinionStrategies: filterOperatingWeeks(Array.isArray(data.weeklyArchiveOpinionStrategies) ? data.weeklyArchiveOpinionStrategies : []),
    weeklyArchiveMarkets: filterOperatingWeeks(Array.isArray(data.weeklyArchiveMarkets) ? data.weeklyArchiveMarkets : []),
    weeklyArchiveResults: filterOperatingWeeks(Array.isArray(data.weeklyArchiveResults) ? data.weeklyArchiveResults : []),
    weeklyArchiveRawJson: filterOperatingWeeks(Array.isArray(data.weeklyArchiveRawJson) ? data.weeklyArchiveRawJson : [])
  };
}

function filterOperatingWeeks(rows) {
  return rows.filter(row => !row?.week || String(row.week) >= OPERATING_WEEK_CUTOFF);
}

function normalizeSettings(settings = {}) {
  const evaluationLagWeeks = Number(settings.evaluationLagWeeks);
  return {
    pnlUnit: String(settings.pnlUnit || DEFAULT_SETTINGS.pnlUnit),
    evaluationLagWeeks: Number.isFinite(evaluationLagWeeks) && evaluationLagWeeks > 0
      ? evaluationLagWeeks
      : DEFAULT_SETTINGS.evaluationLagWeeks
  };
}

function normalizeMarket(item) {
  const ktb3y = Number(item.ktb3y);
  const ktb10y = Number(item.ktb10y);
  const msb2y = Number(item.msb2y);
  const creditAA2y = Number(item.creditAA2y);
  return {
    ...item,
    ktb3y,
    ktb10y,
    msb2y,
    creditAA2y,
    curveSpread: Number.isFinite(Number(item.curveSpread)) ? Number(item.curveSpread) : round((ktb10y - ktb3y) * 100, 1),
    creditSpread: Number.isFinite(Number(item.creditSpread)) ? Number(item.creditSpread) : round((creditAA2y - msb2y) * 100, 1)
  };
}

function normalizeDailyRate(item) {
  const date = normalizeDateText(item.date);
  const ktb3y = Number(item.ktb3y);
  const ktb10y = Number(item.ktb10y);
  const msb2y = Number(item.msb2y);
  const creditAA2y = Number(item.creditAA2y);
  const curveSpread = Number(item.curveSpread);
  const creditSpread = Number(item.creditSpread);
  if (!date || [ktb3y, ktb10y, msb2y, creditAA2y].some(value => !Number.isFinite(value))) return null;
  return {
    date,
    ktb3y,
    ktb10y,
    msb2y,
    creditAA2y,
    curveSpread: Number.isFinite(curveSpread) ? curveSpread : round((ktb10y - ktb3y) * 100, 1),
    creditSpread: Number.isFinite(creditSpread) ? creditSpread : round((creditAA2y - msb2y) * 100, 1)
  };
}

function normalizeWeekMapping(item) {
  const week = item?.week;
  const baseDate = normalizeDateText(item?.baseDate || item?.asOf);
  if (!week || !baseDate) return null;
  return {
    week,
    baseDate,
    memo: item.memo || "",
    createdAt: item.createdAt || item.updatedAt || new Date().toISOString(),
    updatedAt: item.updatedAt || item.createdAt || new Date().toISOString()
  };
}

function legacyMarketMappings(marketData = []) {
  return Array.isArray(marketData)
    ? marketData.map(row => normalizeWeekMapping({ week: row.week, baseDate: row.asOf, memo: row.memo })).filter(Boolean)
    : [];
}

function saveState() {
  state = normalizeState(state);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function microsoftConfigReady() {
  return Boolean(
    MICROSOFT_CONFIG.clientId &&
    MICROSOFT_CONFIG.tenantId &&
    MICROSOFT_CONFIG.sharePointHost &&
    !MICROSOFT_CONFIG.sharePointHost.includes("*")
  );
}

function currentPageRedirectUri() {
  if (typeof window === "undefined" || !window.location) return "";
  const url = new URL(window.location.href);
  url.search = "";
  url.hash = "";
  return url.href;
}

function microsoftRedirectUri() {
  const configured = String(MICROSOFT_CONFIG.redirectUri || "").trim();
  return configured && configured !== "auto"
    ? configured
    : currentPageRedirectUri();
}

function redirectOrigin() {
  try {
    return new URL(microsoftRedirectUri()).origin;
  } catch {
    return "";
  }
}

function redirectOriginMatches() {
  const expected = redirectOrigin();
  return !expected || (typeof window !== "undefined" && window.location.origin === expected);
}

function shortErrorMessage(error) {
  return String(error?.message || error || "unknown-error").slice(0, 140);
}

function setMicrosoftStatus(message, mode = "") {
  if (!els.msAuthStatus) return;
  els.msAuthStatus.textContent = message;
  els.authToolbar?.classList.toggle("connected", mode === "connected");
  els.authToolbar?.classList.toggle("error", mode === "error");
}

function setMicrosoftAccountBadge(account) {
  const label = account?.name || account?.username || "";
  if (els.msAccountBadge) {
    els.msAccountBadge.textContent = label ? `로그인: ${label}` : "로그인 안 됨";
    els.msAccountBadge.classList.toggle("signed-in", Boolean(label));
  }
  if (els.msSignInButton) {
    els.msSignInButton.hidden = Boolean(label);
  }
}

function formatKstDateTime(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "-";
  const parts = new Intl.DateTimeFormat("ko-KR", {
    timeZone: "Asia/Seoul",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", hour12: false
  }).formatToParts(d).reduce((acc, p) => (acc[p.type] = p.value, acc), {});
  return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}`;
}

function latestUpdatedAtIso() {
  const pools = [
    state.weeklyOpinions,
    state.weeklySummaries,
    state.marketWeekMappings,
    state.marketData
  ];
  let latest = "";
  pools.forEach(list => {
    (Array.isArray(list) ? list : []).forEach(row => {
      const u = row && row.updatedAt ? String(row.updatedAt) : "";
      if (u && u > latest) latest = u;
    });
  });
  return latest;
}

function renderLastUpdatedBadge() {
  if (!els.lastUpdatedBadge) return;
  const iso = latestUpdatedAtIso();
  els.lastUpdatedBadge.textContent = `최종 업데이트: ${iso ? formatKstDateTime(iso) : "-"}`;
}

function getMsalClient() {
  if (!microsoftConfigReady()) {
    throw new Error("microsoft-config-missing");
  }
  if (!redirectOriginMatches()) {
    throw new Error(`redirect-origin-mismatch:${microsoftRedirectUri()}`);
  }
  if (!window.msal?.PublicClientApplication) {
    throw new Error("msal-not-loaded");
  }
  if (!msalClient) {
    msalClient = new window.msal.PublicClientApplication({
      auth: {
        clientId: MICROSOFT_CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${MICROSOFT_CONFIG.tenantId}`,
        redirectUri: microsoftRedirectUri()
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
      }
    });
  }
  return msalClient;
}

async function signInMicrosoft() {
  try {
    const client = getMsalClient();
    const accounts = client.getAllAccounts();
    if (accounts.length) {
      msAccount = accounts[0];
    } else {
      const result = await client.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
      msAccount = result.account;
    }
    setMicrosoftStatus(`${msAccount?.name || msAccount?.username || "Microsoft"} 로그인됨`, "connected");
    setMicrosoftAccountBadge(msAccount);
    return msAccount;
  } catch (error) {
    console.error(error);
    if (error.message === "microsoft-config-missing") {
      setMicrosoftStatus("clientId, tenantId, SharePoint host 설정 필요", "error");
      alert("Microsoft 앱 설정값을 먼저 app.js 상단 MICROSOFT_CONFIG에 입력해야 합니다.");
      return null;
    }
    if (error.message.startsWith("redirect-origin-mismatch")) {
      setMicrosoftStatus("localhost 주소로 다시 열어야 합니다.", "error");
      alert(`현재 주소와 Microsoft 리디렉션 주소가 다릅니다.\n아래 주소로 다시 열어 주세요.\n\n${microsoftRedirectUri()}`);
      return null;
    }
    if (error.message === "msal-not-loaded") {
      setMicrosoftStatus("MSAL 라이브러리를 불러오지 못했습니다.", "error");
      alert("MSAL 라이브러리를 불러오지 못했습니다. 네트워크 연결 또는 회사 보안 정책을 확인해 주세요.");
      return null;
    }
    setMicrosoftStatus("Microsoft 로그인 실패", "error");
    return null;
  }
}

async function graphFetch(path, options = {}) {
  const client = getMsalClient();
  const account = msAccount || client.getAllAccounts()[0] || await signInMicrosoft();
  if (!account) throw new Error("microsoft-account-missing");
  const token = await client.acquireTokenSilent({ account, scopes: GRAPH_SCOPES })
    .catch(() => client.acquireTokenPopup({ account, scopes: GRAPH_SCOPES }));
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`graph-${response.status}: ${detail}`);
  }
  if (response.status === 204) return {};
  const text = await response.text();
  return text ? JSON.parse(text) : {};
}

function sharePointSitePath() {
  return MICROSOFT_CONFIG.sitePath.startsWith("/")
    ? MICROSOFT_CONFIG.sitePath
    : `/${MICROSOFT_CONFIG.sitePath}`;
}

async function getSharePointSite() {
  return graphFetch(`/sites/${MICROSOFT_CONFIG.sharePointHost}:${sharePointSitePath()}?$select=id,displayName,webUrl`);
}

async function getSharePointListsByName(siteId) {
  const lists = await graphFetch(`/sites/${encodeURIComponent(siteId)}/lists?$select=id,displayName,webUrl`);
  return new Map((lists.value || []).map(list => [list.displayName, list]));
}

async function readSharePointListItems(siteId, list, top = 500) {
  if (!list) throw new Error("sharepoint-list-missing");
  const items = [];
  let path = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(list.id)}/items?$expand=fields&$top=${top}`;
  while (path) {
    const data = await graphFetch(path);
    items.push(...(data.value || []));
    path = graphPathFromNextLink(data["@odata.nextLink"]);
  }
  return items;
}

async function getSharePointColumns(siteId, list) {
  if (!list) throw new Error("sharepoint-list-missing");
  const data = await graphFetch(`/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(list.id)}/columns?$select=name,displayName`);
  return data.value || [];
}

function graphPathFromNextLink(nextLink) {
  if (!nextLink) return "";
  const marker = "/v1.0";
  const index = nextLink.indexOf(marker);
  return index >= 0 ? nextLink.slice(index + marker.length) : nextLink;
}

function columnLookupKey(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function resolveSharePointFieldName(columns, fieldName) {
  if (fieldName === "Title") return "Title";
  const wanted = columnLookupKey(fieldName);
  const column = columns.find(item =>
    columnLookupKey(item.name) === wanted ||
    columnLookupKey(item.displayName) === wanted
  );
  return column?.name || fieldName;
}

function resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, fieldName) {
  if (fieldName === "Title") return "Title";
  const wanted = columnLookupKey(fieldName);
  const sampleKey = sampleFieldKeys.find(key => columnLookupKey(key) === wanted);
  if (sampleKey) return sampleKey;
  const column = columns.find(item =>
    columnLookupKey(item.name) === wanted ||
    columnLookupKey(item.displayName) === wanted
  );
  return column?.name || fieldName;
}

function resolveSharePointFields(columns, fields, sampleFieldKeys = []) {
  return Object.fromEntries(
    Object.entries(fields).map(([key, value]) => [resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, key), value])
  );
}

function sharePointFieldExists(columns, sampleFieldKeys, fieldName) {
  if (fieldName === "Title") return true;
  const wanted = columnLookupKey(fieldName);
  return (
    sampleFieldKeys.some(key => columnLookupKey(key) === wanted) ||
    columns.some(column =>
      columnLookupKey(column.name) === wanted ||
      columnLookupKey(column.displayName) === wanted
    )
  );
}

function resolveExistingSharePointFields(columns, fields, sampleFieldKeys = []) {
  return Object.fromEntries(
    Object.entries(fields)
      .filter(([key]) => sharePointFieldExists(columns, sampleFieldKeys, key))
      .map(([key, value]) => [resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, key), value])
  );
}

function applySharePointFieldAliases(fields, aliases) {
  const next = { ...fields };
  Object.entries(aliases).forEach(([appField, sharePointField]) => {
    if (appField in next && appField !== sharePointField) {
      next[sharePointField] = next[appField];
      delete next[appField];
    }
  });
  return next;
}

function omitSharePointFields(fields, fieldNames) {
  const next = { ...fields };
  fieldNames.forEach(fieldName => {
    delete next[fieldName];
  });
  return next;
}

function sharePointColumnDiagnostics(columns, sampleFieldKeys = []) {
  const columnText = columns
    .map(column => `${column.displayName || "(no display)"} -> ${column.name}`)
    .join("; ");
  const sampleText = sampleFieldKeys.join(", ");
  return `columns: ${columnText || "(none)"} | sample fields: ${sampleText || "(none)"}`;
}

function memberIdFromSortOrder(sortOrder) {
  return `FI${String(Number(sortOrder) || 0).padStart(2, "0")}`;
}

function fieldValue(fields, key, fallback = "") {
  return fields?.[key] ?? fallback;
}

function fieldValueAny(fields, keys, fallback = "") {
  const keyList = Array.isArray(keys) ? keys : [keys];
  const found = keyList.find(key => fields?.[key] !== undefined);
  return found ? fields[found] : fallback;
}

function numberField(fields, key) {
  const value = Number(fieldValue(fields, key));
  return Number.isFinite(value) ? value : 0;
}

function numberFieldAny(fields, keys) {
  const value = Number(fieldValueAny(fields, keys));
  return Number.isFinite(value) ? value : 0;
}

function nullableNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function sharePointNumber(value) {
  const number = nullableNumber(value);
  return number === null ? null : number;
}

function booleanField(fields, key, defaultValue = true) {
  const value = fieldValue(fields, key, defaultValue);
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const text = String(value).trim().toLowerCase();
  if (["false", "0", "no", "inactive"].includes(text)) return false;
  if (["true", "1", "yes", "active"].includes(text)) return true;
  return defaultValue;
}

function textListField(fields, key) {
  const value = fieldValue(fields, key, "");
  if (Array.isArray(value)) return value;
  return String(value || "")
    .split(/[|,]/)
    .map(item => item.trim())
    .filter(Boolean);
}

function mapSharePointMember(item) {
  const fields = item.fields || {};
  const appId = String(fieldValue(fields, "appId", "") || "").trim();
  const title = String(fieldValue(fields, "Title", "") || "").trim();
  return {
    id: appId || `sp-member-${item.id}`,
    name: String(fieldValue(fields, "name", title) || "").trim(),
    team: String(fieldValue(fields, "team", "") || "").trim(),
    email: String(fieldValue(fields, "email", "") || "").trim().toLowerCase(),
    active: booleanField(fields, "active", true),
    sortOrder: numberField(fields, "sortOrder")
  };
}

function sharePointMemberFields(member) {
  return {
    Title: member.name || member.id,
    appId: member.id,
    team: member.team || "",
    active: member.active !== false,
    sortOrder: Number(member.sortOrder) || 0
  };
}

async function saveSharePointMember(member) {
  const site = await getSharePointSite();
  const lists = await getSharePointListsByName(site.id);
  const list = lists.get("DAOL_FI_Members");
  if (!list) throw new Error("list-not-found:DAOL_FI_Members");
  const columns = await getSharePointColumns(site.id, list);
  const items = await readSharePointListItems(site.id, list);
  const sampleFieldKeys = [...new Set(items.flatMap(item => Object.keys(item.fields || {})))];
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const existing = items.find(item => {
    const fields = item.fields || {};
    return fieldValue(fields, appIdFieldName) === member.id || fieldValue(fields, "Title") === member.name;
  });
  const fields = resolveSharePointFields(columns, sharePointMemberFields(member), sampleFieldKeys);
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

async function findSharePointMemberItem(member) {
  const site = await getSharePointSite();
  const lists = await getSharePointListsByName(site.id);
  const list = lists.get("DAOL_FI_Members");
  if (!list) throw new Error("list-not-found:DAOL_FI_Members");
  const columns = await getSharePointColumns(site.id, list);
  const items = await readSharePointListItems(site.id, list);
  const sampleFieldKeys = [...new Set(items.flatMap(item => Object.keys(item.fields || {})))];
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const item = items.find(candidate => {
    const fields = candidate.fields || {};
    return fieldValue(fields, appIdFieldName) === member.id || fieldValue(fields, "Title") === member.name;
  });
  return { site, list, item };
}

async function deleteSharePointMember(member) {
  const { site, list, item } = await findSharePointMemberItem(member);
  if (!item) return { mode: "missing" };
  await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(item.id)}`,
    { method: "DELETE" }
  );
  return { mode: "deleted", itemId: item.id };
}

function mapSharePointOpinion(item) {
  const fields = item.fields || {};
  const appId = String(fieldValue(fields, "appId", "") || "").trim();
  const title = String(fieldValue(fields, "Title", "") || "").trim();
  const [titleWeek = "", titleMemberId = ""] = title.split("_");
  const confidenceValues = ["rateConfidence", "curveConfidence", "creditConfidence"].map(key => nullableNumber(fieldValue(fields, key, "")));
  return {
    id: appId || title || `sp-opinion-${item.id}`,
    week: String(fieldValue(fields, "week", titleWeek) || "").trim(),
    memberId: (String(fieldValueAny(fields, ["memberId", "memberid", "mmberId", "mmberid"], "") || "").trim()) || titleMemberId,
    rateDuration: nullableNumber(fieldValue(fields, "rateDuration", "")),
    curveDuration: nullableNumber(fieldValue(fields, "curveDuration", "")),
    creditDuration: nullableNumber(fieldValue(fields, "creditDuration", "")),
    ktb3yRangeLow: nullableNumber(fieldValue(fields, "ktb3yRangeLow", "")),
    ktb3yRangeHigh: nullableNumber(fieldValue(fields, "ktb3yRangeHigh", "")),
    curveSpreadRangeLow: nullableNumber(fieldValue(fields, "curveSpreadRangeLow", "")),
    curveSpreadRangeHigh: nullableNumber(fieldValue(fields, "curveSpreadRangeHigh", "")),
    creditSpreadRangeLow: nullableNumber(fieldValueAny(fields, ["creditSpreadRangeLow", "creditRangeLow"], "")),
    creditSpreadRangeHigh: nullableNumber(fieldValueAny(fields, ["creditSpreadRangeHigh", "creditRangeHigh"], "")),
    confidence: nullableNumber(fieldValue(fields, "confidence", "")) ?? nullableNumber(Math.round(avg(confidenceValues))),
    rateConfidence: nullableNumber(fieldValue(fields, "rateConfidence", "")),
    curveConfidence: nullableNumber(fieldValue(fields, "curveConfidence", "")),
    creditConfidence: nullableNumber(fieldValue(fields, "creditConfidence", "")),
    rationaleTags: textListField(fields, "rationaleTags"),
    rationaleText: String(fieldValue(fields, "rationaleText", "") || ""),
    rateRationaleText: String(fieldValue(fields, "rateRationaleText", "") || ""),
    curveRationaleText: String(fieldValue(fields, "curveRationaleText", "") || ""),
    creditRationaleText: String(fieldValue(fields, "creditRationaleText", "") || ""),
    createdAt: String(fieldValue(fields, "createdAt", "") || ""),
    updatedAt: String(fieldValue(fields, "updatedAt", "") || "")
  };
}

function sharePointOpinionFields(opinion) {
  return {
    Title: `${opinion.week}_${opinion.memberId}`,
    appId: opinion.id || `${opinion.week}_${opinion.memberId}`,
    week: opinion.week,
    memberId: opinion.memberId,
    rateDuration: sharePointNumber(opinion.rateDuration),
    curveDuration: sharePointNumber(opinion.curveDuration),
    creditDuration: sharePointNumber(opinion.creditDuration),
    ktb3yRangeLow: sharePointNumber(opinion.ktb3yRangeLow),
    ktb3yRangeHigh: sharePointNumber(opinion.ktb3yRangeHigh),
    curveSpreadRangeLow: sharePointNumber(opinion.curveSpreadRangeLow),
    curveSpreadRangeHigh: sharePointNumber(opinion.curveSpreadRangeHigh),
    creditSpreadRangeLow: sharePointNumber(opinion.creditSpreadRangeLow),
    creditSpreadRangeHigh: sharePointNumber(opinion.creditSpreadRangeHigh),
    confidence: sharePointNumber(opinion.confidence),
    rateConfidence: sharePointNumber(opinion.rateConfidence),
    curveConfidence: sharePointNumber(opinion.curveConfidence),
    creditConfidence: sharePointNumber(opinion.creditConfidence),
    rationaleTags: Array.isArray(opinion.rationaleTags) ? opinion.rationaleTags.join("|") : String(opinion.rationaleTags || ""),
    rationaleText: opinion.rationaleText || "",
    rateRationaleText: opinion.rateRationaleText || "",
    curveRationaleText: opinion.curveRationaleText || "",
    creditRationaleText: opinion.creditRationaleText || "",
    createdAt: opinion.createdAt || new Date().toISOString(),
    updatedAt: opinion.updatedAt || new Date().toISOString()
  };
}

async function saveSharePointWeeklyOpinion(opinion) {
  const site = await getSharePointSite();
  const lists = await getSharePointListsByName(site.id);
  const list = lists.get("FI_WeeklyOpinions");
  if (!list) throw new Error("list-not-found:FI_WeeklyOpinions");
  const columns = await getSharePointColumns(site.id, list);
  const items = await readSharePointListItems(site.id, list);
  const sampleFieldKeys = [...new Set(items.flatMap(item => Object.keys(item.fields || {})))];
  const weekFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "week");
  const memberIdFieldNames = ["memberId", "memberid", "mmberId", "mmberid"].map(name =>
    resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, name)
  );
  const existing = items.find(item => {
    const fields = item.fields || {};
    const title = `${opinion.week}_${opinion.memberId}`;
    const existingMemberId = fieldValueAny(fields, memberIdFieldNames, "");
    return (
      fieldValue(fields, "Title") === title ||
      fieldValue(fields, "appId") === opinion.id ||
      (fieldValue(fields, weekFieldName) === opinion.week && existingMemberId === opinion.memberId)
    );
  });
  const fields = omitSharePointFields(
    applySharePointFieldAliases(
      resolveSharePointFields(columns, sharePointOpinionFields(opinion), sampleFieldKeys),
      {
        creditSpreadRangeLow: "creditRangeLow",
        creditSpreadRangeHigh: "creditRangeHigh"
      }
    ),
    ["confidence", "rationaleTags", "rationaleText"]
  );
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

function mapSharePointWeeklySummary(item) {
  const fields = item.fields || {};
  const appId = String(fieldValue(fields, "appId", "") || "").trim();
  const week = String(fieldValue(fields, "week", fieldValue(fields, "Title", "")) || "").trim();
  return {
    id: appId || `sp-summary-${item.id}`,
    week,
    memberId: SUMMARY_MEMBER_ID,
    rateDuration: nullableNumber(fieldValue(fields, "rateDuration", "")),
    curveDuration: nullableNumber(fieldValue(fields, "curveDuration", "")),
    creditDuration: nullableNumber(fieldValue(fields, "creditDuration", "")),
    ktb3yRangeLow: nullableNumber(fieldValue(fields, "ktb3yRangeLow", "")),
    ktb3yRangeHigh: nullableNumber(fieldValue(fields, "ktb3yRangeHigh", "")),
    curveSpreadRangeLow: nullableNumber(fieldValue(fields, "curveSpreadRangeLow", "")),
    curveSpreadRangeHigh: nullableNumber(fieldValue(fields, "curveSpreadRangeHigh", "")),
    creditSpreadRangeLow: nullableNumber(fieldValueAny(fields, ["creditSpreadRangeLow", "creditRangeLow"], "")),
    creditSpreadRangeHigh: nullableNumber(fieldValueAny(fields, ["creditSpreadRangeHigh", "creditRangeHigh"], "")),
    confidence: nullableNumber(fieldValue(fields, "confidence", "")),
    rateConfidence: nullableNumber(fieldValue(fields, "rateConfidence", "")),
    curveConfidence: nullableNumber(fieldValue(fields, "curveConfidence", "")),
    creditConfidence: nullableNumber(fieldValue(fields, "creditConfidence", "")),
    overallSummaryText: String(fieldValue(fields, "overallSummaryText", "") || ""),
    rateSummaryText: String(fieldValue(fields, "rateSummaryText", "") || ""),
    curveSummaryText: String(fieldValue(fields, "curveSummaryText", "") || ""),
    creditSummaryText: String(fieldValue(fields, "creditSummaryText", "") || ""),
    status: String(fieldValue(fields, "status", "completed") || "completed"),
    draftSource: String(fieldValue(fields, "draftSource", "") || ""),
    completedAt: String(fieldValue(fields, "completedAt", "") || ""),
    createdAt: String(fieldValue(fields, "createdAt", "") || ""),
    updatedAt: String(fieldValue(fields, "updatedAt", "") || "")
  };
}

function sharePointWeeklySummaryFields(summary) {
  return {
    Title: summary.week,
    appId: summary.id || summary.week,
    week: summary.week,
    rateDuration: Number(summary.rateDuration) || 0,
    curveDuration: Number(summary.curveDuration) || 0,
    creditDuration: Number(summary.creditDuration) || 0,
    ktb3yRangeLow: Number(summary.ktb3yRangeLow) || 0,
    ktb3yRangeHigh: Number(summary.ktb3yRangeHigh) || 0,
    curveSpreadRangeLow: Number(summary.curveSpreadRangeLow) || 0,
    curveSpreadRangeHigh: Number(summary.curveSpreadRangeHigh) || 0,
    creditSpreadRangeLow: Number(summary.creditSpreadRangeLow) || 0,
    creditSpreadRangeHigh: Number(summary.creditSpreadRangeHigh) || 0,
    confidence: Number(summary.confidence) || 0,
    rateConfidence: Number(summary.rateConfidence) || 0,
    curveConfidence: Number(summary.curveConfidence) || 0,
    creditConfidence: Number(summary.creditConfidence) || 0,
    overallSummaryText: summary.overallSummaryText || "",
    rateSummaryText: summary.rateSummaryText || "",
    curveSummaryText: summary.curveSummaryText || "",
    creditSummaryText: summary.creditSummaryText || "",
    status: summary.status || "completed",
    draftSource: summary.draftSource || "",
    completedAt: summary.completedAt || "",
    createdAt: summary.createdAt || new Date().toISOString(),
    updatedAt: summary.updatedAt || new Date().toISOString()
  };
}

async function ensureSharePointWeeklySummaryStorage(siteId) {
  const lists = await getSharePointListsByName(siteId);
  const list = await ensureSharePointArchiveList(siteId, lists, WEEKLY_SUMMARY_LIST_NAME);
  await ensureSharePointArchiveColumns(siteId, list, WEEKLY_SUMMARY_SCHEMA);
  return list;
}

async function saveSharePointWeeklySummary(summary) {
  const site = await getSharePointSite();
  const list = await ensureSharePointWeeklySummaryStorage(site.id);
  const columns = await getSharePointColumns(site.id, list);
  const items = await readSharePointListItems(site.id, list);
  const sampleFieldKeys = [...new Set(items.flatMap(item => Object.keys(item.fields || {})))];
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const weekFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "week");
  const existing = items.find(item => {
    const fields = item.fields || {};
    return (
      fieldValue(fields, appIdFieldName) === summary.id ||
      fieldValue(fields, weekFieldName) === summary.week ||
      fieldValue(fields, "Title") === summary.week
    );
  });
  const fields = resolveSharePointFields(columns, sharePointWeeklySummaryFields(summary), sampleFieldKeys);
  if (existing) {
    await writeSharePointArchiveFields(site.id, list.id, fields, existing.id);
    return { mode: "updated", itemId: existing.id };
  }
  const created = await writeSharePointArchiveFields(site.id, list.id, fields);
  return { mode: "created", itemId: created.id };
}

function mapSharePointMarket(item) {
  const fields = item.fields || {};
  return normalizeMarket({
    week: String(fieldValue(fields, "week", fieldValue(fields, "Title", "")) || "").trim(),
    asOf: normalizeDateText(fieldValue(fields, "asOf", "")),
    ktb3y: numberField(fields, "ktb3y"),
    ktb10y: numberField(fields, "ktb10y"),
    curveSpread: fieldValue(fields, "curveSpread", ""),
    msb2y: numberField(fields, "msb2y"),
    creditAA2y: numberField(fields, "creditAAm2y"),
    creditSpread: fieldValue(fields, "creditSpread", ""),
    memo: String(fieldValue(fields, "memo", "") || "")
  });
}

function sharePointMarketFields(market) {
  return {
    Title: market.week,
    week: market.week,
    asOf: market.asOf,
    ktb3y: Number(market.ktb3y) || 0,
    ktb10y: Number(market.ktb10y) || 0,
    curveSpread: Number(market.curveSpread) || 0,
    msb2y: Number(market.msb2y) || 0,
    creditAA2y: Number(market.creditAA2y) || 0,
    creditSpread: Number(market.creditSpread) || 0,
    memo: market.memo || ""
  };
}

async function saveSharePointMarket(market) {
  const site = await getSharePointSite();
  const lists = await getSharePointListsByName(site.id);
  const list = lists.get("FI_MarketData");
  if (!list) throw new Error("list-not-found:FI_MarketData");
  const columns = await getSharePointColumns(site.id, list);
  const items = await readSharePointListItems(site.id, list);
  const sampleFieldKeys = [...new Set(items.flatMap(item => Object.keys(item.fields || {})))];
  const weekFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "week");
  const existing = items.find(item => {
    const fields = item.fields || {};
    return fieldValue(fields, "Title") === market.week || fieldValue(fields, weekFieldName) === market.week;
  });
  const fields = applySharePointFieldAliases(
    resolveSharePointFields(columns, sharePointMarketFields(market), sampleFieldKeys),
    { creditAA2y: "creditAAm2y" }
  );
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

async function getSharePointListContext(listName, top = 5000) {
  const site = await getSharePointSite();
  const lists = await getSharePointListsByName(site.id);
  const list = lists.get(listName);
  if (!list) throw new Error(`list-not-found:${listName}`);
  const columns = await getSharePointColumns(site.id, list);
  const items = await readSharePointListItems(site.id, list, top);
  const sampleFieldKeys = [...new Set(items.flatMap(item => Object.keys(item.fields || {})))];
  return { site, list, columns, items, sampleFieldKeys };
}

function sharePointColumnDefinition(name, type) {
  const column = { name };
  if (type === "number") column.number = {};
  else if (type === "boolean") column.boolean = {};
  else if (type === "multiline") column.text = { allowMultipleLines: true };
  else column.text = {};
  return column;
}

async function createSharePointList(siteId, listName) {
  return graphFetch(
    `/sites/${encodeURIComponent(siteId)}/lists`,
    {
      method: "POST",
      body: {
        displayName: listName,
        list: { template: "genericList" }
      }
    }
  );
}

async function ensureSharePointArchiveList(siteId, lists, listName) {
  if (lists.has(listName)) return lists.get(listName);
  const created = await createSharePointList(siteId, listName);
  const list = {
    id: created.id,
    displayName: created.displayName || listName,
    webUrl: created.webUrl || ""
  };
  lists.set(listName, list);
  return list;
}

async function ensureSharePointArchiveColumns(siteId, list, schema) {
  if (!schema?.length) return;
  let columns = await getSharePointColumns(siteId, list);
  const sampleFieldKeys = columns.map(column => column.name).filter(Boolean);
  for (const [name, type] of schema) {
    if (sharePointFieldExists(columns, sampleFieldKeys, name)) continue;
    try {
      await graphFetch(
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(list.id)}/columns`,
        { method: "POST", body: sharePointColumnDefinition(name, type) }
      );
      columns = await getSharePointColumns(siteId, list);
      sampleFieldKeys.splice(0, sampleFieldKeys.length, ...columns.map(column => column.name).filter(Boolean));
    } catch (error) {
      console.warn(`Archive column create skipped: ${list.displayName}.${name}`, error);
    }
  }
}

async function ensureSharePointArchiveStorage(siteId) {
  const lists = await getSharePointListsByName(siteId);
  for (const [listName, schema] of Object.entries(ARCHIVE_LIST_SCHEMAS)) {
    try {
      const list = await ensureSharePointArchiveList(siteId, lists, listName);
      await ensureSharePointArchiveColumns(siteId, list, schema);
    } catch (error) {
      console.warn(`Archive storage ensure skipped: ${listName}`, error);
    }
  }
}

function mapSharePointDailyRate(item) {
  const fields = item.fields || {};
  return normalizeDailyRate({
    date: normalizeDateText(fieldValue(fields, "date", fieldValue(fields, "Title", ""))),
    ktb3y: numberField(fields, "ktb3y"),
    ktb10y: numberField(fields, "ktb10y"),
    msb2y: numberField(fields, "msb2y"),
    creditAA2y: numberFieldAny(fields, ["creditAA2y", "creditAAm2y"]),
    curveSpread: fieldValue(fields, "curveSpread", ""),
    creditSpread: fieldValue(fields, "creditSpread", "")
  });
}

function sharePointDailyRateFields(rate) {
  return {
    Title: rate.date,
    date: rate.date,
    ktb3y: Number(rate.ktb3y) || 0,
    ktb10y: Number(rate.ktb10y) || 0,
    curveSpread: Number(rate.curveSpread) || 0,
    msb2y: Number(rate.msb2y) || 0,
    creditAA2y: Number(rate.creditAA2y) || 0,
    creditSpread: Number(rate.creditSpread) || 0
  };
}

async function saveSharePointDailyRates(rates) {
  const { site, list, columns, items, sampleFieldKeys } = await getSharePointListContext("FI_DailyRates");
  const dateFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "date");
  const existingByDate = new Map(items.map(item => {
    const fields = item.fields || {};
    return [String(fieldValue(fields, dateFieldName, fieldValue(fields, "Title", "")) || ""), item];
  }));
  let created = 0;
  let updated = 0;
  for (const rate of rates) {
    const fields = applySharePointFieldAliases(
      resolveSharePointFields(columns, sharePointDailyRateFields(rate), sampleFieldKeys),
      { creditAA2y: "creditAAm2y" }
    );
    const existing = existingByDate.get(rate.date);
    if (existing) {
      await graphFetch(
        `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
        { method: "PATCH", body: fields }
      );
      updated += 1;
    } else {
      const item = await graphFetch(
        `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
        { method: "POST", body: { fields } }
      );
      existingByDate.set(rate.date, item);
      created += 1;
    }
  }
  return { created, updated, total: rates.length };
}

function mapSharePointWeekMapping(item) {
  const fields = item.fields || {};
  return normalizeWeekMapping({
    week: String(fieldValue(fields, "week", fieldValue(fields, "Title", "")) || "").trim(),
    baseDate: normalizeDateText(fieldValue(fields, "baseDate", "")),
    memo: String(fieldValue(fields, "memo", "") || ""),
    createdAt: String(fieldValue(fields, "createdAt", "") || ""),
    updatedAt: String(fieldValue(fields, "updatedAt", "") || "")
  });
}

function sharePointWeekMappingFields(mapping) {
  return {
    Title: mapping.week,
    week: mapping.week,
    baseDate: mapping.baseDate,
    memo: mapping.memo || "",
    createdAt: mapping.createdAt || new Date().toISOString(),
    updatedAt: mapping.updatedAt || new Date().toISOString()
  };
}

async function saveSharePointWeekMapping(mapping) {
  const { site, list, columns, items, sampleFieldKeys } = await getSharePointListContext("FI_MarketWeekMappings");
  const weekFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "week");
  const existing = items.find(item => {
    const fields = item.fields || {};
    return fieldValue(fields, "Title") === mapping.week || fieldValue(fields, weekFieldName) === mapping.week;
  });
  const fields = resolveSharePointFields(columns, sharePointWeekMappingFields(mapping), sampleFieldKeys);
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

function mapSharePointMappingHistory(item) {
  const fields = item.fields || {};
  const title = String(fieldValue(fields, "Title", "") || "").trim();
  return {
    id: String(fieldValue(fields, "appId", title || `sp-market-map-${item.id}`) || "").trim(),
    week: String(fieldValue(fields, "week", "") || "").trim(),
    previousBaseDate: normalizeDateText(fieldValue(fields, "previousBaseDate", "")),
    baseDate: normalizeDateText(fieldValue(fields, "baseDate", "")),
    memo: String(fieldValue(fields, "memo", "") || ""),
    changedAt: String(fieldValue(fields, "changedAt", "") || "")
  };
}

function sharePointMappingHistoryFields(history) {
  const title = history.id || `${history.week}_${history.baseDate}_${history.changedAt || ""}`;
  return {
    Title: title,
    appId: history.id || title,
    week: history.week,
    previousBaseDate: history.previousBaseDate || "",
    baseDate: history.baseDate,
    memo: history.memo || "",
    changedAt: history.changedAt || new Date().toISOString()
  };
}

async function saveSharePointMappingHistory(history) {
  const { site, list, columns, items, sampleFieldKeys } = await getSharePointListContext("FI_MarketMappingHistory");
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const title = history.id || `${history.week}_${history.baseDate}_${history.changedAt || ""}`;
  const existing = items.find(item => {
    const fields = item.fields || {};
    return fieldValue(fields, appIdFieldName) === history.id || fieldValue(fields, "Title") === title;
  });
  const fields = resolveSharePointFields(columns, sharePointMappingHistoryFields(history), sampleFieldKeys);
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

function mapSharePointUploadedRateFile(item) {
  const fields = item.fields || {};
  const title = String(fieldValue(fields, "Title", "") || "").trim();
  return {
    id: String(fieldValue(fields, "appId", `sp-rate-file-${item.id}`) || "").trim(),
    fileName: String(fieldValue(fields, "fileName", title) || "").trim(),
    uploadedAt: String(fieldValue(fields, "uploadedAt", "") || ""),
    rowCount: numberField(fields, "rowCount"),
    minDate: normalizeDateText(fieldValue(fields, "minDate", "")),
    maxDate: normalizeDateText(fieldValue(fields, "maxDate", "")),
    fileUrl: String(fieldValue(fields, "fileUrl", "") || "")
  };
}

function sharePointUploadedRateFileFields(file) {
  return {
    Title: file.fileName || file.id,
    appId: file.id,
    fileName: file.fileName || "",
    uploadedAt: file.uploadedAt || new Date().toISOString(),
    rowCount: Number(file.rowCount) || 0,
    minDate: file.minDate || "",
    maxDate: file.maxDate || "",
    fileUrl: file.fileUrl || ""
  };
}

async function saveSharePointUploadedRateFile(file) {
  const { site, list, columns, items, sampleFieldKeys } = await getSharePointListContext("FI_UploadedRateFiles");
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const fileNameFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "fileName");
  const uploadedAtFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "uploadedAt");
  const existing = items.find(item => {
    const fields = item.fields || {};
    return (
      fieldValue(fields, appIdFieldName) === file.id ||
      (fieldValue(fields, fileNameFieldName) === file.fileName && fieldValue(fields, uploadedAtFieldName) === file.uploadedAt)
    );
  });
  const fields = resolveSharePointFields(columns, sharePointUploadedRateFileFields(file), sampleFieldKeys);
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

function mapSharePointSettings(item) {
  const fields = item.fields || {};
  const title = String(fieldValue(fields, "Title", "") || "").trim();
  const appId = String(fieldValue(fields, "appId", title) || "").trim();
  if (title !== "default" && appId !== "default") return null;
  return normalizeSettings({
    pnlUnit: fieldValue(fields, "pnlUnit", DEFAULT_SETTINGS.pnlUnit),
    evaluationLagWeeks: fieldValue(fields, "evaluationLagWeeks", DEFAULT_SETTINGS.evaluationLagWeeks)
  });
}

function sharePointSettingsFields(settings) {
  const normalized = normalizeSettings(settings);
  return {
    Title: "default",
    appId: "default",
    pnlUnit: normalized.pnlUnit,
    evaluationLagWeeks: normalized.evaluationLagWeeks,
    updatedAt: new Date().toISOString()
  };
}

async function saveSharePointSettings(settings) {
  const { site, list, columns, items, sampleFieldKeys } = await getSharePointListContext("FI_Settings");
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const existing = items.find(item => {
    const fields = item.fields || {};
    return fieldValue(fields, "Title") === "default" || fieldValue(fields, appIdFieldName) === "default";
  });
  const fields = resolveSharePointFields(columns, sharePointSettingsFields(settings), sampleFieldKeys);
  if (existing) {
    await graphFetch(
      `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(existing.id)}/fields`,
      { method: "PATCH", body: fields }
    );
    return { mode: "updated", itemId: existing.id };
  }
  const created = await graphFetch(
    `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items`,
    { method: "POST", body: { fields } }
  );
  return { mode: "created", itemId: created.id };
}

function mapSharePointArchiveHeader(item) {
  const fields = item.fields || {};
  const title = String(fieldValue(fields, "Title", "") || "").trim();
  return {
    archiveId: String(fieldValue(fields, "archiveId", fieldValue(fields, "appId", title || `sp-archive-${item.id}`)) || "").trim(),
    week: String(fieldValue(fields, "week", title) || "").trim(),
    version: numberField(fields, "version") || 1,
    status: String(fieldValue(fields, "status", "active") || "active"),
    archivedAt: String(fieldValue(fields, "archivedAt", "") || ""),
    archivedBy: String(fieldValue(fields, "archivedBy", "") || ""),
    memo: String(fieldValue(fields, "memo", "") || ""),
    calculationVersion: numberField(fields, "calculationVersion") || ARCHIVE_CALCULATION_VERSION,
    sourceHash: String(fieldValue(fields, "sourceHash", "") || "")
  };
}

function archiveHeaderFields(archive) {
  return {
    Title: archive.week,
    appId: archive.archiveId,
    archiveId: archive.archiveId,
    week: archive.week,
    version: Number(archive.version) || 1,
    status: archive.status || "active",
    archivedAt: archive.archivedAt || new Date().toISOString(),
    archivedBy: archive.archivedBy || "",
    memo: archive.memo || "",
    calculationVersion: Number(archive.calculationVersion) || ARCHIVE_CALCULATION_VERSION,
    sourceHash: archive.sourceHash || ""
  };
}

function mapSharePointArchiveMember(item) {
  const fields = item.fields || {};
  return {
    id: String(fieldValue(fields, "appId", `sp-archive-member-${item.id}`) || ""),
    archiveId: String(fieldValue(fields, "archiveId", "") || ""),
    week: String(fieldValue(fields, "week", "") || ""),
    version: numberField(fields, "version") || 1,
    memberId: String(fieldValue(fields, "memberId", "") || ""),
    memberName: String(fieldValue(fields, "memberName", fieldValue(fields, "Title", "")) || ""),
    team: String(fieldValue(fields, "team", "") || ""),
    active: booleanField(fields, "active", true),
    sortOrder: numberField(fields, "sortOrder")
  };
}

function archiveMemberFields(row) {
  return {
    Title: `${row.week}_${row.memberId}`,
    appId: row.id,
    archiveId: row.archiveId,
    week: row.week,
    version: Number(row.version) || 1,
    memberId: row.memberId,
    memberName: row.memberName || "",
    team: row.team || "",
    active: row.active !== false,
    sortOrder: Number(row.sortOrder) || 0
  };
}

function mapSharePointArchiveOpinionStrategy(item) {
  const fields = item.fields || {};
  return {
    id: String(fieldValue(fields, "appId", `sp-archive-opinion-${item.id}`) || ""),
    archiveId: String(fieldValue(fields, "archiveId", "") || ""),
    week: String(fieldValue(fields, "week", "") || ""),
    version: numberField(fields, "version") || 1,
    sourceOpinionId: String(fieldValue(fields, "sourceOpinionId", "") || ""),
    memberId: String(fieldValue(fields, "memberId", "") || ""),
    memberName: String(fieldValue(fields, "memberName", "") || ""),
    team: String(fieldValue(fields, "team", "") || ""),
    strategyKey: String(fieldValue(fields, "strategyKey", "") || ""),
    duration: numberField(fields, "duration"),
    rangeLow: numberField(fields, "rangeLow"),
    rangeHigh: numberField(fields, "rangeHigh"),
    confidence: numberField(fields, "confidence"),
    rationaleText: String(fieldValue(fields, "rationaleText", "") || ""),
    rationaleTags: textListField(fields, "rationaleTags"),
    sourceCreatedAt: String(fieldValue(fields, "sourceCreatedAt", "") || ""),
    sourceUpdatedAt: String(fieldValue(fields, "sourceUpdatedAt", "") || "")
  };
}

function archiveOpinionStrategyFields(row) {
  return {
    Title: `${row.week}_${row.memberId}_${row.strategyKey}`,
    appId: row.id,
    archiveId: row.archiveId,
    week: row.week,
    version: Number(row.version) || 1,
    sourceOpinionId: row.sourceOpinionId || "",
    memberId: row.memberId,
    memberName: row.memberName || "",
    team: row.team || "",
    strategyKey: row.strategyKey,
    duration: Number(row.duration) || 0,
    rangeLow: Number(row.rangeLow) || 0,
    rangeHigh: Number(row.rangeHigh) || 0,
    confidence: Number(row.confidence) || 0,
    rationaleText: row.rationaleText || "",
    rationaleTags: Array.isArray(row.rationaleTags) ? row.rationaleTags.join("|") : String(row.rationaleTags || ""),
    sourceCreatedAt: row.sourceCreatedAt || "",
    sourceUpdatedAt: row.sourceUpdatedAt || ""
  };
}

function mapSharePointArchiveMarket(item) {
  const fields = item.fields || {};
  return {
    id: String(fieldValue(fields, "appId", `sp-archive-market-${item.id}`) || ""),
    archiveId: String(fieldValue(fields, "archiveId", "") || ""),
    week: String(fieldValue(fields, "week", "") || ""),
    version: numberField(fields, "version") || 1,
    marketRole: String(fieldValue(fields, "marketRole", "") || ""),
    marketWeek: String(fieldValue(fields, "marketWeek", "") || ""),
    asOf: normalizeDateText(fieldValue(fields, "asOf", "")),
    baseDate: normalizeDateText(fieldValue(fields, "baseDate", "")),
    ktb3y: numberField(fields, "ktb3y"),
    ktb10y: numberField(fields, "ktb10y"),
    curveSpread: numberField(fields, "curveSpread"),
    msb2y: numberField(fields, "msb2y"),
    creditAA2y: numberFieldAny(fields, ["creditAA2y", "creditAAm2y"]),
    creditSpread: numberField(fields, "creditSpread"),
    memo: String(fieldValue(fields, "memo", "") || "")
  };
}

function archiveMarketFields(row) {
  return {
    Title: `${row.week}_${row.marketRole}`,
    appId: row.id,
    archiveId: row.archiveId,
    week: row.week,
    version: Number(row.version) || 1,
    marketRole: row.marketRole,
    marketWeek: row.marketWeek || "",
    asOf: row.asOf || "",
    baseDate: row.baseDate || row.asOf || "",
    ktb3y: Number(row.ktb3y) || 0,
    ktb10y: Number(row.ktb10y) || 0,
    curveSpread: Number(row.curveSpread) || 0,
    msb2y: Number(row.msb2y) || 0,
    creditAA2y: Number(row.creditAA2y) || 0,
    creditSpread: Number(row.creditSpread) || 0,
    memo: row.memo || ""
  };
}

function mapSharePointArchiveResult(item) {
  const fields = item.fields || {};
  return {
    id: String(fieldValue(fields, "appId", `sp-archive-result-${item.id}`) || ""),
    archiveId: String(fieldValue(fields, "archiveId", "") || ""),
    week: String(fieldValue(fields, "week", "") || ""),
    version: numberField(fields, "version") || 1,
    sourceOpinionId: String(fieldValue(fields, "sourceOpinionId", "") || ""),
    memberId: String(fieldValue(fields, "memberId", "") || ""),
    memberName: String(fieldValue(fields, "memberName", "") || ""),
    strategyKey: String(fieldValue(fields, "strategyKey", "") || ""),
    evaluatedWeek: String(fieldValue(fields, "evaluatedWeek", "") || ""),
    marketDelta: numberField(fields, "marketDelta"),
    pnl: numberField(fields, "pnl"),
    ytdPnl: numberField(fields, "ytdPnl"),
    rangeHit: booleanField(fields, "rangeHit", false)
  };
}

function archiveResultFields(row) {
  return {
    Title: `${row.week}_${row.memberId}_${row.strategyKey}`,
    appId: row.id,
    archiveId: row.archiveId,
    week: row.week,
    version: Number(row.version) || 1,
    sourceOpinionId: row.sourceOpinionId || "",
    memberId: row.memberId,
    memberName: row.memberName || "",
    strategyKey: row.strategyKey,
    evaluatedWeek: row.evaluatedWeek || "",
    marketDelta: nullableNumber(row.marketDelta),
    pnl: nullableNumber(row.pnl),
    ytdPnl: nullableNumber(row.ytdPnl),
    rangeHit: row.rangeHit === true
  };
}

function mapSharePointArchiveRawJson(item) {
  const fields = item.fields || {};
  return {
    id: String(fieldValue(fields, "appId", `sp-archive-raw-${item.id}`) || ""),
    archiveId: String(fieldValue(fields, "archiveId", "") || ""),
    week: String(fieldValue(fields, "week", "") || ""),
    version: numberField(fields, "version") || 1,
    json: String(fieldValue(fields, "json", "") || "")
  };
}

function archiveRawJsonFields(row) {
  return {
    Title: `${row.week}_v${row.version}`,
    appId: row.id,
    archiveId: row.archiveId,
    week: row.week,
    version: Number(row.version) || 1,
    json: row.json || ""
  };
}

async function saveSharePointArchiveRows(listName, rows, fieldsFactory) {
  const { site, list, columns, items, sampleFieldKeys } = await getSharePointListContext(listName);
  const appIdFieldName = resolveSharePointFieldNameFromSources(columns, sampleFieldKeys, "appId");
  const existingByAppId = new Map(items.map(item => [String(fieldValue(item.fields || {}, appIdFieldName, "")), item]));
  let created = 0;
  let updated = 0;
  for (const row of rows) {
    const fields = applySharePointFieldAliases(
      resolveExistingSharePointFields(columns, fieldsFactory(row), sampleFieldKeys),
      { creditAA2y: "creditAAm2y" }
    );
    const existing = existingByAppId.get(row.id || row.archiveId);
    if (existing) {
      await writeSharePointArchiveFields(site.id, list.id, fields, existing.id);
      updated += 1;
    } else {
      const item = await writeSharePointArchiveFields(site.id, list.id, fields);
      existingByAppId.set(row.id || row.archiveId, item);
      created += 1;
    }
  }
  return { listName, created, updated, total: rows.length };
}

function invalidSharePointFieldName(error) {
  const text = String(error?.message || "");
  return text.match(/Field '([^']+)' is not recognized/)?.[1] || "";
}

async function writeSharePointArchiveFields(siteId, listId, fields, itemId = "") {
  const bodyFields = { ...fields };
  for (let attempt = 0; attempt < 12; attempt += 1) {
    try {
      if (itemId) {
        await graphFetch(
          `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/fields`,
          { method: "PATCH", body: bodyFields }
        );
        return { id: itemId };
      }
      return graphFetch(
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`,
        { method: "POST", body: { fields: bodyFields } }
      );
    } catch (error) {
      const invalidField = invalidSharePointFieldName(error);
      if (!invalidField || !(invalidField in bodyFields) || invalidField === "Title") throw error;
      delete bodyFields[invalidField];
    }
  }
  throw new Error("sharepoint-archive-too-many-invalid-fields");
}

async function saveSharePointArchiveSnapshot(snapshot) {
  const site = await getSharePointSite();
  await ensureSharePointArchiveStorage(site.id);
  const rowsByList = [
    [ARCHIVE_LIST_NAMES.archives, state.weeklyArchives.filter(item => item.week === snapshot.header.week), archiveHeaderFields],
    [ARCHIVE_LIST_NAMES.members, snapshot.members, archiveMemberFields],
    [ARCHIVE_LIST_NAMES.opinionStrategies, snapshot.opinionStrategies, archiveOpinionStrategyFields],
    [ARCHIVE_LIST_NAMES.markets, snapshot.markets, archiveMarketFields],
    [ARCHIVE_LIST_NAMES.results, snapshot.results, archiveResultFields],
    [ARCHIVE_LIST_NAMES.rawJson, snapshot.rawJson, archiveRawJsonFields]
  ];
  const results = [];
  const skipped = [];
  for (const [listName, rows, fieldsFactory] of rowsByList) {
    try {
      results.push(await saveSharePointArchiveRows(listName, rows, fieldsFactory));
    } catch (error) {
      if (String(error.message || "").includes(`list-not-found:${listName}`)) {
        skipped.push(listName);
      } else {
        throw error;
      }
    }
  }
  return { results, skipped };
}

async function readOptionalArchiveList(siteId, lists, listName, mapper) {
  if (!lists.has(listName)) return [];
  const items = await readSharePointListItems(siteId, lists.get(listName), 5000);
  return items.map(mapper).filter(Boolean);
}

async function loadSharePointState() {
  try {
    setMicrosoftStatus("SharePoint 데이터 불러오는 중...");
    await signInMicrosoft();
    const site = await getSharePointSite();
    const lists = await getSharePointListsByName(site.id);
    const requiredLists = [
      "DAOL_FI_Members",
      "FI_WeeklyOpinions",
      "FI_MarketData",
      "FI_DailyRates",
      "FI_MarketWeekMappings",
      "FI_MarketMappingHistory",
      "FI_UploadedRateFiles"
    ];
    const missing = requiredLists.filter(name => !lists.has(name));
    if (missing.length) throw new Error(`list-not-found:${missing.join(",")}`);

    const [
      memberItems,
      opinionItems,
      marketItems,
      dailyRateItems,
      weekMappingItems,
      mappingHistoryItems,
      uploadedRateFileItems,
      settingsItems,
      archiveItems,
      archiveMemberItems,
      archiveOpinionStrategyItems,
      archiveMarketItems,
      archiveResultItems,
      archiveRawJsonItems,
      weeklySummaryItems
    ] = await Promise.all([
      readSharePointListItems(site.id, lists.get("DAOL_FI_Members")),
      readSharePointListItems(site.id, lists.get("FI_WeeklyOpinions")),
      readSharePointListItems(site.id, lists.get("FI_MarketData")),
      readSharePointListItems(site.id, lists.get("FI_DailyRates"), 5000),
      readSharePointListItems(site.id, lists.get("FI_MarketWeekMappings"), 5000),
      readSharePointListItems(site.id, lists.get("FI_MarketMappingHistory"), 5000),
      readSharePointListItems(site.id, lists.get("FI_UploadedRateFiles"), 5000),
      lists.has("FI_Settings") ? readSharePointListItems(site.id, lists.get("FI_Settings")) : Promise.resolve([]),
      readOptionalArchiveList(site.id, lists, ARCHIVE_LIST_NAMES.archives, mapSharePointArchiveHeader),
      readOptionalArchiveList(site.id, lists, ARCHIVE_LIST_NAMES.members, mapSharePointArchiveMember),
      readOptionalArchiveList(site.id, lists, ARCHIVE_LIST_NAMES.opinionStrategies, mapSharePointArchiveOpinionStrategy),
      readOptionalArchiveList(site.id, lists, ARCHIVE_LIST_NAMES.markets, mapSharePointArchiveMarket),
      readOptionalArchiveList(site.id, lists, ARCHIVE_LIST_NAMES.results, mapSharePointArchiveResult),
      readOptionalArchiveList(site.id, lists, ARCHIVE_LIST_NAMES.rawJson, mapSharePointArchiveRawJson),
      lists.has(WEEKLY_SUMMARY_LIST_NAME)
        ? readSharePointListItems(site.id, lists.get(WEEKLY_SUMMARY_LIST_NAME), 5000)
        : Promise.resolve([])
    ]);
    const sharePointSettings = settingsItems.map(mapSharePointSettings).find(Boolean);

    const nextState = normalizeState({
      ...state,
      members: memberItems.map(mapSharePointMember).filter(member => member.id && member.name),
      weeklyOpinions: opinionItems.map(mapSharePointOpinion).filter(opinion => opinion.week && opinion.memberId),
      marketData: marketItems.map(mapSharePointMarket).filter(market => market.week && market.asOf),
      dailyRates: dailyRateItems.map(mapSharePointDailyRate).filter(Boolean),
      marketWeekMappings: weekMappingItems.map(mapSharePointWeekMapping).filter(Boolean),
      marketMappingHistory: mappingHistoryItems.map(mapSharePointMappingHistory).filter(item => item.week && item.baseDate),
      uploadedRateFiles: uploadedRateFileItems.map(mapSharePointUploadedRateFile).filter(item => item.id && item.fileName),
      settings: sharePointSettings || state.settings,
      weeklyArchives: archiveItems.filter(item => item.archiveId && item.week),
      weeklyArchiveMembers: archiveMemberItems.filter(item => item.archiveId && item.memberId),
      weeklyArchiveOpinionStrategies: archiveOpinionStrategyItems.filter(item => item.archiveId && item.memberId && item.strategyKey),
      weeklyArchiveMarkets: archiveMarketItems.filter(item => item.archiveId && item.marketRole),
      weeklyArchiveResults: archiveResultItems.filter(item => item.archiveId && item.memberId && item.strategyKey),
      weeklyArchiveRawJson: archiveRawJsonItems.filter(item => item.archiveId),
      weeklySummaries: weeklySummaryItems.map(mapSharePointWeeklySummary).filter(item => item.week)
    });
    state = nextState;
    saveState();
    if (!sharePointSettings && lists.has("FI_Settings")) {
      saveSharePointSettings(state.settings).catch(error => console.warn("FI_Settings default upsert failed", error));
    }
    if (!state.weeklyOpinions.some(item => item.week === selectedWeek) && !state.marketData.some(item => item.week === selectedWeek)) {
      selectedWeek = latestWeek();
    }
    render();
    const rawApplied = await loadBundledRatesIfAvailable({ showStatus: false });
    setMicrosoftStatus(
      `SharePoint 불러오기 성공: members ${state.members.length}, opinions ${state.weeklyOpinions.length}, market ${state.marketData.length}, rates ${state.dailyRates.length}${rawApplied ? " (local raw applied)" : ""}, mappings ${state.marketWeekMappings.length}, settings ${sharePointSettings ? "loaded" : "local"}`,
      "connected"
    );
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`불러오기 실패: ${shortErrorMessage(error)}`, "error");
    alert(`SharePoint 데이터 불러오기 실패: ${error.message}\n\n기존 로컬 화면은 유지됩니다.`);
  }
}

async function testSharePointConnection() {
  try {
    setMicrosoftStatus("SharePoint 연결 확인 중...");
    await signInMicrosoft();
    const site = await getSharePointSite();
    const lists = await getSharePointListsByName(site.id);
    const list = lists.get(MICROSOFT_CONFIG.testListName);
    if (!list) throw new Error(`list-not-found:${MICROSOFT_CONFIG.testListName}`);
    const items = await readSharePointListItems(site.id, list, 5);
    const count = items.length || 0;
    setMicrosoftStatus(`${MICROSOFT_CONFIG.testListName} ${count}건 읽기 성공`, "connected");
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`연결 실패: ${shortErrorMessage(error)}`, "error");
    alert(`SharePoint 연결 테스트 실패: ${error.message}`);
  }
}

function id(prefix) {
  return `${prefix}-${crypto.randomUUID()}`;
}

function round(value, digits = 2) {
  const factor = 10 ** digits;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function fmt(value, digits = 2) {
  const number = finiteNumber(value);
  return Number.isFinite(number) ? number.toFixed(digits) : "-";
}

function fmtSigned(value, digits = 2) {
  const n = finiteNumber(value);
  if (!Number.isFinite(n)) return "-";
  return `${n > 0 ? "+" : ""}${n.toFixed(digits)}`;
}

function avg(values) {
  const valid = values.map(Number).filter(Number.isFinite);
  if (!valid.length) return NaN;
  return valid.reduce((sum, value) => sum + value, 0) / valid.length;
}

function median(values) {
  const valid = values.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
  if (!valid.length) return NaN;
  const mid = Math.floor(valid.length / 2);
  return valid.length % 2 ? valid[mid] : (valid[mid - 1] + valid[mid]) / 2;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formData(form) {
  return Object.fromEntries(new FormData(form).entries());
}

function setStrategyOmitState(strategyKey, skipped) {
  if (!els.opinionForm) return;
  const strategy = STRATEGIES.find(item => item.key === strategyKey);
  if (!strategy) return;
  const form = els.opinionForm.elements;
  const names = [strategy.durationKey, strategy.rangeLowKey, strategy.rangeHighKey,
                 `${strategy.key}Confidence`, `${strategy.key}RangeSkip`];
  names.forEach(name => {
    const field = form[name];
    if (!field) return;
    field.required = false;
    if (skipped) {
      if (field.type === "checkbox") field.checked = false;
      else field.value = "";
    }
    field.classList.toggle("input-locked", skipped);
  });
}

function setRangeOmitState(strategyKey, skipped) {
  if (!els.opinionForm) return;
  const strategy = STRATEGIES.find(item => item.key === strategyKey);
  if (!strategy) return;
  const form = els.opinionForm.elements;
  [strategy.rangeLowKey, strategy.rangeHighKey].forEach(name => {
    const field = form[name];
    if (!field) return;
    field.required = false;
    if (skipped) field.value = "";
    field.classList.toggle("input-locked", skipped);
  });
}

function updateOpinionStrategyControls() {
  STRATEGIES.forEach(strategy => {
    const status = els.opinionForm?.elements[`${strategy.key}Status`]?.value || "active";
    const stratSkipped = status === "skip";
    setStrategyOmitState(strategy.key, stratSkipped);
    if (!stratSkipped) {
      const rangeSkip = !!els.opinionForm?.elements[`${strategy.key}RangeSkip`]?.checked;
      setRangeOmitState(strategy.key, rangeSkip);
    }
  });
}

function setDurationSelectValue(selectEl, value) {
  if (!selectEl) return;
  selectEl.querySelectorAll("option[data-transient]").forEach(option => option.remove());
  const num = nullableNumber(value);
  if (num === null) { selectEl.value = ""; return; }
  const canonical = String(num);
  if (![...selectEl.options].some(option => option.value === canonical)) {
    const option = document.createElement("option");
    option.value = canonical;
    option.textContent = `${fmtSigned(num, 2)} (비표준)`;
    option.dataset.transient = "1";
    selectEl.appendChild(option);
  }
  selectEl.value = canonical;
}

function applyOpinionToForm(opinion) {
  if (!els.opinionForm) return;
  const form = els.opinionForm.elements;
  form.week.value = opinion?.week || selectedOpinionWeek;
  form.memberId.value = opinion?.memberId || form.memberId.value;
  STRATEGIES.forEach(strategy => {
    const durNull  = nullableNumber(opinion?.[strategy.durationKey]) === null;
    const lowNull  = nullableNumber(opinion?.[strategy.rangeLowKey]) === null;
    const highNull = nullableNumber(opinion?.[strategy.rangeHighKey]) === null;
    const skipped  = !opinion || (durNull && lowNull && highNull);
    if (form[`${strategy.key}Status`]) form[`${strategy.key}Status`].value = skipped ? "skip" : "active";
    setDurationSelectValue(form[strategy.durationKey], opinion?.[strategy.durationKey]);
    form[strategy.rangeHighKey].value = opinion?.[strategy.rangeHighKey] ?? "";
    form[strategy.rangeLowKey].value  = opinion?.[strategy.rangeLowKey] ?? "";
    if (form[`${strategy.key}RangeSkip`]) {
      form[`${strategy.key}RangeSkip`].checked = !!opinion && !skipped && lowNull && highNull;
    }
    form[`${strategy.key}Confidence`].value = opinion?.[`${strategy.key}Confidence`] ?? "";
    form[`${strategy.key}RationaleText`].value = opinion ? strategyRationale(opinion, strategy.key) : "";
  });
  updateOpinionStrategyControls();
}

function selectedOpinionMemberName() {
  const text = els.opinionMember?.selectedOptions?.[0]?.textContent || "";
  return text.split(/[·\u00b7ㆍ∙•]/)[0].trim();
}

function normalizeLookupName(value) {
  return String(value || "").normalize("NFC").replace(/\s+/g, "").trim();
}

function selectedOpinionMemberLookupNames(memberId) {
  return [
    memberById(memberId).name,
    selectedOpinionMemberName(),
    els.opinionMember?.selectedOptions?.[0]?.textContent || ""
  ]
    .map(normalizeLookupName)
    .filter(Boolean);
}

function findOpinionForSelection(week, memberId) {
  const exact = state.weeklyOpinions.find(row => row.week === week && row.memberId === memberId);
  if (exact) return exact;
  const lookupNames = selectedOpinionMemberLookupNames(memberId);
  if (!lookupNames.length) return null;
  return state.weeklyOpinions.find(row =>
    row.week === week &&
    lookupNames.includes(normalizeLookupName(memberById(row.memberId).name))
  ) || null;
}

function opinionFormHasUnsavedInput() {
  if (!els.opinionForm) return false;
  const data = formData(els.opinionForm);
  return STRATEGIES.some(strategy => {
    if (data[`${strategy.key}Status`] === "skip") return false;
    const num = [strategy.durationKey, strategy.rangeLowKey, strategy.rangeHighKey]
      .some(key => String(data[key] ?? "").trim() !== "");
    const text = String(data[`${strategy.key}RationaleText`] ?? "").trim() !== "";
    return num || text;
  });
}

function clearOpinionInputs(keepSelection = true) {
  const week = els.opinionWeek?.value || selectedOpinionWeek;
  const memberId = els.opinionMember?.value || "";
  els.opinionForm.reset();
  renderFormsDefaults();
  if (keepSelection) {
    els.opinionWeek.value = week;
    els.opinionMember.value = memberId;
    selectedOpinionWeek = week;
  }
  STRATEGIES.forEach(strategy => {
    if (els.opinionForm.elements[`${strategy.key}Status`]) {
      els.opinionForm.elements[`${strategy.key}Status`].value = "active";
    }
    setDurationSelectValue(els.opinionForm.elements[strategy.durationKey], null);
    els.opinionForm.elements[`${strategy.key}Confidence`].value = "3";
  });
  updateOpinionStrategyControls();
}

function loadSelectedOpinionIntoForm({ notify = true } = {}) {
  const week = els.opinionWeek?.value || selectedOpinionWeek;
  const memberId = els.opinionMember?.value || "";
  const opinion = findOpinionForSelection(week, memberId);
  const memberName = selectedOpinionMemberName() || memberById(memberId).name;
  if (opinion) {
    const formOpinion = opinion.memberId && ![...els.opinionMember.options].some(option => option.value === opinion.memberId)
      ? { ...opinion, memberId }
      : opinion;
    applyOpinionToForm(formOpinion);
    if (notify) alert(`${weekDisplayLabel(week)} / ${memberName} 의견을 불러왔습니다.`);
    return opinion;
  }
  clearOpinionInputs(true);
  if (notify) alert(`${weekDisplayLabel(week)} / ${memberName}에 저장된 의견이 없습니다.`);
  return null;
}

function updateOpinionLookupStatus(message = "") {
  if (!els.opinionLookupStatus) return;
  if (message) {
    els.opinionLookupStatus.textContent = message;
    return;
  }
  const week = els.opinionWeek?.value || selectedOpinionWeek;
  const rows = opinionsForWeek(week);
  const memberNames = rows.map(row => memberById(row.memberId).name).filter(Boolean);
  els.opinionLookupStatus.textContent = rows.length
    ? `${weekDisplayLabel(week)} 저장 의견 ${rows.length}건: ${memberNames.join(", ")}`
    : `${weekDisplayLabel(week)} 저장 의견 없음. 필요하면 Login/Load 후 조회하세요.`;
}

function numericInputValue(value) {
  return nullableNumber(value);
}

function strategyIsSkipped(data, strategy) {
  return data[`${strategy.key}Status`] === "skip";
}

function rangeIsSkipped(data, strategy) {
  return data[`${strategy.key}RangeSkip`] === "on";
}

function finiteNumber(value) {
  const number = nullableNumber(value);
  return number === null ? NaN : number;
}

function normalizeDateText(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") return excelDateToIso(value);
  const text = String(value).trim();
  if (/^\d+(\.\d+)?$/.test(text)) return excelDateToIso(Number(text));
  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function excelDateToIso(serial) {
  const days = Number(serial);
  if (!Number.isFinite(days)) return "";
  const date = new Date(Date.UTC(1899, 11, 30 + Math.floor(days)));
  return date.toISOString().slice(0, 10);
}

function getCurrentWeek(date = new Date()) {
  const copy = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = copy.getUTCDay() || 7;
  copy.setUTCDate(copy.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(copy.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((copy - yearStart) / 86400000) + 1) / 7);
  return `${copy.getUTCFullYear()}-W${String(weekNo).padStart(2, "0")}`;
}

function weekStartDate(week) {
  const match = String(week || "").match(/^(\d{4})-W(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const weekNo = Number(match[2]);
  if (!Number.isFinite(year) || !Number.isFinite(weekNo)) return null;
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const day = jan4.getUTCDay() || 7;
  jan4.setUTCDate(jan4.getUTCDate() - day + 1 + ((weekNo - 1) * 7));
  return jan4;
}

function addWeeksToWeek(week, count) {
  const date = weekStartDate(week);
  if (!date) return getCurrentWeek();
  date.setUTCDate(date.getUTCDate() + (Number(count) || 0) * 7);
  return getCurrentWeek(date);
}

function meetingDateForWeek(week) {
  const date = weekStartDate(week);
  if (!date) return null;
  date.setUTCDate(date.getUTCDate() - 7);
  return date;
}

function formatDateDot(date) {
  if (!date || Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10).replaceAll("-", ".");
}

function meetingNumber(week) {
  const date = meetingDateForWeek(week);
  const firstDate = meetingDateForWeek(FIRST_MEETING_WEEK);
  if (!date || !firstDate) return "";
  const meetingIndex = Math.round((date - firstDate) / (7 * 86400000)) + 1;
  const yearText = String(date.getUTCFullYear()).slice(2);
  return Number.isFinite(meetingIndex) ? `${yearText}\uB144 ${meetingIndex}\uBC88\uC9F8 \uD68C\uC758` : "";
}

function weekDisplayLabel(week) {
  const dateText = formatDateDot(meetingDateForWeek(week));
  const meetingText = meetingNumber(week);
  return [dateText, meetingText].filter(Boolean).join(" \u00B7 ") || String(week || "");
}
function weekShortLabel(week) {
  return formatDateDot(meetingDateForWeek(week)) || String(week || "");
}

function nextOpinionWeek() {
  return addWeeksToWeek(getCurrentWeek(), 1);
}

function opinionWeekOptions() {
  const start = nextOpinionWeek();
  const futureWeeks = Array.from({ length: 12 }, (_, index) => addWeeksToWeek(start, index));
  const savedOpinionWeeks = state.weeklyOpinions.map(item => item.week);
  const savedSummaryWeeks = state.weeklySummaries.map(item => item.week);
  const mappedWeeks = state.marketWeekMappings.map(item => item.week);
  return [...new Set([...savedOpinionWeeks, ...savedSummaryWeeks, ...mappedWeeks, ...futureWeeks, selectedOpinionWeek])]
    .filter(Boolean)
    .sort();
}

function weeklySummaryWeekOptions() {
  return [...new Set([
    ...sortedWeeks(),
    ...state.weeklySummaries.map(item => item.week),
    selectedWeeklySummaryWeek,
    nextOpinionWeek()
  ])].filter(Boolean).sort().reverse();
}

function sortedWeeks() {
  return [...new Set([
    ...state.weeklyOpinions.map(item => item.week),
    ...state.weeklySummaries.map(item => item.week),
    ...state.marketData.map(item => item.week),
    ...state.marketWeekMappings.map(item => item.week),
    nextOpinionWeek()
  ])].filter(Boolean).sort();
}

function latestWeek() {
  return sortedWeeks().at(-1) || nextOpinionWeek();
}

function activeMembers() {
  return [...state.members]
    .filter(member => member.active !== false)
    .sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0) || a.name.localeCompare(b.name));
}

function orderedMembers() {
  return [...state.members]
    .sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0) || a.name.localeCompare(b.name));
}

function memberById(memberId) {
  if (memberId === SUMMARY_MEMBER_ID) return SUMMARY_MEMBER;
  return state.members.find(member => member.id === memberId) || { name: "미등록", team: "-" };
}

function opinionsForWeek(week) {
  return state.weeklyOpinions.filter(item => item.week === week);
}

function weeklySummaryForWeek(week) {
  return state.weeklySummaries.find(item => item.week === week);
}

function summaryPerformanceRows() {
  return state.weeklySummaries
    .filter(item => item.status === "completed")
    .filter(item => STRATEGIES.every(strategy =>
      Number.isFinite(finiteNumber(item[strategy.durationKey])) &&
      Number.isFinite(finiteNumber(item[strategy.rangeLowKey])) &&
      Number.isFinite(finiteNumber(item[strategy.rangeHighKey]))
    ))
    .map(item => ({
      ...item,
      memberId: SUMMARY_MEMBER_ID,
      rationaleText: item.overallSummaryText || "",
      rateRationaleText: item.rateSummaryText || "",
      curveRationaleText: item.curveSummaryText || "",
      creditRationaleText: item.creditSummaryText || ""
    }));
}

function performanceOpinions() {
  return [...state.weeklyOpinions, ...summaryPerformanceRows()];
}

function marketByWeek(week) {
  const saved = state.marketData.find(item => item.week === week);
  if (saved) return saved;
  const mapping = state.marketWeekMappings.find(item => item.week === week);
  const daily = mapping ? dailyRateByDate(mapping.baseDate) : null;
  return daily ? dailyRateToMarket(week, daily, mapping.memo) : undefined;
}

function dailyRateByDate(date) {
  return state.dailyRates.find(item => item.date === date);
}

function dailyRateToMarket(week, daily, memo = "") {
  return makeMarket(week, daily.date, daily.ktb3y, daily.ktb10y, daily.msb2y, daily.creditAA2y, memo);
}

function marketDateForWeek(week) {
  const mapping = state.marketWeekMappings.find(item => item.week === week);
  if (mapping?.baseDate) return mapping.baseDate;
  return marketByWeek(week)?.asOf || "";
}

function addDaysIso(dateText, days) {
  const date = new Date(`${dateText}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function shortDateLabel(dateText) {
  return String(dateText || "").slice(5).replace("-", ".");
}

function nextWeekAfter(week) {
  const weeks = state.marketData.map(item => item.week).sort();
  const index = weeks.indexOf(week);
  return index >= 0 ? weeks[index + 1] : undefined;
}

function previousWeekBefore(week) {
  const weeks = state.marketData.map(item => item.week).sort();
  const index = weeks.indexOf(week);
  return index > 0 ? weeks[index - 1] : undefined;
}

function marketDelta(week, strategy) {
  const start = marketByWeek(week);
  const end = marketByWeek(nextWeekAfter(week));
  if (!start || !end) return NaN;
  if (strategy.key === "rate") {
    return round((Number(end.ktb3y) - Number(start.ktb3y)) * 100, 2);
  }
  return round(Number(end[strategy.marketKey]) - Number(start[strategy.marketKey]), 2);
}

function strategyPnl(opinion, strategy) {
  const delta = marketDelta(opinion.week, strategy);
  const duration = finiteNumber(opinion[strategy.durationKey]);
  if (!Number.isFinite(delta) || !Number.isFinite(duration)) return NaN;
  return round(-strategy.multiplier * duration * delta, 3);
}

function totalPnl(opinion) {
  const values = STRATEGIES.map(strategy => strategyPnl(opinion, strategy));
  if (values.some(value => !Number.isFinite(value))) return NaN;
  return round(values.reduce((sum, value) => sum + value, 0), 3);
}

function strategyRationale(opinion, strategyKey) {
  return opinion[`${strategyKey}RationaleText`] || opinion.rationaleText || "";
}

function strategyConfidence(opinion, strategyKey) {
  const value = Number(opinion[`${strategyKey}Confidence`] ?? opinion.confidence);
  return Number.isFinite(value) ? value : NaN;
}

function strategyYtdPnl(opinion, strategy) {
  const rows = evaluatedRows(opinion.week.slice(0, 4))
    .filter(row => row.opinion.memberId === opinion.memberId && row.opinion.week <= opinion.week);
  const values = rows.map(row => row.pnls[strategy.key]).filter(Number.isFinite);
  return values.length ? round(values.reduce((sum, value) => sum + value, 0), 2) : NaN;
}

function evaluatedRows(year = selectedWeek.slice(0, 4)) {
  return performanceOpinions()
    .filter(opinion => opinion.week.startsWith(year))
    .map(opinion => {
      const pnls = Object.fromEntries(STRATEGIES.map(strategy => [strategy.key, strategyPnl(opinion, strategy)]));
      return {
        opinion,
        member: memberById(opinion.memberId),
        nextWeek: nextWeekAfter(opinion.week),
        pnls,
        total: Object.values(pnls).some(value => !Number.isFinite(value)) ? NaN : round(Object.values(pnls).reduce((sum, value) => sum + value, 0), 3)
      };
    })
    .filter(row => Number.isFinite(row.total));
}

function rankingRows(sortMetric = "rate") {
  const rows = evaluatedRows();
  const byMember = new Map();
  state.members.forEach(member => {
    byMember.set(member.id, {
      member,
      rate: 0,
      curve: 0,
      credit: 0,
      total: 0,
      count: 0
    });
  });
  if (rows.some(row => row.opinion.memberId === SUMMARY_MEMBER_ID)) {
    byMember.set(SUMMARY_MEMBER_ID, {
      member: SUMMARY_MEMBER,
      rate: 0,
      curve: 0,
      credit: 0,
      total: 0,
      count: 0
    });
  }
  rows.forEach(row => {
    const bucket = byMember.get(row.opinion.memberId);
    if (!bucket) return;
    STRATEGIES.forEach(strategy => {
      bucket[strategy.key] += row.pnls[strategy.key];
    });
    bucket.total += row.total;
    bucket.count += 1;
  });
  return [...byMember.values()]
    .map(item => ({
      ...item,
      rate: round(item.rate, 2),
      curve: round(item.curve, 2),
      credit: round(item.credit, 2),
      total: round(item.total, 2)
    }))
    .sort((a, b) => b[sortMetric] - a[sortMetric]);
}

function weeklyPerformanceRows() {
  const rows = evaluatedRows();
  const weeks = [...new Set(rows.map(row => row.opinion.week))].sort();
  if (!weeks.length) return { week: "", rows: [] };
  const targetWeek = rows.some(row => row.opinion.week === selectedWeek)
    ? selectedWeek
    : [...weeks].reverse().find(week => week <= selectedWeek) || weeks.at(-1);
  return {
    week: targetWeek,
    rows: rows
      .filter(row => row.opinion.week === targetWeek)
      .sort((a, b) => b.pnls[selectedPerformanceMetric] - a.pnls[selectedPerformanceMetric])
  };
}

function latestArchiveForWeek(week) {
  return state.weeklyArchives
    .filter(item => item.week === week)
    .sort((a, b) => Number(b.version || 0) - Number(a.version || 0))[0];
}

function activeArchiveForWeek(week) {
  return state.weeklyArchives
    .filter(item => item.week === week && item.status === "active")
    .sort((a, b) => Number(b.version || 0) - Number(a.version || 0))[0] || latestArchiveForWeek(week);
}

function archiveById(archiveId) {
  return state.weeklyArchives.find(item => item.archiveId === archiveId);
}

function selectedArchive() {
  const archives = [...state.weeklyArchives]
    .sort((a, b) => b.week.localeCompare(a.week) || Number(b.version || 0) - Number(a.version || 0));
  if (!archives.length) return undefined;
  const explicit = selectedArchiveId ? archiveById(selectedArchiveId) : undefined;
  if (explicit) return explicit;
  return activeArchiveForWeek(selectedWeek) || archives[0];
}

function simpleHash(text) {
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = ((hash << 5) - hash + text.charCodeAt(i)) | 0;
  }
  return String(hash >>> 0);
}

function archivedByName() {
  return msAccount?.name || msAccount?.username || "local";
}

function strategyRangeHit(opinion, strategy, market) {
  const actual = finiteNumber(market?.[strategy.marketKey]);
  const low = finiteNumber(opinion[strategy.rangeLowKey]);
  const high = finiteNumber(opinion[strategy.rangeHighKey]);
  return Number.isFinite(actual) && Number.isFinite(low) && Number.isFinite(high) && actual >= low && actual <= high;
}

function buildWeeklyArchiveSnapshot(week, memo = "") {
  const archivedAt = new Date().toISOString();
  const previousVersion = Math.max(0, ...state.weeklyArchives.filter(item => item.week === week).map(item => Number(item.version) || 0));
  const version = previousVersion + 1;
  const archiveId = `archive_${week}_v${version}`;
  const opinions = opinionsForWeek(week);
  const memberIds = new Set(opinions.map(opinion => opinion.memberId));
  const members = state.members
    .filter(member => memberIds.has(member.id) || member.active !== false)
    .map(member => ({
      id: `${archiveId}_member_${member.id}`,
      archiveId,
      week,
      version,
      memberId: member.id,
      memberName: member.name,
      team: member.team,
      active: member.active !== false,
      sortOrder: Number(member.sortOrder) || 0
    }));
  const baseMarket = marketByWeek(week);
  const evaluationWeek = nextWeekAfter(week);
  const evaluationMarket = evaluationWeek ? marketByWeek(evaluationWeek) : undefined;
  const mapping = state.marketWeekMappings.find(item => item.week === week);
  const evaluationMapping = evaluationWeek ? state.marketWeekMappings.find(item => item.week === evaluationWeek) : undefined;
  const markets = [
    baseMarket ? {
      id: `${archiveId}_market_base`,
      archiveId,
      week,
      version,
      marketRole: "base",
      marketWeek: week,
      asOf: baseMarket.asOf,
      baseDate: mapping?.baseDate || baseMarket.asOf,
      ktb3y: baseMarket.ktb3y,
      ktb10y: baseMarket.ktb10y,
      curveSpread: baseMarket.curveSpread,
      msb2y: baseMarket.msb2y,
      creditAA2y: baseMarket.creditAA2y,
      creditSpread: baseMarket.creditSpread,
      memo: baseMarket.memo || mapping?.memo || ""
    } : null,
    evaluationMarket ? {
      id: `${archiveId}_market_evaluation`,
      archiveId,
      week,
      version,
      marketRole: "evaluation",
      marketWeek: evaluationWeek,
      asOf: evaluationMarket.asOf,
      baseDate: evaluationMapping?.baseDate || evaluationMarket.asOf,
      ktb3y: evaluationMarket.ktb3y,
      ktb10y: evaluationMarket.ktb10y,
      curveSpread: evaluationMarket.curveSpread,
      msb2y: evaluationMarket.msb2y,
      creditAA2y: evaluationMarket.creditAA2y,
      creditSpread: evaluationMarket.creditSpread,
      memo: evaluationMarket.memo || evaluationMapping?.memo || ""
    } : null
  ].filter(Boolean);
  const opinionStrategies = opinions.flatMap(opinion => {
    const member = memberById(opinion.memberId);
    return STRATEGIES.map(strategy => ({
      id: `${archiveId}_opinion_${opinion.memberId}_${strategy.key}`,
      archiveId,
      week,
      version,
      sourceOpinionId: opinion.id || "",
      memberId: opinion.memberId,
      memberName: member.name,
      team: member.team,
      strategyKey: strategy.key,
      duration: nullableNumber(opinion[strategy.durationKey]),
      rangeLow: nullableNumber(opinion[strategy.rangeLowKey]),
      rangeHigh: nullableNumber(opinion[strategy.rangeHighKey]),
      confidence: nullableNumber(strategyConfidence(opinion, strategy.key)),
      rationaleText: strategyRationale(opinion, strategy.key),
      rationaleTags: Array.isArray(opinion.rationaleTags) ? opinion.rationaleTags : [],
      sourceCreatedAt: opinion.createdAt || "",
      sourceUpdatedAt: opinion.updatedAt || ""
    }));
  });
  const results = opinions.flatMap(opinion => {
    const member = memberById(opinion.memberId);
    return STRATEGIES.map(strategy => ({
      id: `${archiveId}_result_${opinion.memberId}_${strategy.key}`,
      archiveId,
      week,
      version,
      sourceOpinionId: opinion.id || "",
      memberId: opinion.memberId,
      memberName: member.name,
      strategyKey: strategy.key,
      evaluatedWeek: evaluationWeek || "",
      marketDelta: marketDelta(week, strategy),
      pnl: strategyPnl(opinion, strategy),
      ytdPnl: strategyYtdPnl(opinion, strategy),
      rangeHit: strategyRangeHit(opinion, strategy, baseMarket)
    }));
  });
  const rawPayload = {
    archiveId,
    week,
    version,
    archivedAt,
    archivedBy: archivedByName(),
    memo,
    source: {
      members: state.members,
      weeklyOpinions: opinions,
      marketData: [baseMarket, evaluationMarket].filter(Boolean),
      marketWeekMappings: [mapping, evaluationMapping].filter(Boolean)
    },
    analysis: { members, opinionStrategies, markets, results }
  };
  const rawJsonText = JSON.stringify(rawPayload);
  const header = {
    archiveId,
    week,
    version,
    status: "active",
    archivedAt,
    archivedBy: archivedByName(),
    memo,
    calculationVersion: ARCHIVE_CALCULATION_VERSION,
    sourceHash: simpleHash(rawJsonText)
  };
  return {
    header,
    members,
    opinionStrategies,
    markets,
    results,
    rawJson: [{
      id: `${archiveId}_raw`,
      archiveId,
      week,
      version,
      json: JSON.stringify({ ...rawPayload, header })
    }]
  };
}

function upsertLocalArchiveSnapshot(snapshot) {
  state.weeklyArchives = state.weeklyArchives.map(item =>
    item.week === snapshot.header.week && item.status === "active"
      ? { ...item, status: "superseded" }
      : item
  );
  state.weeklyArchives.push(snapshot.header);
  state.weeklyArchiveMembers.push(...snapshot.members);
  state.weeklyArchiveOpinionStrategies.push(...snapshot.opinionStrategies);
  state.weeklyArchiveMarkets.push(...snapshot.markets);
  state.weeklyArchiveResults.push(...snapshot.results);
  state.weeklyArchiveRawJson.push(...snapshot.rawJson);
}

function performanceMetricLabel(metricKey) {
  if (metricKey === "rate") return "금리 방향성";
  if (metricKey === "curve") return "커브";
  if (metricKey === "credit") return "크레딧 스프레드 성과";
  return "금리 방향성";
}

function consensus(week) {
  const rows = opinionsForWeek(week);
  const result = { count: rows.length };
  STRATEGIES.forEach(strategy => {
    result[strategy.key] = {
      avgDuration: avg(rows.map(item => item[strategy.durationKey])),
      medDuration: median(rows.map(item => item[strategy.durationKey])),
      avgLow: avg(rows.map(item => item[strategy.rangeLowKey])),
      avgHigh: avg(rows.map(item => item[strategy.rangeHighKey])),
      medLow: median(rows.map(item => item[strategy.rangeLowKey])),
      medHigh: median(rows.map(item => item[strategy.rangeHighKey]))
    };
  });
  return result;
}

function strategyDirectionLabel(strategy, value) {
  const number = Number(value);
  if (!Number.isFinite(number) || Math.abs(number) < 0.005) return "중립";
  if (strategy.key === "rate") return number > 0 ? "금리 하락 우위" : "금리 상승 경계";
  if (strategy.key === "curve") return number > 0 ? "커브 축소 우위" : "커브 확대 우위";
  return number > 0 ? "크레딧 스프레드 축소 우위" : "크레딧 스프레드 확대 경계";
}

function rationaleSnippet(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 140);
}

function strategyDraftText(week, strategy) {
  const rows = opinionsForWeek(week);
  if (!rows.length) return "";
  const c = consensus(week)[strategy.key];
  const digits = strategy.key === "rate" ? 3 : 1;
  const positive = rows.filter(row => finiteNumber(row[strategy.durationKey]) > 0).length;
  const negative = rows.filter(row => finiteNumber(row[strategy.durationKey]) < 0).length;
  const neutral = rows.length - positive - negative;
  const rationales = rows
    .map(row => {
      const member = memberById(row.memberId);
      const text = rationaleSnippet(strategyRationale(row, strategy.key));
      return text ? `- ${member.name}: ${text}` : "";
    })
    .filter(Boolean)
    .join("\n");
  return [
    `${strategy.title}은 ${strategyDirectionLabel(strategy, c.avgDuration)}로 종합됩니다.`,
    `평균 입력값은 ${fmtSigned(c.avgDuration, 2)}이고, 중앙값은 ${fmtSigned(c.medDuration, 2)}입니다. 예상 레인지는 평균 기준 ${fmt(c.avgLow, digits)}~${fmt(c.avgHigh, digits)}${strategy.rangeUnit}입니다.`,
    `의견 분포는 플러스 ${positive}명, 마이너스 ${negative}명, 중립 ${neutral}명입니다.`,
    rationales ? `주요 근거:\n${rationales}` : ""
  ].filter(Boolean).join("\n\n");
}

function generateWeeklySummaryDraft(week) {
  const rows = opinionsForWeek(week);
  if (!rows.length) return null;
  const c = consensus(week);
  const overallLines = [
    `${week} 종합의견 초안입니다. 총 ${rows.length}명의 의견을 기준으로 작성했습니다.`,
    `금리는 ${strategyDirectionLabel(STRATEGIES[0], c.rate.avgDuration)}, 커브는 ${strategyDirectionLabel(STRATEGIES[1], c.curve.avgDuration)}, 크레딧은 ${strategyDirectionLabel(STRATEGIES[2], c.credit.avgDuration)}로 요약됩니다.`,
    "아래 초안은 각 조직원의 입력값과 판단 근거를 기계적으로 종합한 내용이므로, 확정 전에 문장과 결론을 검토해 주세요."
  ];
  const draftNumber = value => Number.isFinite(finiteNumber(value)) ? finiteNumber(value) : "";
  return {
    overallSummaryText: overallLines.join("\n\n"),
    rateSummaryText: strategyDraftText(week, STRATEGIES[0]),
    curveSummaryText: strategyDraftText(week, STRATEGIES[1]),
    creditSummaryText: strategyDraftText(week, STRATEGIES[2]),
    rateDuration: draftNumber(round(c.rate.avgDuration, 2)),
    curveDuration: draftNumber(round(c.curve.avgDuration, 2)),
    creditDuration: draftNumber(round(c.credit.avgDuration, 2)),
    ktb3yRangeLow: draftNumber(round(c.rate.avgLow, 3)),
    ktb3yRangeHigh: draftNumber(round(c.rate.avgHigh, 3)),
    curveSpreadRangeLow: draftNumber(round(c.curve.avgLow, 1)),
    curveSpreadRangeHigh: draftNumber(round(c.curve.avgHigh, 1)),
    creditSpreadRangeLow: draftNumber(round(c.credit.avgLow, 1)),
    creditSpreadRangeHigh: draftNumber(round(c.credit.avgHigh, 1)),
    rateConfidence: draftNumber(Math.round(avg(rows.map(row => strategyConfidence(row, "rate"))))),
    curveConfidence: draftNumber(Math.round(avg(rows.map(row => strategyConfidence(row, "curve"))))),
    creditConfidence: draftNumber(Math.round(avg(rows.map(row => strategyConfidence(row, "credit")))))
  };
}

function render() {
  renderWeeks();
  renderArchiveState();
  renderMemberSelect();
  renderFormsDefaults();
  renderWeeklySummaryView();
  renderOverview();
  renderDurationCharts();
  renderPerformance();
  renderArchiveView();
  renderRateUploadState();
  renderMarketDatePreview();
  renderMarketList();
  renderMembers();
  renderLastUpdatedBadge();
}

function renderWeeks() {
  const weeks = sortedWeeks().reverse();
  if (!weeks.includes(selectedWeek)) selectedWeek = weeks[0] || getCurrentWeek();
  els.weekSelect.innerHTML = weeks.map(week => `<option value="${week}">${weekDisplayLabel(week)}</option>`).join("");
  els.weekSelect.value = selectedWeek;
}

function renderArchiveState() {
  if (!els.archiveStatus) return;
  const archive = activeArchiveForWeek(selectedWeek);
  if (!archive) {
    els.archiveStatus.innerHTML = `<strong>Archive</strong><span>Not closed</span>`;
    if (els.archiveWeekButton) els.archiveWeekButton.textContent = "Archive week";
    return;
  }
  const archivedAt = archive.archivedAt ? archive.archivedAt.slice(0, 16).replace("T", " ") : "-";
  els.archiveStatus.innerHTML = `
    <strong>Archive v${Number(archive.version) || 1}</strong>
    <span>${escapeHtml(archivedAt)} · ${escapeHtml(archive.archivedBy || "-")}</span>
  `;
  if (els.archiveWeekButton) els.archiveWeekButton.textContent = "Re-archive week";
}

function findSelfMemberByLogin() {
  const upn = String(msAccount?.username || "").trim().toLowerCase();
  if (!upn) return null;
  return state.members.find(member =>
    member.active !== false &&
    String(member.email || "").trim().toLowerCase() === upn
  ) || null;
}

function renderMemberSelect() {
  const self = findSelfMemberByLogin();

  if (self) {
    // 본인 1개 옵션만 → 다른 멤버 선택 불가(잠금 효과). disabled는 쓰지 않는다
    // (disabled면 formData가 memberId를 수집하지 않아 저장이 깨진다).
    els.opinionMember.innerHTML =
      `<option value="${self.id}">${escapeHtml(self.name)} · ${escapeHtml(self.team)}</option>`;
    els.opinionMember.value = self.id;
    els.opinionMember.disabled = false;
    els.opinionMember.classList.add("locked-self");
    return;
  }

  const upn = String(msAccount?.username || "").trim();
  console.warn(
    "[A2] 로그인 계정과 일치하는 멤버를 찾지 못했습니다. " +
    "DAOL_FI_Members의 email 칸이 이 UPN과 정확히 일치하는지 확인하세요. UPN =",
    upn || "(로그인 안 됨)"
  );
  els.opinionMember.innerHTML = upn
    ? `<option value="">계정 매칭 실패 — 관리자에게 email 확인 요청</option>`
    : `<option value="">로그인이 필요합니다</option>`;
  els.opinionMember.value = "";
  els.opinionMember.disabled = false;
  els.opinionMember.classList.remove("locked-self");
}

function renderFormsDefaults() {
  const options = opinionWeekOptions();
  if (!options.includes(selectedOpinionWeek)) selectedOpinionWeek = options[0] || nextOpinionWeek();
  els.opinionWeek.innerHTML = options.map(week => `<option value="${week}">${weekDisplayLabel(week)}</option>`).join("");
  els.opinionWeek.value = selectedOpinionWeek;
  updateOpinionStrategyControls();
  updateOpinionLookupStatus();
  els.marketWeek.innerHTML = sortedWeeks().map(week => `<option value="${week}">${weekDisplayLabel(week)}</option>`).join("");
  els.marketWeek.value = selectedWeek;
  const mapping = state.marketWeekMappings.find(item => item.week === selectedWeek);
  if (mapping) {
    els.marketBaseDate.value = mapping.baseDate;
    els.marketForm.elements.memo.value = mapping.memo || "";
  } else if (!els.marketBaseDate.value) {
    els.marketBaseDate.value = latestDailyRateDate() || new Date().toISOString().slice(0, 10);
  }
}

function renderWeeklySummaryView() {
  if (!els.weeklySummaryForm || !els.weeklySummaryWeek) return;
  const options = weeklySummaryWeekOptions();
  if (!options.includes(selectedWeeklySummaryWeek)) selectedWeeklySummaryWeek = options[0] || selectedWeek || nextOpinionWeek();
  els.weeklySummaryWeek.innerHTML = options.map(week => `<option value="${week}">${weekDisplayLabel(week)}</option>`).join("");
  populateWeeklySummaryForm(selectedWeeklySummaryWeek);
}

function populateWeeklySummaryForm(week) {
  if (!els.weeklySummaryForm || !els.weeklySummaryWeek) return;
  selectedWeeklySummaryWeek = week || selectedWeeklySummaryWeek || selectedWeek;
  const summary = weeklySummaryForWeek(selectedWeeklySummaryWeek);
  els.weeklySummaryWeek.value = selectedWeeklySummaryWeek;
  els.weeklySummaryForm.elements.overallSummaryText.value = summary?.overallSummaryText || "";
  els.weeklySummaryForm.elements.rateDuration.value = summary?.rateDuration ?? "";
  els.weeklySummaryForm.elements.ktb3yRangeHigh.value = summary?.ktb3yRangeHigh ?? "";
  els.weeklySummaryForm.elements.ktb3yRangeLow.value = summary?.ktb3yRangeLow ?? "";
  els.weeklySummaryForm.elements.rateConfidence.value = summary?.rateConfidence || "3";
  els.weeklySummaryForm.elements.rateSummaryText.value = summary?.rateSummaryText || "";
  els.weeklySummaryForm.elements.curveDuration.value = summary?.curveDuration ?? "";
  els.weeklySummaryForm.elements.curveSpreadRangeHigh.value = summary?.curveSpreadRangeHigh ?? "";
  els.weeklySummaryForm.elements.curveSpreadRangeLow.value = summary?.curveSpreadRangeLow ?? "";
  els.weeklySummaryForm.elements.curveConfidence.value = summary?.curveConfidence || "3";
  els.weeklySummaryForm.elements.curveSummaryText.value = summary?.curveSummaryText || "";
  els.weeklySummaryForm.elements.creditDuration.value = summary?.creditDuration ?? "";
  els.weeklySummaryForm.elements.creditSpreadRangeHigh.value = summary?.creditSpreadRangeHigh ?? "";
  els.weeklySummaryForm.elements.creditSpreadRangeLow.value = summary?.creditSpreadRangeLow ?? "";
  els.weeklySummaryForm.elements.creditConfidence.value = summary?.creditConfidence || "3";
  els.weeklySummaryForm.elements.creditSummaryText.value = summary?.creditSummaryText || "";
  if (els.weeklySummaryStatus) {
    const opinionCount = opinionsForWeek(selectedWeeklySummaryWeek).length;
    if (summary) {
      const updatedAt = summary.updatedAt ? summary.updatedAt.slice(0, 16).replace("T", " ") : "-";
      els.weeklySummaryStatus.textContent = `완료됨 · ${updatedAt} · 원 의견 ${opinionCount}건`;
    } else {
      els.weeklySummaryStatus.textContent = `저장된 종합의견 없음 · 원 의견 ${opinionCount}건`;
    }
  }
}

function renderOverview() {
  if (els.topRankingMetricSelect) els.topRankingMetricSelect.value = selectedTopRankingMetric;
  const rows = opinionsForWeek(selectedWeek);
  const c = consensus(selectedWeek);
  els.kpiParticipants.textContent = `${rows.length}명`;
  els.kpiActiveMembers.textContent = `활성 본부원 ${activeMembers().length}명`;
  els.kpiRateDuration.textContent = fmtSigned(c.rate.avgDuration, 2);
  els.kpiCurveDuration.textContent = fmtSigned(c.curve.avgDuration, 2);
  els.kpiCreditDuration.textContent = fmtSigned(c.credit.avgDuration, 2);

  const rankingMetric = selectedTopRankingMetric;
  const ranking = rankingRows(rankingMetric);
  els.topRanking.innerHTML = ranking.slice(0, 3).map((row, index) => `
    <div class="ranking-item">
      <div>
        <strong>${index + 1}. ${escapeHtml(row.member.name)}</strong>
        <span>${escapeHtml(row.member.team)} · 평가 ${row.count}주</span>
      </div>
      <strong class="${scoreClass(row[rankingMetric])}">${fmtSigned(row[rankingMetric], 2)}bp</strong>
    </div>
  `).join("") || emptyBlock("확정된 성과가 없습니다.");

  els.fullRankingTable.innerHTML = ranking.map((row, index) => `
    <tr>
      <td>${index + 1}</td>
      <td><strong>${escapeHtml(row.member.name)}</strong></td>
      <td class="${scoreClass(row.rate)}">${fmtSigned(row.rate, 2)}</td>
      <td class="${scoreClass(row.curve)}">${fmtSigned(row.curve, 2)}</td>
      <td class="${scoreClass(row.credit)}">${fmtSigned(row.credit, 2)}</td>
    </tr>
  `).join("") || emptyRow(5);

  els.weekConsensus.innerHTML = STRATEGIES.map(strategy => {
    const item = c[strategy.key];
    const digits = strategy.key === "rate" ? 3 : 1;
    return `
      <div class="summary-item consensus-item">
        <strong>${strategy.title}</strong>
        <span><b>주간 평균</b> <em>${fmtSigned(item.avgDuration, 2)}</em></span>
        <span><b>중앙값</b> <em>${fmtSigned(item.medDuration, 2)}</em></span>
        <span><b>평균 레인지</b> <em>${fmt(item.avgLow, digits)}~${fmt(item.avgHigh, digits)}${strategy.rangeUnit}</em></span>
      </div>
    `;
  }).join("");

  renderMarketMovement();
}

function renderMarketMovement() {
  const current = marketByWeek(selectedWeek);
  if (!current) {
    els.marketMovementList.innerHTML = emptyBlock("선택 주차의 시장 데이터가 없습니다.");
    return;
  }

  const previousWeek = previousWeekBefore(selectedWeek);
  const previous = previousWeek ? marketByWeek(previousWeek) : undefined;
  const movementRows = [
    {
      label: "국고3년 금리",
      value: current.ktb3y,
      previous: previous?.ktb3y,
      valueText: `${fmt(current.ktb3y, 3)}%`,
      deltaText: previous ? `${fmtSigned((current.ktb3y - previous.ktb3y) * 100, 1)}bp` : "전주 없음"
    },
    {
      label: "국고10년 금리",
      value: current.ktb10y,
      previous: previous?.ktb10y,
      valueText: `${fmt(current.ktb10y, 3)}%`,
      deltaText: previous ? `${fmtSigned((current.ktb10y - previous.ktb10y) * 100, 1)}bp` : "전주 없음"
    },
    {
      label: "국고3년-10년 금리 스프레드",
      value: current.curveSpread,
      previous: previous?.curveSpread,
      valueText: `${fmt(current.curveSpread, 1)}bp`,
      deltaText: previous ? `${fmtSigned(current.curveSpread - previous.curveSpread, 1)}bp` : "전주 없음"
    },
    {
      label: "여전채 AA-2년 - 통안채 2년 스프레드",
      value: current.creditSpread,
      previous: previous?.creditSpread,
      valueText: `${fmt(current.creditSpread, 1)}bp`,
      deltaText: previous ? `${fmtSigned(current.creditSpread - previous.creditSpread, 1)}bp` : "전주 없음"
    }
  ];

  els.marketMovementList.innerHTML = movementRows.map(row => {
    const delta = Number(row.value) - Number(row.previous);
    return `
      <div class="market-movement-item">
        <span>${escapeHtml(row.label)}</span>
        <strong>${row.valueText}</strong>
        <small class="${previous ? scoreClass(delta) : "score-flat"}">전주 대비 ${row.deltaText}</small>
      </div>
    `;
  }).join("");
}

function renderDurationCharts() {
  els.durationCharts.innerHTML = STRATEGIES.map(strategy => chartPanelHtml(`duration-${strategy.key}`, strategy.title, "본부원별 입력값과 시장 데이터를 함께 표시합니다.")).join("");
  STRATEGIES.forEach(strategy => {
    const table = document.querySelector(`#duration-${strategy.key}-table`);
    table.innerHTML = durationTableRows(strategy);
    drawStrategyChart(document.querySelector(`#duration-${strategy.key}`), strategy, strategyChartModes[strategy.key]);
  });
}

function chartPanelHtml(idValue, title, subtitle) {
  const strategyKey = idValue.replace("duration-", "");
  const selectedMode = strategyChartModes[strategyKey] || "duration";
  return `
    <article class="panel">
      <div class="panel-head">
        <div>
          <h3>${title}</h3>
          <p>${subtitle}</p>
        </div>
        <label class="inline-control" for="${idValue}-mode">
          보기
          <select id="${idValue}-mode" data-chart-mode="${strategyKey}">
            <option value="duration"${selectedMode === "duration" ? " selected" : ""}>듀레이션</option>
            <option value="range"${selectedMode === "range" ? " selected" : ""}>레인지</option>
          </select>
        </label>
      </div>
      <div class="chart-wrap">
        <svg id="${idValue}"></svg>
        <div id="${idValue}-tooltip" class="chart-tooltip" role="tooltip"></div>
      </div>
      <div id="${idValue}-table" class="week-accordion"></div>
    </article>
  `;
}

function durationTableRows(strategy) {
  const rows = state.weeklyOpinions
    .filter(opinion => Number.isFinite(finiteNumber(opinion[strategy.durationKey])))
    .sort((a, b) => b.week.localeCompare(a.week));
  const weeks = [...new Set(rows.map(opinion => opinion.week))];

  return weeks.map(week => {
    const weekRows = rows.filter(opinion => opinion.week === week);
    const avgDuration = avg(weekRows.map(opinion => opinion[strategy.durationKey]));
    const avgPnl = avg(weekRows.map(opinion => strategyPnl(opinion, strategy)));
    const detailRows = weekRows.map(opinion => {
      const member = memberById(opinion.memberId);
      const pnl = strategyPnl(opinion, strategy);
      const ytdPnl = strategyYtdPnl(opinion, strategy);
      const confidence = strategyConfidence(opinion, strategy.key);
      return `
        <tr>
          <td>${escapeHtml(member.name)}</td>
          <td>${fmtSigned(opinion[strategy.durationKey], 2)}</td>
          <td>${fmt(opinion[strategy.rangeLowKey], strategy.key === "rate" ? 3 : 1)}~${fmt(opinion[strategy.rangeHighKey], strategy.key === "rate" ? 3 : 1)}${strategy.rangeUnit}</td>
          <td class="rationale-cell">${escapeHtml(strategyRationale(opinion, strategy.key))}</td>
          <td>${Number.isFinite(confidence) ? `<span class="pill">확신도 ${confidence}</span>` : "-"}</td>
          <td class="${scoreClass(pnl)}">${Number.isFinite(pnl) ? `${fmtSigned(pnl, 2)}bp` : "평가 대기"}</td>
          <td class="${scoreClass(ytdPnl)}">${Number.isFinite(ytdPnl) ? `${fmtSigned(ytdPnl, 2)}bp` : "평가 대기"}</td>
        </tr>
      `;
    }).join("");

    return `
      <details class="week-disclosure">
        <summary>
          <span class="week-summary-main">
            <strong>${escapeHtml(weekDisplayLabel(week))}</strong>
            <span>${weekRows.length}명 참여</span>
          </span>
          <span class="week-summary-meta">
            <span>평균 입력 ${fmtSigned(avgDuration, 2)}</span>
            <span class="${scoreClass(avgPnl)}">평균 손익 ${Number.isFinite(avgPnl) ? `${fmtSigned(avgPnl, 2)}bp` : "평가 대기"}</span>
          </span>
        </summary>
        <div class="table-wrap">
          <table class="week-detail-table">
            <thead>
              <tr>
                <th>본부원</th>
                <th>입력값</th>
                <th>레인지</th>
                <th>판단 근거</th>
                <th>확신도</th>
                <th>주간 손익</th>
                <th>YTD 손익</th>
              </tr>
            </thead>
            <tbody>${detailRows}</tbody>
          </table>
        </div>
      </details>
    `;
  }).join("") || `<div class="empty-state">저장된 데이터가 없습니다.</div>`;
}

function rangeTableRows(strategy) {
  const rows = state.weeklyOpinions
    .filter(opinion => Number.isFinite(finiteNumber(opinion[strategy.rangeLowKey])) && Number.isFinite(finiteNumber(opinion[strategy.rangeHighKey])))
    .sort((a, b) => b.week.localeCompare(a.week));
  const weeks = [...new Set(rows.map(opinion => opinion.week))];

  return weeks.map(week => {
    const weekRows = rows.filter(opinion => opinion.week === week);
    const actual = marketByWeek(week)?.[strategy.marketKey];
    const hasActual = Number.isFinite(finiteNumber(actual));
    const hitCount = hasActual
      ? weekRows.filter(opinion => finiteNumber(actual) >= finiteNumber(opinion[strategy.rangeLowKey]) && finiteNumber(actual) <= finiteNumber(opinion[strategy.rangeHighKey])).length
      : 0;
    const actualText = hasActual ? `${fmt(actual, strategy.key === "rate" ? 3 : 1)}${strategy.unit}` : "-";
    const detailRows = weekRows.map(opinion => {
      const member = memberById(opinion.memberId);
      const hit = hasActual && finiteNumber(actual) >= finiteNumber(opinion[strategy.rangeLowKey]) && finiteNumber(actual) <= finiteNumber(opinion[strategy.rangeHighKey]);
      return `
        <tr>
          <td>${escapeHtml(member.name)}</td>
          <td>${actualText}</td>
          <td>${fmt(opinion[strategy.rangeLowKey], strategy.key === "rate" ? 3 : 1)}~${fmt(opinion[strategy.rangeHighKey], strategy.key === "rate" ? 3 : 1)}${strategy.rangeUnit}</td>
          <td class="rationale-cell">${escapeHtml(strategyRationale(opinion, strategy.key))}</td>
          <td>${hasActual ? `<span class="pill">${hit ? "적중" : "이탈"}</span>` : "대기"}</td>
        </tr>
      `;
    }).join("");

    return `
      <details class="week-disclosure">
        <summary>
          <span class="week-summary-main">
            <strong>${escapeHtml(weekDisplayLabel(week))}</strong>
            <span>${weekRows.length}명 참여</span>
          </span>
          <span class="week-summary-meta">
            <span>실제 ${actualText}</span>
            <span>${hasActual ? `적중 ${hitCount}/${weekRows.length}` : "평가 대기"}</span>
          </span>
        </summary>
        <div class="table-wrap">
          <table class="week-detail-table">
            <thead>
              <tr>
                <th>본부원</th>
                <th>실제값</th>
                <th>예상 레인지</th>
                <th>판단 근거</th>
                <th>결과</th>
              </tr>
            </thead>
            <tbody>${detailRows}</tbody>
          </table>
        </div>
      </details>
    `;
  }).join("") || `<div class="empty-state">저장된 데이터가 없습니다.</div>`;
}

function renderPerformance() {
  if (els.performanceMetricSelect) els.performanceMetricSelect.value = selectedPerformanceMetric;
  const weekly = weeklyPerformanceRows();
  els.weeklyPerformanceTable.innerHTML = weekly.rows.map(row => `
    <tr>
      <td><strong>${escapeHtml(row.member.name)}</strong><br><span>${escapeHtml(row.member.team)}</span></td>
      <td class="${scoreClass(row.pnls.rate)}">${fmtSigned(row.pnls.rate, 2)}</td>
      <td class="${scoreClass(row.pnls.curve)}">${fmtSigned(row.pnls.curve, 2)}</td>
      <td class="${scoreClass(row.pnls.credit)}">${fmtSigned(row.pnls.credit, 2)}</td>
      <td>${escapeHtml(row.opinion.week)} → ${escapeHtml(row.nextWeek)}</td>
    </tr>
  `).join("") || emptyRow(5);

  const ranking = rankingRows(selectedPerformanceMetric);
  els.performanceMemberTable.innerHTML = ranking.map(row => `
    <tr>
      <td><strong>${escapeHtml(row.member.name)}</strong><br><span>${escapeHtml(row.member.team)}</span></td>
      <td class="${scoreClass(row.rate)}">${fmtSigned(row.rate, 2)}</td>
      <td class="${scoreClass(row.curve)}">${fmtSigned(row.curve, 2)}</td>
      <td class="${scoreClass(row.credit)}">${fmtSigned(row.credit, 2)}</td>
      <td>${row.count}</td>
    </tr>
  `).join("") || emptyRow(5);

  els.performanceDetailTable.innerHTML = performanceDetailRows();

  drawWeeklyPerformanceChart();
  drawPerformanceChart();
}

function renderArchiveView() {
  const archiveSummaryList = document.querySelector("#archiveSummaryList");
  const archiveSelectedDetail = document.querySelector("#archiveSelectedDetail");
  const archiveWeekFilter = document.querySelector("#archiveWeekFilter");
  const archiveVersionFilter = document.querySelector("#archiveVersionFilter");
  const archiveMarketSnapshot = document.querySelector("#archiveMarketSnapshot");
  const archiveOpinionTable = document.querySelector("#archiveOpinionTable");
  const archiveResultTable = document.querySelector("#archiveResultTable");
  if (!archiveSummaryList || !archiveSelectedDetail || !archiveOpinionTable || !archiveResultTable) return;
  const archives = [...state.weeklyArchives]
    .sort((a, b) => b.week.localeCompare(a.week) || Number(b.version || 0) - Number(a.version || 0));
  const currentArchive = selectedArchive();
  if (currentArchive) selectedArchiveId = currentArchive.archiveId;
  const archiveWeeks = [...new Set(archives.map(archive => archive.week))];
  const versionsForWeek = currentArchive
    ? archives.filter(archive => archive.week === currentArchive.week)
    : [];

  if (archiveWeekFilter) {
    archiveWeekFilter.innerHTML = archiveWeeks
      .map(week => `<option value="${escapeHtml(week)}"${currentArchive?.week === week ? " selected" : ""}>${escapeHtml(weekDisplayLabel(week))}</option>`)
      .join("") || `<option value="">No archived week</option>`;
    archiveWeekFilter.disabled = !archiveWeeks.length;
  }

  if (archiveVersionFilter) {
    archiveVersionFilter.innerHTML = versionsForWeek
      .map(archive => {
        const label = `v${Number(archive.version) || 1} · ${archive.status || "-"} · ${archive.archivedAt ? archive.archivedAt.slice(0, 16).replace("T", " ") : "-"}`;
        return `<option value="${escapeHtml(archive.archiveId)}"${currentArchive?.archiveId === archive.archiveId ? " selected" : ""}>${escapeHtml(label)}</option>`;
      })
      .join("") || `<option value="">No version</option>`;
    archiveVersionFilter.disabled = !versionsForWeek.length;
  }

  archiveSummaryList.innerHTML = archives.map(archive => {
    const isSelected = currentArchive?.archiveId === archive.archiveId;
    const archivedAt = archive.archivedAt ? archive.archivedAt.slice(0, 16).replace("T", " ") : "-";
    return `
      <button class="archive-summary-item ${isSelected ? "selected" : ""}" type="button" data-archive-id="${escapeHtml(archive.archiveId)}">
        <div>
          <strong>${escapeHtml(archive.week)} v${Number(archive.version) || 1}</strong>
          <span>${escapeHtml(archive.status || "-")} · ${escapeHtml(archivedAt)}</span>
        </div>
        <span>${escapeHtml(archive.archivedBy || "-")}</span>
      </button>
    `;
  }).join("") || emptyBlock("No archived weeks yet.");

  if (!currentArchive) {
    archiveSelectedDetail.innerHTML = emptyBlock("No archive selected.");
    if (archiveMarketSnapshot) archiveMarketSnapshot.innerHTML = emptyBlock("No market snapshot.");
    archiveOpinionTable.innerHTML = emptyRow(7);
    archiveResultTable.innerHTML = emptyRow(8);
    return;
  }

  const archiveId = currentArchive.archiveId;
  const members = state.weeklyArchiveMembers.filter(row => row.archiveId === archiveId);
  const opinions = state.weeklyArchiveOpinionStrategies.filter(row => row.archiveId === archiveId);
  const markets = state.weeklyArchiveMarkets.filter(row => row.archiveId === archiveId);
  const results = state.weeklyArchiveResults.filter(row => row.archiveId === archiveId);
  const rawJson = state.weeklyArchiveRawJson.filter(row => row.archiveId === archiveId);

  archiveSelectedDetail.innerHTML = `
    <div class="preview-grid">
      <span><b>Week</b> ${escapeHtml(currentArchive.week)}</span>
      <span><b>Version</b> ${Number(currentArchive.version) || 1}</span>
      <span><b>Status</b> ${escapeHtml(currentArchive.status || "-")}</span>
      <span><b>Archived</b> ${escapeHtml(currentArchive.archivedAt ? currentArchive.archivedAt.slice(0, 16).replace("T", " ") : "-")}</span>
      <span><b>Archived by</b> ${escapeHtml(currentArchive.archivedBy || "-")}</span>
      <span><b>Members</b> ${members.length}</span>
      <span><b>Opinions</b> ${opinions.length}</span>
      <span><b>Markets</b> ${markets.length}</span>
      <span><b>Results</b> ${results.length}</span>
      <span><b>Raw JSON</b> ${rawJson.length}</span>
    </div>
    ${currentArchive.memo ? `<div class="archive-memo">${escapeHtml(currentArchive.memo)}</div>` : ""}
  `;

  if (archiveMarketSnapshot) {
    archiveMarketSnapshot.innerHTML = markets.length ? `
      <div class="archive-market-grid">
        ${markets
          .sort((a, b) => String(a.marketRole).localeCompare(String(b.marketRole)))
          .map(row => `
            <div class="archive-market-card">
              <strong>${row.marketRole === "base" ? "Base market" : "Evaluation market"}</strong>
              <span><b>Week</b> ${escapeHtml(row.marketWeek || "-")}</span>
              <span><b>Base date</b> ${escapeHtml(row.baseDate || row.asOf || "-")}</span>
              <span><b>KTB 3Y</b> ${fmt(row.ktb3y, 3)}%</span>
              <span><b>KTB 10Y</b> ${fmt(row.ktb10y, 3)}%</span>
              <span><b>3Y10Y</b> ${fmt(row.curveSpread, 1)}bp</span>
              <span><b>MSB 2Y</b> ${fmt(row.msb2y, 3)}%</span>
              <span><b>AA- 2Y</b> ${fmt(row.creditAA2y, 3)}%</span>
              <span><b>Credit</b> ${fmt(row.creditSpread, 1)}bp</span>
            </div>
          `).join("")}
      </div>
    ` : emptyBlock("No market snapshot.");
  }

  archiveOpinionTable.innerHTML = opinions
    .sort((a, b) => a.memberId.localeCompare(b.memberId) || a.strategyKey.localeCompare(b.strategyKey))
    .map(row => `
      <tr>
        <td>${escapeHtml(row.week)}</td>
        <td><strong>${escapeHtml(row.memberName || row.memberId)}</strong><br><span>${escapeHtml(row.memberId)}</span></td>
        <td>${escapeHtml(row.strategyKey)}</td>
        <td>${fmtSigned(row.duration, 2)}</td>
        <td>${fmt(row.rangeLow, 2)}~${fmt(row.rangeHigh, 2)}</td>
        <td>${fmt(row.confidence, 0)}</td>
        <td>${escapeHtml(row.rationaleText || "-")}</td>
      </tr>
    `).join("") || emptyRow(7);

  archiveResultTable.innerHTML = results
    .sort((a, b) => a.memberId.localeCompare(b.memberId) || a.strategyKey.localeCompare(b.strategyKey))
    .map(row => `
      <tr>
        <td>${escapeHtml(row.week)}</td>
        <td><strong>${escapeHtml(row.memberName || row.memberId)}</strong><br><span>${escapeHtml(row.memberId)}</span></td>
        <td>${escapeHtml(row.strategyKey)}</td>
        <td>${escapeHtml(row.evaluatedWeek || "-")}</td>
        <td>${fmtSigned(row.marketDelta, 2)}</td>
        <td class="${scoreClass(row.pnl)}">${fmtSigned(row.pnl, 2)}</td>
        <td class="${scoreClass(row.ytdPnl)}">${fmtSigned(row.ytdPnl, 2)}</td>
        <td>${row.rangeHit ? "Y" : "-"}</td>
      </tr>
    `).join("") || emptyRow(8);
}

function performanceDetailRows() {
  const rows = evaluatedRows().sort((a, b) => b.opinion.week.localeCompare(a.opinion.week));
  const weeks = [...new Set(rows.map(row => row.opinion.week))];

  return weeks.map(week => {
    const metric = selectedPerformanceMetric;
    const weekRows = rows.filter(row => row.opinion.week === week).sort((a, b) => b.pnls[metric] - a.pnls[metric]);
    const avgMetric = avg(weekRows.map(row => row.pnls[metric]));
    const detailRows = weekRows.map(row => `
      <tr>
        <td><strong>${escapeHtml(row.member.name)}</strong><br><span>${escapeHtml(row.member.team)}</span></td>
        <td>${escapeHtml(weekShortLabel(row.opinion.week))} -> ${escapeHtml(weekShortLabel(row.nextWeek) || "-")}</td>
        <td class="${scoreClass(row.pnls.rate)}">${fmtSigned(row.pnls.rate, 2)}</td>
        <td class="${scoreClass(row.pnls.curve)}">${fmtSigned(row.pnls.curve, 2)}</td>
        <td class="${scoreClass(row.pnls.credit)}">${fmtSigned(row.pnls.credit, 2)}</td>
      </tr>
    `).join("");

    return `
      <details class="week-disclosure">
        <summary>
          <span class="week-summary-main">
            <strong>${escapeHtml(weekDisplayLabel(week))}</strong>
            <span>${weekRows.length}명 평가</span>
          </span>
          <span class="week-summary-meta">
            <span>평가 구간 ${escapeHtml(weekShortLabel(week))} -> ${escapeHtml(weekShortLabel(weekRows[0]?.nextWeek) || "-")}</span>
            <span class="${scoreClass(avgMetric)}">평균 ${performanceMetricLabel(metric)} ${Number.isFinite(avgMetric) ? fmtSigned(avgMetric, 2) : "-"}</span>
          </span>
        </summary>
        <div class="table-wrap">
          <table class="week-detail-table">
            <thead>
              <tr>
                <th>본부원</th>
                <th>평가 구간</th>
                <th>금리</th>
                <th>커브</th>
                <th>크레딧</th>
              </tr>
            </thead>
            <tbody>${detailRows}</tbody>
          </table>
        </div>
      </details>
    `;
  }).join("") || `<div class="empty-state">평가된 손익 데이터가 없습니다.</div>`;
}

function latestDailyRateDate() {
  return [...state.dailyRates].sort((a, b) => a.date.localeCompare(b.date)).at(-1)?.date || "";
}

function renderRateUploadState() {
  const dates = [...state.dailyRates].map(row => row.date).sort();
  if (els.rateDateOptions) els.rateDateOptions.innerHTML = dates.map(date => `<option value="${date}"></option>`).join("");
  if (!dates.length) {
    els.rateUploadStatus.textContent = "업로드된 금리 시계열이 없습니다.";
    els.rateCoverage.innerHTML = emptyBlock("엑셀 파일을 업로드하면 기준일 후보가 표시됩니다.");
    return;
  }
  const file = state.uploadedRateFiles.at(-1);
  els.rateUploadStatus.textContent = `${file?.fileName || "업로드 파일"} · ${dates.length.toLocaleString()}개 일자 저장`;
  els.rateCoverage.innerHTML = `
    <div class="preview-grid">
      <span><b>기간</b> ${dates[0]} ~ ${dates.at(-1)}</span>
      <span><b>최근 업로드</b> ${file?.uploadedAt ? file.uploadedAt.slice(0, 19).replace("T", " ") : "-"}</span>
    </div>
  `;
}

function renderMarketDatePreview() {
  const week = els.marketWeek.value || selectedWeek;
  const date = els.marketBaseDate.value;
  const daily = dailyRateByDate(date);
  if (!date) {
    els.marketDatePreview.innerHTML = emptyBlock("주차에 연결할 기준일을 선택하세요.");
    return;
  }
  if (!daily) {
    els.marketDatePreview.innerHTML = emptyBlock("선택한 기준일의 업로드 금리 데이터가 없습니다.");
    return;
  }
  els.marketDatePreview.innerHTML = `
    <div class="preview-grid">
      <span><b>${escapeHtml(weekDisplayLabel(week))}</b> 기준일 ${escapeHtml(date)}</span>
      <span>국고3Y ${fmt(daily.ktb3y, 3)}%</span>
      <span>국고10Y ${fmt(daily.ktb10y, 3)}%</span>
      <span>3Y10Y ${fmt(daily.curveSpread, 1)}bp</span>
      <span>통안2Y ${fmt(daily.msb2y, 3)}%</span>
      <span>AA-2Y ${fmt(daily.creditAA2y, 3)}%</span>
      <span>크레딧 ${fmt(daily.creditSpread, 1)}bp</span>
    </div>
  `;
}

function renderMarketList() {
  const rows = sortedWeeks().map(marketByWeek).filter(Boolean).sort((a, b) => b.week.localeCompare(a.week)).slice(0, 6);
  els.latestMarketList.innerHTML = rows.map(row => `
    <div class="list-item">
      <strong>${escapeHtml(row.week)} · ${escapeHtml(row.asOf)}</strong>
      <span>국고3Y ${fmt(row.ktb3y, 3)}% · 3Y10Y ${fmt(row.curveSpread, 1)}bp · 크레딧 ${fmt(row.creditSpread, 1)}bp</span>
    </div>
  `).join("") || emptyBlock("시장 데이터가 없습니다.");
}

function renderMembers() {
  els.memberTable.innerHTML = [...state.members]
    .sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0))
    .map(member => `
      <tr>
        <td><strong>${escapeHtml(member.name)}</strong></td>
        <td>${escapeHtml(member.team)}</td>
        <td><span class="pill">${member.active === false ? "비활성" : "활성"}</span></td>
        <td>
          <button class="ghost-button" type="button" data-edit-member="${member.id}">수정</button>
          <button class="danger-button" type="button" data-delete-member="${member.id}">삭제</button>
        </td>
      </tr>
    `).join("") || emptyRow(4);
}

function scoreClass(value) {
  const number = finiteNumber(value);
  if (!Number.isFinite(number) || number === 0) return "score-flat";
  return number > 0 ? "score-pos" : "score-neg";
}

function emptyRow(colspan) {
  return `<tr><td colspan="${colspan}" class="empty-state">저장된 데이터가 없습니다.</td></tr>`;
}

function emptyBlock(text) {
  return `<div class="list-item"><strong>${text}</strong></div>`;
}

function svgEl(name, attrs = {}) {
  const el = document.createElementNS("http://www.w3.org/2000/svg", name);
  Object.entries(attrs).forEach(([key, value]) => el.setAttribute(key, value));
  return el;
}

function drawBase(svg) {
  svg.innerHTML = "";
  const width = svg.clientWidth || 900;
  const height = svg.clientHeight || 330;
  const pad = { top: 24, right: 58, bottom: 42, left: 58 };
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  return { width, height, pad, plotW: width - pad.left - pad.right, plotH: height - pad.top - pad.bottom };
}

function drawEmpty(svg, message) {
  const { width, height } = drawBase(svg);
  svg.append(svgEl("rect", { x: 1, y: 1, width: width - 2, height: height - 2, rx: 8, fill: "#f8faf8", stroke: "#d9e1db" }));
  const text = svgEl("text", { x: width / 2, y: height / 2, "text-anchor": "middle", class: "chart-label" });
  text.textContent = message;
  svg.append(text);
}

function drawAxes(svg, domain, range, labels) {
  const { width, height, pad, plotW, plotH } = domain.base;
  const sx = index => pad.left + (labels.length <= 1 ? plotW / 2 : (index / (labels.length - 1)) * plotW);
  const sy = value => pad.top + (1 - ((value - range.min) / (range.max - range.min))) * plotH;
  for (let i = 0; i <= 4; i += 1) {
    const y = pad.top + (plotH / 4) * i;
    const value = range.max - ((range.max - range.min) / 4) * i;
    svg.append(svgEl("line", { x1: pad.left, y1: y, x2: width - pad.right, y2: y, class: "grid-line" }));
    const t = svgEl("text", { x: 8, y: y + 4, class: "chart-label" });
    t.textContent = fmt(value, 2);
    svg.append(t);
  }
  labels.forEach((label, index) => {
    if (index === 0 || index === labels.length - 1 || labels.length <= 8) {
      const t = svgEl("text", { x: sx(index), y: height - 14, "text-anchor": "middle", class: "chart-label" });
      t.textContent = weekShortLabel(label).replace(/^\d{4}\./, "");
      svg.append(t);
    }
  });
  svg.append(svgEl("line", { x1: pad.left, y1: height - pad.bottom, x2: width - pad.right, y2: height - pad.bottom, class: "axis" }));
  svg.append(svgEl("line", { x1: pad.left, y1: pad.top, x2: pad.left, y2: height - pad.bottom, class: "axis" }));
  if (range.min < 0 && range.max > 0) {
    const zeroY = sy(0);
    svg.append(svgEl("line", { x1: pad.left, y1: zeroY, x2: width - pad.right, y2: zeroY, class: "zero-line" }));
  }
  return { sx, sy };
}

function drawDateAxes(svg, domain, range, startDate, endDate) {
  const { width, height, pad, plotW, plotH } = domain.base;
  const start = new Date(`${startDate}T00:00:00Z`).getTime();
  const end = new Date(`${endDate}T00:00:00Z`).getTime();
  const span = Math.max(end - start, 86400000);
  const sx = dateText => {
    const time = new Date(`${dateText}T00:00:00Z`).getTime();
    if (!Number.isFinite(time)) return pad.left;
    return pad.left + ((time - start) / span) * plotW;
  };
  const sy = value => pad.top + (1 - ((value - range.min) / (range.max - range.min))) * plotH;
  for (let i = 0; i <= 4; i += 1) {
    const y = pad.top + (plotH / 4) * i;
    const value = range.max - ((range.max - range.min) / 4) * i;
    svg.append(svgEl("line", { x1: pad.left, y1: y, x2: width - pad.right, y2: y, class: "grid-line" }));
    const t = svgEl("text", { x: 8, y: y + 4, class: "chart-label" });
    t.textContent = fmt(value, 2);
    svg.append(t);
  }
  for (let i = 0; i <= 4; i += 1) {
    const x = pad.left + (plotW / 4) * i;
    const date = new Date(start + span * (i / 4)).toISOString().slice(0, 10);
    const label = svgEl("text", { x, y: height - 14, "text-anchor": "middle", class: "chart-label" });
    label.textContent = shortDateLabel(date);
    svg.append(label);
  }
  return { sx, sy, startDate, endDate };
}

function boundedChartX(base, x) {
  return Math.min(
    base.width - base.pad.right - 10,
    Math.max(base.pad.left + 10, x)
  );
}

function offsetDateX(base, axes, dateText, index, count, step = 10) {
  const offset = (index - (count - 1) / 2) * step;
  return boundedChartX(base, axes.sx(dateText) + offset);
}

function drawRightAxis(svg, base, range) {
  const { width, height, pad, plotH } = base;
  const sy = value => pad.top + (1 - ((value - range.min) / (range.max - range.min))) * plotH;
  svg.append(svgEl("line", { x1: width - pad.right, y1: pad.top, x2: width - pad.right, y2: height - pad.bottom, class: "axis aux-axis" }));
  for (let i = 0; i <= 4; i += 1) {
    const y = pad.top + (plotH / 4) * i;
    const value = range.max - ((range.max - range.min) / 4) * i;
    const t = svgEl("text", { x: width - pad.right + 8, y: y + 4, class: "chart-label aux-label" });
    t.textContent = fmt(value, 2);
    svg.append(t);
  }
  return { sy };
}

function appendTitle(element, text) {
  element.dataset.tooltip = text;
  element.setAttribute("tabindex", "0");
  element.classList.add("has-tooltip");
  return element;
}

function setupChartTooltip(svg) {
  const wrap = svg.closest(".chart-wrap");
  const tooltip = wrap?.querySelector(".chart-tooltip");
  if (!wrap || !tooltip) return;

  const hide = () => {
    tooltip.classList.remove("visible");
  };

  const move = (event, target) => {
    const text = target?.dataset?.tooltip;
    if (!text) {
      hide();
      return;
    }

    tooltip.textContent = text;
    tooltip.classList.add("visible");

    const wrapRect = wrap.getBoundingClientRect();
    const tooltipRect = tooltip.getBoundingClientRect();
    const left = Math.min(
      wrapRect.width - tooltipRect.width - 10,
      Math.max(10, event.clientX - wrapRect.left + 12)
    );
    const top = Math.min(
      wrapRect.height - tooltipRect.height - 10,
      Math.max(10, event.clientY - wrapRect.top + 12)
    );

    tooltip.style.left = `${left}px`;
    tooltip.style.top = `${top}px`;
  };

  svg.onmousemove = event => {
    const target = event.target?.closest?.("[data-tooltip]");
    move(event, target);
  };
  svg.onmouseleave = hide;
  svg.onfocusin = event => {
    const target = event.target?.closest?.("[data-tooltip]");
    if (!target) return;
    const targetRect = target.getBoundingClientRect();
    move(
      {
        clientX: targetRect.left + targetRect.width / 2,
        clientY: targetRect.top + targetRect.height / 2
      },
      target
    );
  };
  svg.onfocusout = hide;
}

function chartRange(values, fallbackSpan = 1, includeZero = true) {
  const valid = values.filter(Number.isFinite);
  if (!valid.length) return { min: -fallbackSpan, max: fallbackSpan };
  const min = includeZero ? Math.min(...valid, 0) : Math.min(...valid);
  const max = includeZero ? Math.max(...valid, 0) : Math.max(...valid);
  const span = Math.max(max - min, fallbackSpan);
  return { min: min - span * .2, max: max + span * .2 };
}

function strategyChartDateWindow() {
  const selectedBaseDate = marketDateForWeek(selectedWeek);
  const fallbackBaseDate = sortedWeeks()
    .map(marketDateForWeek)
    .filter(Boolean)
    .sort()
    .at(-1) || latestDailyRateDate();
  const endDate = selectedBaseDate || fallbackBaseDate;
  const startDate = endDate ? addDaysIso(endDate, -30) : "";
  const axisEndDate = startDate && endDate
    ? addDaysIso(startDate, Math.round(((new Date(`${endDate}T00:00:00Z`) - new Date(`${startDate}T00:00:00Z`)) / 86400000) / 0.75))
    : "";
  return { startDate, endDate, axisEndDate };
}

function strategyChartDailyRows(strategy, startDate, endDate) {
  const rowsByDate = new Map([...state.dailyRates]
    .filter(row => row.date >= startDate && row.date <= endDate && Number.isFinite(finiteNumber(row[strategy.marketKey])))
    .map(row => [row.date, { ...row, source: "daily" }]));
  const selectedMarketDate = marketDateForWeek(selectedWeek);
  const selectedMarket = marketByWeek(selectedWeek);
  const selectedMarketValue = finiteNumber(selectedMarket?.[strategy.marketKey]);
  if (
    selectedMarketDate >= startDate &&
    selectedMarketDate <= endDate &&
    Number.isFinite(selectedMarketValue) &&
    !rowsByDate.has(selectedMarketDate)
  ) {
    rowsByDate.set(selectedMarketDate, {
      ...(rowsByDate.get(selectedMarketDate) || {}),
      date: selectedMarketDate,
      [strategy.marketKey]: selectedMarketValue,
      source: "meeting"
    });
  }
  return [...rowsByDate.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function drawStrategyChart(svg, strategy, mode = "duration") {
  const { startDate, endDate, axisEndDate } = strategyChartDateWindow();
  const dailyRows = startDate && endDate ? strategyChartDailyRows(strategy, startDate, endDate) : [];
  const rows = state.weeklyOpinions
    .map(row => ({ ...row, chartDate: marketDateForWeek(row.week) }))
    .filter(row =>
      row.chartDate >= startDate &&
      row.chartDate <= endDate &&
      (mode === "range"
        ? Number.isFinite(finiteNumber(row[strategy.rangeLowKey])) && Number.isFinite(finiteNumber(row[strategy.rangeHighKey]))
        : Number.isFinite(finiteNumber(row[strategy.durationKey])))
    );
  const marketValues = dailyRows.map(row => Number(row[strategy.marketKey])).filter(Number.isFinite);
  if (!startDate || !endDate || (!rows.length && !marketValues.length)) {
    drawEmpty(svg, "전략 의견과 시장 데이터를 입력하면 차트가 표시됩니다.");
    return;
  }

  const base = drawBase(svg);
  const primaryValues = mode === "range"
    ? [
        ...rows.flatMap(row => [finiteNumber(row[strategy.rangeLowKey]), finiteNumber(row[strategy.rangeHighKey])]),
        ...marketValues
      ]
    : rows.map(row => finiteNumber(row[strategy.durationKey]));
  const primaryRange = chartRange(
    primaryValues,
    mode === "range" && strategy.key !== "rate" ? 10 : .1,
    mode !== "range"
  );
  const axes = drawDateAxes(svg, { base }, primaryRange, startDate, axisEndDate || endDate);
  let marketAxis = null;
  if (marketValues.length && mode !== "range") {
    const marketMin = Math.min(...marketValues);
    const marketMax = Math.max(...marketValues);
    const marketSpan = Math.max(marketMax - marketMin, strategy.key === "rate" ? .1 : 10);
    marketAxis = drawRightAxis(svg, base, { min: marketMin - marketSpan * .18, max: marketMax + marketSpan * .18 });
  }
  const colors = ["#0f766e", "#356bb3", "#b9781f", "#b54747", "#6f5aa7", "#4d7d3a"];
  const memberColorById = new Map(orderedMembers().map((member, index) => [member.id, colors[index % colors.length]]));

  if (marketValues.length) {
    const marketPoints = dailyRows
      .map(row => {
        const value = Number(row[strategy.marketKey]);
        const y = mode === "range" ? axes.sy(value) : marketAxis.sy(value);
        return Number.isFinite(value) ? { date: row.date, value, x: axes.sx(row.date), y } : null;
      })
      .filter(Boolean);
    if (marketPoints.length > 1) {
      const path = svgEl("path", {
        d: marketPoints.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" "),
        class: mode === "range" ? "line-path" : "line-path aux-market-line"
      });
      appendTitle(path, `${strategy.title} 시장 데이터`);
      svg.append(path);
    }
    marketPoints.forEach(point => {
      const tooltip = `${point.date} 시장 ${fmt(point.value, strategy.key === "rate" ? 3 : 1)}${strategy.unit}`;
      const hit = svgEl("circle", {
        cx: point.x,
        cy: point.y,
        r: mode === "range" ? 9 : 7,
        class: "market-hit-dot"
      });
      appendTitle(hit, tooltip);
      svg.append(hit);
      if (mode !== "range" && (shortDateLabel(point.date).endsWith("01") || point.date === endDate)) {
        const dot = svgEl("circle", { cx: point.x, cy: point.y, r: 3, class: "dot aux-dot" });
        appendTitle(dot, tooltip);
        svg.append(dot);
      }
    });
  }

  if (mode === "range") {
    const chartDates = [...new Set(rows.map(row => row.chartDate))];
    chartDates.forEach(chartDate => {
      const weekRows = rows.filter(row => row.chartDate === chartDate);
      weekRows.forEach((opinion, rowIndex) => {
        const member = memberById(opinion.memberId);
        const color = memberColorById.get(opinion.memberId) || colors[rowIndex % colors.length];
        const low = finiteNumber(opinion[strategy.rangeLowKey]);
        const high = finiteNumber(opinion[strategy.rangeHighKey]);
        if (!Number.isFinite(low) || !Number.isFinite(high)) return;
        if (!Number.isFinite(low) || !Number.isFinite(high)) return;
        const x = offsetDateX(base, axes, chartDate, rowIndex, weekRows.length, 12);
        const line = svgEl("line", { x1: x, y1: axes.sy(low), x2: x, y2: axes.sy(high), stroke: color, class: "range-line" });
        appendTitle(line, `${weekDisplayLabel(opinion.week)} ${member.name} 레인지 ${fmt(low, strategy.key === "rate" ? 3 : 1)}~${fmt(high, strategy.key === "rate" ? 3 : 1)}${strategy.rangeUnit}`);
        svg.append(line);
        svg.append(appendTitle(svgEl("line", { x1: x - 4, y1: axes.sy(low), x2: x + 4, y2: axes.sy(low), stroke: color, class: "range-line" }), `${member.name} 하단 ${fmt(low, strategy.key === "rate" ? 3 : 1)}${strategy.rangeUnit}`));
        svg.append(appendTitle(svgEl("line", { x1: x - 4, y1: axes.sy(high), x2: x + 4, y2: axes.sy(high), stroke: color, class: "range-line" }), `${member.name} 상단 ${fmt(high, strategy.key === "rate" ? 3 : 1)}${strategy.rangeUnit}`));
      });
    });
    const title = svgEl("text", { x: base.width - base.pad.right, y: 18, "text-anchor": "end", class: "chart-label" });
    title.textContent = `${shortDateLabel(startDate)}~${shortDateLabel(endDate)} 일별 시장`;
    svg.append(title);
    setupChartTooltip(svg);
    return;
  }

  orderedMembers().forEach((member, memberIndex) => {
    const memberRows = rows
      .filter(row => row.memberId === member.id)
      .sort((a, b) => a.chartDate.localeCompare(b.chartDate));
    const points = memberRows.map(opinion => {
      const value = opinion ? finiteNumber(opinion[strategy.durationKey]) : NaN;
      const dateRows = rows.filter(row => row.chartDate === opinion.chartDate);
      const dateIndex = dateRows.findIndex(row => row.memberId === opinion.memberId);
      const x = offsetDateX(base, axes, opinion.chartDate, Math.max(dateIndex, 0), dateRows.length, 10);
      return Number.isFinite(value) ? { week: opinion.week, date: opinion.chartDate, value, x, y: axes.sy(value) } : null;
    })
      .filter(Boolean);
    if (points.length > 1) {
      const path = svgEl("path", {
        d: points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" "),
        fill: "none",
        stroke: colors[memberIndex % colors.length],
        "stroke-width": 2,
        opacity: .9
      });
      appendTitle(path, member.name);
      svg.append(path);
    }
    points.forEach(point => {
      const dot = svgEl("circle", { cx: point.x, cy: point.y, r: 4.5, fill: colors[memberIndex % colors.length], stroke: "#fff", "stroke-width": 1.5 });
      appendTitle(dot, `${weekDisplayLabel(point.week)} ${member.name} ${fmtSigned(point.value, 2)}`);
      svg.append(dot);
    });
  });

  const avgPoints = [...new Set(rows.map(row => row.chartDate))]
    .sort()
    .map(chartDate => {
      const dayRows = rows.filter(row => row.chartDate === chartDate);
      const value = avg(dayRows.map(row => row[strategy.durationKey]));
      return Number.isFinite(value) ? { date: chartDate, value, x: axes.sx(chartDate), y: axes.sy(value) } : null;
    })
    .filter(Boolean);
  if (avgPoints.length > 1) {
    const avgPath = svgEl("path", { d: avgPoints.map((p, i) => `${i ? "L" : "M"} ${p.x} ${p.y}`).join(" "), class: "line-path med" });
    appendTitle(avgPath, `${strategy.title} 평균 듀레이션`);
    svg.append(avgPath);
  }
  const title = svgEl("text", { x: base.width - base.pad.right, y: 18, "text-anchor": "end", class: "chart-label" });
  title.textContent = `${shortDateLabel(startDate)}~${shortDateLabel(endDate)} 일별 시장`;
  svg.append(title);
  setupChartTooltip(svg);
}

function drawDurationChart(svg, strategy) {
  const rows = state.weeklyOpinions.filter(row => Number.isFinite(finiteNumber(row[strategy.durationKey])));
  const weeks = sortedWeeks();
  if (!rows.length || weeks.length < 2) {
    drawEmpty(svg, "전략 의견을 입력하면 시계열이 표시됩니다.");
    return;
  }
  const base = drawBase(svg);
  const values = rows.map(row => finiteNumber(row[strategy.durationKey]));
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 0);
  const span = Math.max(max - min, .1);
  const axes = drawAxes(svg, { base }, { min: min - span * .25, max: max + span * .25 }, weeks);
  const colors = ["#0f766e", "#356bb3", "#b9781f", "#b54747", "#6f5aa7", "#4d7d3a"];

  orderedMembers().forEach((member, memberIndex) => {
    const points = weeks.map((week, index) => {
      const opinion = rows.find(row => row.week === week && row.memberId === member.id);
      const value = opinion ? finiteNumber(opinion[strategy.durationKey]) : NaN;
      return Number.isFinite(value) ? { x: axes.sx(index), y: axes.sy(value) } : null;
    }).filter(Boolean);
    if (points.length > 1) {
      const path = points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ");
      svg.append(svgEl("path", { d: path, fill: "none", stroke: colors[memberIndex % colors.length], "stroke-width": 2, opacity: .9 }));
    }
    points.forEach(point => svg.append(svgEl("circle", { cx: point.x, cy: point.y, r: 4.5, fill: colors[memberIndex % colors.length], stroke: "#fff", "stroke-width": 1.5 })));
  });

  const avgPoints = weeks.map((week, index) => {
    const value = avg(opinionsForWeek(week).map(row => row[strategy.durationKey]));
    return Number.isFinite(value) ? { x: axes.sx(index), y: axes.sy(value) } : null;
  }).filter(Boolean);
  if (avgPoints.length > 1) {
    svg.append(svgEl("path", { d: avgPoints.map((p, i) => `${i ? "L" : "M"} ${p.x} ${p.y}`).join(" "), class: "line-path med" }));
  }
}

function drawRangeChart(svg, strategy) {
  const markets = [...state.marketData].sort((a, b) => a.week.localeCompare(b.week));
  if (markets.length < 2) {
    drawEmpty(svg, "시장 데이터를 2개 주차 이상 입력하세요.");
    return;
  }
  const weeks = markets.map(row => row.week);
  const ranges = state.weeklyOpinions.flatMap(row => [finiteNumber(row[strategy.rangeLowKey]), finiteNumber(row[strategy.rangeHighKey])]).filter(Number.isFinite);
  const marketValues = markets.map(row => Number(row[strategy.marketKey])).filter(Number.isFinite);
  const values = [...ranges, ...marketValues];
  const min = Math.min(...values);
  const max = Math.max(...values);
  const span = Math.max(max - min, strategy.key === "rate" ? .1 : 10);
  const base = drawBase(svg);
  const axes = drawAxes(svg, { base }, { min: min - span * .18, max: max + span * .18 }, weeks);

  const linePoints = markets.map((row, index) => ({ x: axes.sx(index), y: axes.sy(Number(row[strategy.marketKey])) }));
  svg.append(svgEl("path", { d: linePoints.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" "), class: "line-path" }));
  linePoints.forEach(point => svg.append(svgEl("circle", { cx: point.x, cy: point.y, r: 4.5, class: "dot" })));

  weeks.forEach((week, weekIndex) => {
    const rows = opinionsForWeek(week);
    rows.forEach((opinion, rowIndex) => {
      const low = finiteNumber(opinion[strategy.rangeLowKey]);
      const high = finiteNumber(opinion[strategy.rangeHighKey]);
      if (!Number.isFinite(low) || !Number.isFinite(high)) return;
      if (!Number.isFinite(low) || !Number.isFinite(high)) return;
      const offset = (rowIndex - (rows.length - 1) / 2) * 5;
      const x = axes.sx(weekIndex) + offset;
      svg.append(svgEl("line", { x1: x, y1: axes.sy(low), x2: x, y2: axes.sy(high), stroke: "#356bb3", class: "range-line" }));
      svg.append(svgEl("line", { x1: x - 4, y1: axes.sy(low), x2: x + 4, y2: axes.sy(low), stroke: "#356bb3", class: "range-line" }));
      svg.append(svgEl("line", { x1: x - 4, y1: axes.sy(high), x2: x + 4, y2: axes.sy(high), stroke: "#356bb3", class: "range-line" }));
    });
  });
}

function drawWeeklyPerformanceChart() {
  const svg = els.weeklyPerformanceChart;
  const { week, rows } = weeklyPerformanceRows();
  if (!rows.length) {
    drawEmpty(svg, "평가 확정 후 주간 성과 차트가 표시됩니다.");
    return;
  }

  const base = drawBase(svg);
  const labels = rows.map(row => row.member.name);
  const metrics = [
    { key: "rate", label: "금리", color: "#356bb3" },
    { key: "curve", label: "커브", color: "#b9781f" },
    { key: "credit", label: "크레딧", color: "#b54747" }
  ];
  const metricValue = (row, key) => row.pnls[key];
  const values = rows.flatMap(row => metrics.map(metric => metricValue(row, metric.key)));
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 0);
  const span = Math.max(max - min, 1);
  const axes = drawAxes(svg, { base }, { min: min - span * .18, max: max + span * .18 }, []);
  const { width, height, pad, plotW } = base;
  const zeroY = axes.sy(0);
  const slot = plotW / Math.max(labels.length, 1);
  const barCenterX = index => pad.left + slot * index + slot / 2;
  const groupWidth = Math.min(86, Math.max(44, slot * .64));
  const gap = 3;
  const barWidth = Math.max(6, (groupWidth - gap * (metrics.length - 1)) / metrics.length);

  rows.forEach((row, index) => {
    const centerX = barCenterX(index);
    const groupStart = centerX - groupWidth / 2;
    metrics.forEach((metric, metricIndex) => {
      const value = metricValue(row, metric.key);
      const x = groupStart + metricIndex * (barWidth + gap);
      const y = axes.sy(Math.max(value, 0));
      const barHeight = Math.max(2, Math.abs(axes.sy(value) - zeroY));
      const rect = svgEl("rect", {
        x,
        y: value >= 0 ? y : zeroY,
        width: barWidth,
        height: barHeight,
        rx: 3,
        fill: metric.color,
        opacity: value >= 0 ? .88 : .62
      });
      appendTitle(rect, `${row.member.name} ${metric.label} ${fmtSigned(value, 2)}`);
      svg.append(rect);
    });

    const nameLabel = svgEl("text", {
      x: centerX,
      y: height - 14,
      "text-anchor": "middle",
      class: "chart-label"
    });
    nameLabel.textContent = row.member.name;
    svg.append(nameLabel);
  });

  const title = svgEl("text", { x: width - pad.right, y: 18, "text-anchor": "end", class: "chart-label" });
  title.textContent = `${weekDisplayLabel(week)} 주간 성과`;
  svg.append(title);
  metrics.forEach((metric, index) => {
    const legendX = pad.left + index * 72;
    svg.append(svgEl("rect", { x: legendX, y: 12, width: 10, height: 10, rx: 2, fill: metric.color }));
    const legendText = svgEl("text", { x: legendX + 15, y: 21, class: "chart-label" });
    legendText.textContent = metric.label;
    svg.append(legendText);
  });
  svg.append(svgEl("line", { x1: pad.left, y1: zeroY, x2: width - pad.right, y2: zeroY, class: "zero-line" }));
  setupChartTooltip(svg);
}

function drawPerformanceChart() {
  const svg = els.performanceChart;
  const rows = evaluatedRows().sort((a, b) => a.opinion.week.localeCompare(b.opinion.week));
  if (!rows.length) {
    drawEmpty(svg, "평가 확정 후 YTD 추이 차트가 표시됩니다.");
    return;
  }
  const weeks = [...new Set(rows.map(row => row.opinion.week))];
  const metric = selectedPerformanceMetric;
  const members = rankingRows(metric).filter(row => row.count > 0).map(row => row.member);
  const cumulative = new Map(members.map(member => [member.id, []]));
  const running = new Map(members.map(member => [member.id, 0]));

  weeks.forEach(week => {
    const weekRows = rows.filter(row => row.opinion.week === week);
    members.forEach(member => {
      const memberRows = weekRows.filter(row => row.member.id === member.id);
      const weeklyValue = memberRows.reduce((sum, row) => sum + row.pnls[metric], 0);
      const nextValue = round((running.get(member.id) || 0) + weeklyValue, 2);
      running.set(member.id, nextValue);
      cumulative.get(member.id).push({ week, value: nextValue });
    });
  });
  const values = [...cumulative.values()].flat().map(point => point.value);
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 0);
  const span = Math.max(max - min, 1);
  const base = drawBase(svg);
  const axes = drawAxes(svg, { base }, { min: min - span * .2, max: max + span * .2 }, weeks);
  const colors = ["#0f766e", "#356bb3", "#b9781f", "#b54747", "#6f5aa7", "#4d7d3a"];
  members.forEach((member, memberIndex) => {
    const points = cumulative.get(member.id).map((point, index) => ({ x: axes.sx(index), y: axes.sy(point.value) }));
    if (points.length > 1) {
      svg.append(svgEl("path", {
        d: points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" "),
        fill: "none",
        stroke: colors[memberIndex % colors.length],
        "stroke-width": 2.6,
        opacity: .92
      }));
    }
    points.forEach(point => svg.append(svgEl("circle", {
      cx: point.x,
      cy: point.y,
      r: 3.8,
      fill: colors[memberIndex % colors.length],
      stroke: "#fff",
      "stroke-width": 1.4
    })));
  });

  const title = svgEl("text", { x: base.width - base.pad.right, y: 18, "text-anchor": "end", class: "chart-label" });
  title.textContent = `${performanceMetricLabel(metric)} YTD`;
  svg.append(title);
}

async function parseRatesWorkbook(file) {
  const entries = await unzipWorkbook(await file.arrayBuffer());
  const strings = parseSharedStrings(await readZipText(entries, "xl/sharedStrings.xml", false));
  const workbook = parseXml(await readZipText(entries, "xl/workbook.xml"));
  const rels = parseWorkbookRels(await readZipText(entries, "xl/_rels/workbook.xml.rels"));
  const sheets = [...workbook.querySelectorAll("sheet")].map(sheet => ({
    name: sheet.getAttribute("name"),
    path: `xl/${rels.get(sheet.getAttribute("r:id")) || ""}`.replace("xl//", "xl/")
  }));
  const sheet = sheets.find(item => item.name === "Raw") || sheets[0];
  if (!sheet?.path) throw new Error("sheet-not-found");
  const sheetXml = parseXml(await readZipText(entries, sheet.path));
  const rows = sheetRows(sheetXml, strings);
  const header = detectRateHeaderRows(rows);
  const columns = detectRateColumns(rows.get(header.groupRow), rows.get(header.labelRow));
  const rates = [];
  [...rows.keys()].filter(row => row > header.labelRow).sort((a, b) => a - b).forEach(rowIndex => {
    const row = rows.get(rowIndex);
    const date = normalizeDateText(row?.get(columns.date));
    const ktb3y = Number(row?.get(columns.ktb3y));
    const ktb10y = Number(row?.get(columns.ktb10y));
    const msb2y = Number(row?.get(columns.msb2y));
    const creditAA2y = Number(row?.get(columns.creditAA2y));
    if (!date || [ktb3y, ktb10y, msb2y, creditAA2y].some(value => !Number.isFinite(value) || value === 0)) return;
    rates.push(normalizeDailyRate({ date, ktb3y, ktb10y, msb2y, creditAA2y }));
  });
  const unique = new Map(rates.filter(Boolean).map(row => [row.date, row]));
  const sorted = [...unique.values()].sort((a, b) => a.date.localeCompare(b.date));
  if (!sorted.length) throw new Error("no-rate-rows");
  return sorted;
}

async function loadBundledRatesIfAvailable({ showStatus = true } = {}) {
  try {
    const response = await fetch("rates_raw.xlsx", { cache: "no-store" });
    if (!response.ok) return false;
    const rates = await parseRatesWorkbook(await response.blob());
    const bundledMaxDate = rates.at(-1)?.date || "";
    state.dailyRates = rates;
    state.uploadedRateFiles = [
      ...state.uploadedRateFiles.filter(file => file.fileName !== "rates_raw.xlsx"),
      {
        id: "bundled-rates-raw",
        fileName: "rates_raw.xlsx",
        uploadedAt: new Date().toISOString(),
        rowCount: rates.length,
        minDate: rates[0]?.date || "",
        maxDate: bundledMaxDate
      }
    ];
    syncMappedMarketsFromDailyRates();
    saveState();
    render();
    const latest = rates.find(row => row.date === bundledMaxDate);
    if (showStatus && latest) {
      setMicrosoftStatus(
        `로컬 raw 금리 반영: ${bundledMaxDate} 국고3Y ${fmt(latest.ktb3y, 3)}%, 커브 ${fmt(latest.curveSpread, 1)}bp, 크레딧 ${fmt(latest.creditSpread, 1)}bp`,
        "connected"
      );
    }
    return true;
  } catch (error) {
    console.warn("bundled raw rates load skipped", error);
    return false;
  }
}

function detectRateHeaderRows(rows) {
  const rowNumbers = [...rows.keys()].sort((a, b) => a - b);
  const groupRow = rowNumbers.find(row => [...rows.get(row).values()].some(value => String(value).includes("국고채권")));
  if (!groupRow) throw new Error("rate-group-row-not-found");
  const labelRow = rowNumbers.find(row => row > groupRow && [...rows.get(row).values()].some(value => String(value).includes("일자")));
  if (!labelRow) throw new Error("rate-label-row-not-found");
  return { groupRow, labelRow };
}

async function unzipWorkbook(buffer) {
  const bytes = new Uint8Array(buffer);
  const entries = new Map();
  const eocd = findEndOfCentralDirectory(bytes);
  const view = new DataView(bytes.buffer);
  const entryCount = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  for (let i = 0; i < entryCount; i += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("bad-zip-directory");
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = new TextDecoder().decode(bytes.slice(offset + 46, offset + 46 + fileNameLength));
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    entries.set(name, { method, data: bytes.slice(dataStart, dataStart + compressedSize) });
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function findEndOfCentralDirectory(bytes) {
  for (let i = bytes.length - 22; i >= 0; i -= 1) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) return i;
  }
  throw new Error("zip-end-not-found");
}

async function readZipText(entries, path, required = true) {
  const entry = entries.get(path);
  if (!entry) {
    if (required) throw new Error(`missing ${path}`);
    return "";
  }
  let data = entry.data;
  if (entry.method === 8) data = await inflateRaw(data);
  if (entry.method !== 0 && entry.method !== 8) throw new Error("unsupported-zip-method");
  return new TextDecoder("utf-8").decode(data);
}

async function inflateRaw(data) {
  if (!("DecompressionStream" in window)) throw new Error("decompression-not-supported");
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("xml-parse-error");
  return doc;
}

function parseSharedStrings(text) {
  if (!text) return [];
  return [...parseXml(text).querySelectorAll("si")].map(item => [...item.querySelectorAll("t")].map(node => node.textContent || "").join(""));
}

function parseWorkbookRels(text) {
  const map = new Map();
  [...parseXml(text).querySelectorAll("Relationship")].forEach(rel => {
    map.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
  });
  return map;
}

function sheetRows(doc, sharedStrings) {
  const rows = new Map();
  [...doc.querySelectorAll("sheetData row")].forEach(rowNode => {
    const rowIndex = Number(rowNode.getAttribute("r"));
    const row = new Map();
    [...rowNode.querySelectorAll("c")].forEach(cell => {
      const col = columnIndex(cell.getAttribute("r"));
      row.set(col, cellValue(cell, sharedStrings));
    });
    rows.set(rowIndex, row);
  });
  return rows;
}

function cellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") return cell.querySelector("is t")?.textContent || "";
  const raw = cell.querySelector("v")?.textContent || "";
  if (type === "s") return sharedStrings[Number(raw)] || "";
  return raw;
}

function columnIndex(ref = "") {
  const letters = ref.match(/[A-Z]+/)?.[0] || "";
  return [...letters].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function detectRateColumns(groupRow, labelRow) {
  const headerText = value => String(value || "").replace(/\s+/g, "");
  const groups = [...groupRow.entries()]
    .sort((a, b) => a[0] - b[0])
    .filter(([, value]) => value)
    .map(([col, value]) => [col, headerText(value)]);
  const groupFor = col => groups.filter(([groupCol]) => groupCol <= col).at(-1)?.[1] || "";
  const findCol = (groupText, labelText) => {
    const targetGroup = headerText(groupText);
    const targetLabel = headerText(labelText);
    const found = [...labelRow.entries()].find(([col, label]) => {
      const group = groupFor(col);
      const labelValue = headerText(label);
      return group.includes(targetGroup) && labelValue.startsWith(targetLabel);
    });
    if (!found) throw new Error(`missing ${groupText} ${labelText}`);
    return found[0];
  };
  return {
    date: findCol("시가평가 4사평균 국고채권", "일자"),
    ktb3y: findCol("시가평가 4사평균 국고채권", "3년이하"),
    ktb10y: findCol("시가평가 4사평균 국고채권", "10년이하"),
    msb2y: findCol("시가평가 4사평균 통안증권", "2년이하"),
    creditAA2y: findCol("시가평가 4사평균 기타금융채AA-", "2년이하")
  };
}

function syncMappedMarketsFromDailyRates() {
  state.marketWeekMappings.forEach(mapping => {
    const daily = dailyRateByDate(mapping.baseDate);
    if (!daily) return;
    const market = dailyRateToMarket(mapping.week, daily, mapping.memo);
    const existing = state.marketData.findIndex(row => row.week === mapping.week);
    if (existing >= 0) state.marketData[existing] = market;
    else state.marketData.push(market);
  });
}

function upsertOpinion(item) {
  const existing = state.weeklyOpinions.findIndex(row => row.week === item.week && row.memberId === item.memberId);
  const now = new Date().toISOString();
  if (existing >= 0) {
    state.weeklyOpinions[existing] = { ...state.weeklyOpinions[existing], ...item, updatedAt: now };
    return state.weeklyOpinions[existing];
  } else {
    const opinion = { id: id("opinion"), ...item, createdAt: now, updatedAt: now };
    state.weeklyOpinions.push(opinion);
    return opinion;
  }
}

function upsertWeeklySummary(item) {
  const existing = state.weeklySummaries.findIndex(row => row.week === item.week);
  const now = new Date().toISOString();
  if (existing >= 0) {
    state.weeklySummaries[existing] = {
      ...state.weeklySummaries[existing],
      ...item,
      memberId: SUMMARY_MEMBER_ID,
      status: "completed",
      completedAt: now,
      updatedAt: now
    };
    return state.weeklySummaries[existing];
  }
  const summary = {
    id: id("weekly-summary"),
    ...item,
    memberId: SUMMARY_MEMBER_ID,
    status: "completed",
    completedAt: now,
    createdAt: now,
    updatedAt: now
  };
  state.weeklySummaries.push(summary);
  return summary;
}

function upsertMarket(item) {
  const daily = dailyRateByDate(item.baseDate);
  if (!daily) throw new Error("missing-daily-rate");
  const now = new Date().toISOString();
  const market = dailyRateToMarket(item.week, daily, item.memo);
  const existing = state.marketData.findIndex(row => row.week === market.week);
  if (existing >= 0) state.marketData[existing] = market;
  else state.marketData.push(market);

  const previousMapping = state.marketWeekMappings.find(row => row.week === item.week);
  const mapping = {
    week: item.week,
    baseDate: item.baseDate,
    memo: item.memo || "",
    createdAt: previousMapping?.createdAt || now,
    updatedAt: now
  };
  const mappingIndex = state.marketWeekMappings.findIndex(row => row.week === item.week);
  if (mappingIndex >= 0) state.marketWeekMappings[mappingIndex] = mapping;
  else state.marketWeekMappings.push(mapping);
  state.marketMappingHistory.push({
    id: id("market-map"),
    week: item.week,
    previousBaseDate: previousMapping?.baseDate || "",
    baseDate: item.baseDate,
    memo: item.memo || "",
    changedAt: now
  });
  return {
    market,
    mapping,
    history: state.marketMappingHistory.at(-1)
  };
}

function upsertMember(item) {
  if (item.memberId) {
    const existing = state.members.findIndex(member => member.id === item.memberId);
    if (existing >= 0) {
      state.members[existing] = { ...state.members[existing], ...item, id: item.memberId };
      return state.members[existing];
    }
  } else {
    const member = { id: id("member"), ...item };
    state.members.push(member);
    return member;
  }
  const member = { id: item.memberId || id("member"), ...item };
  state.members.push(member);
  return member;
}

function download(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function buildCsv() {
  const opinionHeader = ["type", "week", "memberId", "rateDuration", "curveDuration", "creditDuration", "ktb3yRangeLow", "ktb3yRangeHigh", "curveSpreadRangeLow", "curveSpreadRangeHigh", "creditSpreadRangeLow", "creditSpreadRangeHigh", "confidence", "rateConfidence", "curveConfidence", "creditConfidence", "rationaleTags", "rationaleText", "rateRationaleText", "curveRationaleText", "creditRationaleText"];
  const weeklySummaryHeader = ["type", "week", "rateDuration", "curveDuration", "creditDuration", "ktb3yRangeLow", "ktb3yRangeHigh", "curveSpreadRangeLow", "curveSpreadRangeHigh", "creditSpreadRangeLow", "creditSpreadRangeHigh", "confidence", "rateConfidence", "curveConfidence", "creditConfidence", "overallSummaryText", "rateSummaryText", "curveSummaryText", "creditSummaryText", "status", "draftSource", "completedAt", "createdAt", "updatedAt"];
  const marketHeader = ["type", "week", "asOf", "ktb3y", "ktb10y", "curveSpread", "msb2y", "creditAA2y", "creditSpread", "memo"];
  const dailyRateHeader = ["type", "date", "ktb3y", "ktb10y", "curveSpread", "msb2y", "creditAA2y", "creditSpread"];
  const mappingHeader = ["type", "week", "baseDate", "memo", "createdAt", "updatedAt"];
  const mappingHistoryHeader = ["type", "week", "previousBaseDate", "baseDate", "memo", "changedAt"];
  const archiveHeader = ["type", "archiveId", "week", "version", "status", "archivedAt", "archivedBy", "memo", "calculationVersion", "sourceHash"];
  const archiveOpinionHeader = ["type", "archiveId", "week", "version", "memberId", "memberName", "team", "strategyKey", "duration", "rangeLow", "rangeHigh", "confidence", "rationaleText"];
  const archiveResultHeader = ["type", "archiveId", "week", "version", "memberId", "memberName", "strategyKey", "evaluatedWeek", "marketDelta", "pnl", "ytdPnl", "rangeHit"];
  const memberHeader = ["type", "id", "name", "team", "active", "sortOrder"];
  const rows = [
    memberHeader,
    ...state.members.map(row => ["member", row.id, row.name, row.team, row.active !== false, row.sortOrder]),
    [],
    opinionHeader,
    ...state.weeklyOpinions.map(row => ["opinion", row.week, row.memberId, row.rateDuration, row.curveDuration, row.creditDuration, row.ktb3yRangeLow, row.ktb3yRangeHigh, row.curveSpreadRangeLow, row.curveSpreadRangeHigh, row.creditSpreadRangeLow, row.creditSpreadRangeHigh, row.confidence, row.rateConfidence, row.curveConfidence, row.creditConfidence, (row.rationaleTags || []).join("|"), row.rationaleText, row.rateRationaleText, row.curveRationaleText, row.creditRationaleText]),
    [],
    weeklySummaryHeader,
    ...state.weeklySummaries.map(row => ["weeklySummary", row.week, row.rateDuration, row.curveDuration, row.creditDuration, row.ktb3yRangeLow, row.ktb3yRangeHigh, row.curveSpreadRangeLow, row.curveSpreadRangeHigh, row.creditSpreadRangeLow, row.creditSpreadRangeHigh, row.confidence, row.rateConfidence, row.curveConfidence, row.creditConfidence, row.overallSummaryText, row.rateSummaryText, row.curveSummaryText, row.creditSummaryText, row.status, row.draftSource, row.completedAt, row.createdAt, row.updatedAt]),
    [],
    marketHeader,
    ...state.marketData.map(row => ["market", row.week, row.asOf, row.ktb3y, row.ktb10y, row.curveSpread, row.msb2y, row.creditAA2y, row.creditSpread, row.memo]),
    [],
    dailyRateHeader,
    ...state.dailyRates.map(row => ["dailyRate", row.date, row.ktb3y, row.ktb10y, row.curveSpread, row.msb2y, row.creditAA2y, row.creditSpread]),
    [],
    mappingHeader,
    ...state.marketWeekMappings.map(row => ["marketWeekMapping", row.week, row.baseDate, row.memo, row.createdAt, row.updatedAt]),
    [],
    mappingHistoryHeader,
    ...state.marketMappingHistory.map(row => ["marketMappingHistory", row.week, row.previousBaseDate, row.baseDate, row.memo, row.changedAt]),
    [],
    archiveHeader,
    ...state.weeklyArchives.map(row => ["weeklyArchive", row.archiveId, row.week, row.version, row.status, row.archivedAt, row.archivedBy, row.memo, row.calculationVersion, row.sourceHash]),
    [],
    archiveOpinionHeader,
    ...state.weeklyArchiveOpinionStrategies.map(row => ["weeklyArchiveOpinionStrategy", row.archiveId, row.week, row.version, row.memberId, row.memberName, row.team, row.strategyKey, row.duration, row.rangeLow, row.rangeHigh, row.confidence, row.rationaleText]),
    [],
    archiveResultHeader,
    ...state.weeklyArchiveResults.map(row => ["weeklyArchiveResult", row.archiveId, row.week, row.version, row.memberId, row.memberName, row.strategyKey, row.evaluatedWeek, row.marketDelta, row.pnl, row.ytdPnl, row.rangeHit])
  ];
  return rows.map(row => row.map(csvEscape).join(",")).join("\n");
}

if (microsoftConfigReady()) {
  setMicrosoftStatus(redirectOriginMatches() ? "Microsoft login ready" : "Open the registered redirect URL.", redirectOriginMatches() ? "" : "error");
  try {
    const cachedAccount = getMsalClient().getAllAccounts?.()[0];
    if (cachedAccount) {
      msAccount = cachedAccount;
      setMicrosoftAccountBadge(cachedAccount);
    }
  } catch (e) {
    /* config/redirect 미준비 시 무시 — 로그인 버튼 누르면 정상 표시됨 */
  }
} else {
  setMicrosoftStatus("Microsoft config required", "error");
}

async function signInMicrosoftAndLoad() {
  const account = await signInMicrosoft();
  if (account) {
    await loadSharePointState();
  }
}

els.msSignInButton?.addEventListener("click", signInMicrosoftAndLoad);
els.graphTestButton?.addEventListener("click", testSharePointConnection);
els.sharePointLoadButton?.addEventListener("click", loadSharePointState);

els.navButtons.forEach(button => {
  button.addEventListener("click", () => {
    const view = button.dataset.view;
    els.navButtons.forEach(item => item.classList.toggle("active", item === button));
    els.views.forEach(item => item.classList.toggle("active", item.id === `${view}View`));
    els.pageTitle.textContent = document.querySelector(`#${view}View`).dataset.title;
    render();
  });
});

els.weekSelect.addEventListener("change", event => {
  selectedWeek = event.target.value;
  selectedArchiveId = "";
  render();
});

els.opinionWeek?.addEventListener("change", event => {
  const newWeek = event.target.value;

  if (opinionFormHasUnsavedInput() &&
      !confirm("저장하지 않은 입력 내용이 있습니다.\n버리고 선택한 회의의 저장 의견을 불러올까요?")) {
    event.target.value = selectedOpinionWeek;
    return;
  }

  selectedOpinionWeek = newWeek;
  updateOpinionLookupStatus();
  loadSelectedOpinionIntoForm({ notify: false });
});

els.loadOpinionButton?.addEventListener("click", async () => {
  const week = els.opinionWeek?.value || selectedOpinionWeek;
  const memberId = els.opinionMember?.value || "";
  if (!state.weeklyOpinions.length) {
    updateOpinionLookupStatus("저장 의견을 불러오는 중입니다. Microsoft 로그인 창이 뜨면 로그인해 주세요.");
    await loadSharePointState();
    if ([...els.opinionWeek.options].some(option => option.value === week)) {
      els.opinionWeek.value = week;
      selectedOpinionWeek = week;
    }
    if ([...els.opinionMember.options].some(option => option.value === memberId)) {
      els.opinionMember.value = memberId;
    }
  }
  const loaded = loadSelectedOpinionIntoForm({ notify: false });
  const message = loaded
    ? `불러옴: ${weekDisplayLabel(week)} / ${memberById(memberId).name}`
    : `저장된 의견 없음: ${weekDisplayLabel(week)} / ${memberById(memberId).name}`;
  updateOpinionLookupStatus(message);
  if (!loaded) {
    const loadedCount = state.weeklyOpinions.length;
    alert(loadedCount
      ? `${message}\n\n선택한 회의와 구성원을 다시 확인해 주세요.`
      : `${message}\n\n현재 화면에 불러온 의견 데이터가 없습니다. Login 후 Load를 먼저 눌러 SharePoint 데이터를 불러와 주세요.`);
  }
});

function handleOpinionStrategyStatusChange(event) {
  const target = event.target;
  if (target.dataset.strategyStatus || /RangeSkip$/.test(target.name || "")) {
    updateOpinionStrategyControls();
  }
}

els.opinionForm?.addEventListener("change", handleOpinionStrategyStatusChange);
els.opinionForm?.addEventListener("input", handleOpinionStrategyStatusChange);

els.weeklySummaryWeek?.addEventListener("change", event => {
  populateWeeklySummaryForm(event.target.value);
});

els.archiveWeekFilter?.addEventListener("change", event => {
  const week = event.target.value;
  const archive = activeArchiveForWeek(week) || latestArchiveForWeek(week);
  selectedArchiveId = archive?.archiveId || "";
  renderArchiveView();
});

els.archiveVersionFilter?.addEventListener("change", event => {
  selectedArchiveId = event.target.value;
  renderArchiveView();
});

els.archiveSummaryList?.addEventListener("click", event => {
  const item = event.target.closest("[data-archive-id]");
  if (!item) return;
  selectedArchiveId = item.dataset.archiveId || "";
  renderArchiveView();
});

els.durationCharts.addEventListener("change", event => {
  const strategyKey = event.target.dataset.chartMode;
  if (!strategyKey) return;
  strategyChartModes[strategyKey] = event.target.value;
  const strategy = STRATEGIES.find(item => item.key === strategyKey);
  if (strategy) drawStrategyChart(document.querySelector(`#duration-${strategy.key}`), strategy, strategyChartModes[strategy.key]);
});

els.currentWeekButton.addEventListener("click", () => {
  selectedWeek = nextOpinionWeek();
  selectedArchiveId = "";
  render();
});

els.opinionForm.addEventListener("submit", async event => {
  event.preventDefault();
  const data = formData(els.opinionForm);
  const activeStrategies = STRATEGIES.filter(strategy => !strategyIsSkipped(data, strategy));
  const rangeIncomplete = activeStrategies.find(strategy =>
    !rangeIsSkipped(data, strategy) &&
    (nullableNumber(data[strategy.rangeLowKey]) === null || nullableNumber(data[strategy.rangeHighKey]) === null)
  );
  if (rangeIncomplete) {
    alert(`${rangeIncomplete.title}: 레인지 상·하단을 모두 입력하거나 '레인지 생략'을 선택하세요.`);
    return;
  }
  const rangeOrderBad = activeStrategies.find(strategy =>
    !rangeIsSkipped(data, strategy) &&
    Number(data[strategy.rangeLowKey]) > Number(data[strategy.rangeHighKey])
  );
  if (rangeOrderBad) { alert("레인지 하단은 상단보다 낮아야 합니다."); return; }
  if (!data.memberId) {
    alert("로그인 계정과 일치하는 본부원을 찾지 못해 저장할 수 없습니다.\nDAOL_FI_Members의 email 칸을 확인하세요.");
    return;
  }
  const member = memberById(data.memberId);
  const existingOpinion = state.weeklyOpinions.find(row => row.week === data.week && row.memberId === data.memberId);
  const confirmMessage = [
    `${weekDisplayLabel(data.week)} / ${member.name} 의견을 저장할까요?`,
    existingOpinion ? "이미 저장된 의견이 있어 기존 내용을 수정합니다." : "새 의견으로 저장합니다."
  ].join("\n");
  if (!confirm(confirmMessage)) return;
  const confidenceValues = activeStrategies
    .map(strategy => numericInputValue(data[`${strategy.key}Confidence`]))
    .filter(value => value !== null);
  const rationaleParts = [
    ["금리 방향성", data.rateRationaleText],
    ["커브", data.curveRationaleText],
    ["크레딧", data.creditRationaleText]
  ].filter(([, text]) => String(text || "").trim());
  const savedOpinion = upsertOpinion({
    week: data.week,
    memberId: data.memberId,
    rateDuration: strategyIsSkipped(data, STRATEGIES[0]) ? null : numericInputValue(data.rateDuration),
    curveDuration: strategyIsSkipped(data, STRATEGIES[1]) ? null : numericInputValue(data.curveDuration),
    creditDuration: strategyIsSkipped(data, STRATEGIES[2]) ? null : numericInputValue(data.creditDuration),
    ktb3yRangeLow: (strategyIsSkipped(data, STRATEGIES[0]) || rangeIsSkipped(data, STRATEGIES[0])) ? null : numericInputValue(data.ktb3yRangeLow),
    ktb3yRangeHigh: (strategyIsSkipped(data, STRATEGIES[0]) || rangeIsSkipped(data, STRATEGIES[0])) ? null : numericInputValue(data.ktb3yRangeHigh),
    curveSpreadRangeLow: (strategyIsSkipped(data, STRATEGIES[1]) || rangeIsSkipped(data, STRATEGIES[1])) ? null : numericInputValue(data.curveSpreadRangeLow),
    curveSpreadRangeHigh: (strategyIsSkipped(data, STRATEGIES[1]) || rangeIsSkipped(data, STRATEGIES[1])) ? null : numericInputValue(data.curveSpreadRangeHigh),
    creditSpreadRangeLow: (strategyIsSkipped(data, STRATEGIES[2]) || rangeIsSkipped(data, STRATEGIES[2])) ? null : numericInputValue(data.creditSpreadRangeLow),
    creditSpreadRangeHigh: (strategyIsSkipped(data, STRATEGIES[2]) || rangeIsSkipped(data, STRATEGIES[2])) ? null : numericInputValue(data.creditSpreadRangeHigh),
    confidence: nullableNumber(Math.round(avg(confidenceValues))),
    rateConfidence: strategyIsSkipped(data, STRATEGIES[0]) ? null : numericInputValue(data.rateConfidence),
    curveConfidence: strategyIsSkipped(data, STRATEGIES[1]) ? null : numericInputValue(data.curveConfidence),
    creditConfidence: strategyIsSkipped(data, STRATEGIES[2]) ? null : numericInputValue(data.creditConfidence),
    rationaleTags: [],
    rationaleText: rationaleParts.map(([label, text]) => `${label}: ${String(text).trim()}`).join("\n"),
    rateRationaleText: data.rateRationaleText.trim(),
    curveRationaleText: data.curveRationaleText.trim(),
    creditRationaleText: data.creditRationaleText.trim()
  });
  selectedOpinionWeek = data.week;
  saveState();
  render();
  clearOpinionInputs(false);
  try {
    setMicrosoftStatus("SharePoint 의견 저장 중...");
    const result = await saveSharePointWeeklyOpinion(savedOpinion);
    setMicrosoftStatus(`SharePoint 의견 ${result.mode === "created" ? "생성" : "업데이트"} 성공`, "connected");
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`의견 로컬 저장됨, SharePoint 실패: ${shortErrorMessage(error)}`, "error");
    alert(`의견은 이 브라우저에 저장됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
  }
});

els.weeklySummaryDraftButton?.addEventListener("click", () => {
  const week = els.weeklySummaryWeek?.value || selectedWeeklySummaryWeek;
  const draft = generateWeeklySummaryDraft(week);
  if (!draft) {
    alert("선택한 주차에 입력된 조직원 의견이 없어 초안을 만들 수 없습니다.");
    return;
  }
  els.weeklySummaryForm.elements.overallSummaryText.value = draft.overallSummaryText;
  els.weeklySummaryForm.elements.rateDuration.value = draft.rateDuration;
  els.weeklySummaryForm.elements.ktb3yRangeHigh.value = draft.ktb3yRangeHigh;
  els.weeklySummaryForm.elements.ktb3yRangeLow.value = draft.ktb3yRangeLow;
  els.weeklySummaryForm.elements.rateConfidence.value = Number.isFinite(draft.rateConfidence) ? draft.rateConfidence : "3";
  els.weeklySummaryForm.elements.rateSummaryText.value = draft.rateSummaryText;
  els.weeklySummaryForm.elements.curveDuration.value = draft.curveDuration;
  els.weeklySummaryForm.elements.curveSpreadRangeHigh.value = draft.curveSpreadRangeHigh;
  els.weeklySummaryForm.elements.curveSpreadRangeLow.value = draft.curveSpreadRangeLow;
  els.weeklySummaryForm.elements.curveConfidence.value = Number.isFinite(draft.curveConfidence) ? draft.curveConfidence : "3";
  els.weeklySummaryForm.elements.curveSummaryText.value = draft.curveSummaryText;
  els.weeklySummaryForm.elements.creditDuration.value = draft.creditDuration;
  els.weeklySummaryForm.elements.creditSpreadRangeHigh.value = draft.creditSpreadRangeHigh;
  els.weeklySummaryForm.elements.creditSpreadRangeLow.value = draft.creditSpreadRangeLow;
  els.weeklySummaryForm.elements.creditConfidence.value = Number.isFinite(draft.creditConfidence) ? draft.creditConfidence : "3";
  els.weeklySummaryForm.elements.creditSummaryText.value = draft.creditSummaryText;
  if (els.weeklySummaryStatus) {
    els.weeklySummaryStatus.textContent = "AI 초안이 입력되었습니다. 검토 후 완료 저장하세요.";
  }
});

els.weeklySummaryForm?.addEventListener("submit", async event => {
  event.preventDefault();
  const data = formData(els.weeklySummaryForm);
  const rangePairs = [
    ["ktb3yRangeLow", "ktb3yRangeHigh"],
    ["curveSpreadRangeLow", "curveSpreadRangeHigh"],
    ["creditSpreadRangeLow", "creditSpreadRangeHigh"]
  ];
  if (rangePairs.some(([low, high]) => Number(data[low]) > Number(data[high]))) {
    alert("레인지 하단은 상단보다 낮아야 합니다.");
    return;
  }
  const confidenceValues = [data.rateConfidence, data.curveConfidence, data.creditConfidence].map(Number);
  const savedSummary = upsertWeeklySummary({
    week: data.week,
    memberId: SUMMARY_MEMBER_ID,
    rateDuration: Number(data.rateDuration),
    curveDuration: Number(data.curveDuration),
    creditDuration: Number(data.creditDuration),
    ktb3yRangeLow: Number(data.ktb3yRangeLow),
    ktb3yRangeHigh: Number(data.ktb3yRangeHigh),
    curveSpreadRangeLow: Number(data.curveSpreadRangeLow),
    curveSpreadRangeHigh: Number(data.curveSpreadRangeHigh),
    creditSpreadRangeLow: Number(data.creditSpreadRangeLow),
    creditSpreadRangeHigh: Number(data.creditSpreadRangeHigh),
    confidence: Math.round(avg(confidenceValues)),
    rateConfidence: Number(data.rateConfidence),
    curveConfidence: Number(data.curveConfidence),
    creditConfidence: Number(data.creditConfidence),
    overallSummaryText: String(data.overallSummaryText || "").trim(),
    rateSummaryText: String(data.rateSummaryText || "").trim(),
    curveSummaryText: String(data.curveSummaryText || "").trim(),
    creditSummaryText: String(data.creditSummaryText || "").trim(),
    draftSource: "member-opinion-draft"
  });
  selectedWeeklySummaryWeek = data.week;
  saveState();
  render();
  try {
    setMicrosoftStatus("SharePoint 종합의견 저장 중...");
    const result = await saveSharePointWeeklySummary(savedSummary);
    setMicrosoftStatus(`SharePoint 종합의견 ${result.mode === "created" ? "생성" : "업데이트"} 성공`, "connected");
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`종합의견 로컬 저장됨, SharePoint 실패: ${shortErrorMessage(error)}`, "error");
    alert(`종합의견은 이 브라우저에 저장됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
  }
});

els.marketForm.addEventListener("submit", async event => {
  event.preventDefault();
  const data = formData(els.marketForm);
  let savedMarket = null;
  let savedMapping = null;
  let savedHistory = null;
  try {
    const result = upsertMarket(data);
    savedMarket = result.market;
    savedMapping = result.mapping;
    savedHistory = result.history;
  } catch {
    alert("선택한 기준일의 업로드 금리 데이터가 없습니다.");
    return;
  }
  selectedWeek = data.week;
  saveState();
  render();
  try {
    setMicrosoftStatus("SharePoint 시장 데이터 저장 중...");
    const [marketResult, mappingResult, historyResult] = await Promise.all([
      saveSharePointMarket(savedMarket),
      saveSharePointWeekMapping(savedMapping),
      saveSharePointMappingHistory(savedHistory)
    ]);
    setMicrosoftStatus(`SharePoint market ${marketResult.mode}, mapping ${mappingResult.mode}, history ${historyResult.mode}`, "connected");
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`시장 데이터 로컬 저장됨, SharePoint 실패: ${shortErrorMessage(error)}`, "error");
    alert(`시장 데이터는 이 브라우저에 저장됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
  }
});

els.memberForm.addEventListener("submit", async event => {
  event.preventDefault();
  const data = formData(els.memberForm);
  const sortOrder = data.sortOrder.trim();
  if (!/^\d+$/.test(sortOrder)) {
    alert("ID는 숫자만 입력하세요.");
    els.memberForm.elements.sortOrder.focus();
    return;
  }
  const memberId = data.memberId || memberIdFromSortOrder(sortOrder);
  if (!data.memberId && state.members.some(member => member.id === memberId)) {
    alert(`${memberId} ID가 이미 있습니다. 다른 ID 숫자를 입력해 주세요.`);
    els.memberForm.elements.sortOrder.focus();
    return;
  }
  const savedMember = upsertMember({
    memberId,
    name: data.name.trim(),
    team: data.team.trim(),
    sortOrder: Number(sortOrder),
    active: data.active === "true"
  });
  els.memberForm.reset();
  saveState();
  render();
  try {
    setMicrosoftStatus("SharePoint 조직원 저장 중...");
    const result = await saveSharePointMember(savedMember);
    setMicrosoftStatus(`SharePoint 조직원 ${result.mode === "created" ? "생성" : "업데이트"} 성공`, "connected");
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`조직원 로컬 저장됨, SharePoint 실패: ${shortErrorMessage(error)}`, "error");
    alert(`조직원은 이 브라우저에 저장됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
  }
});

els.memberForm.elements.sortOrder.addEventListener("input", event => {
  event.target.value = event.target.value.replace(/\D/g, "");
});

els.memberTable.addEventListener("click", async event => {
  const deleteId = event.target.dataset.deleteMember;
  if (!deleteId) return;
  event.preventDefault();
  event.stopImmediatePropagation();

  const member = state.members.find(item => item.id === deleteId);
  if (!member) return;
  const hasOpinions = state.weeklyOpinions.some(row => row.memberId === deleteId);

  if (hasOpinions) {
    if (!confirm("과거 의견이 있어 삭제 대신 비활성화합니다. 계속할까요?")) return;
    member.active = false;
    saveState();
    render();
    try {
      setMicrosoftStatus("SharePoint 조직원 비활성화 저장 중...");
      await saveSharePointMember(member);
      setMicrosoftStatus("SharePoint 조직원 비활성화 성공", "connected");
    } catch (error) {
      console.error(error);
      setMicrosoftStatus(`조직원 로컬 비활성화됨, SharePoint 실패: ${shortErrorMessage(error)}`, "error");
      alert(`조직원은 이 브라우저에서 비활성화됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
    }
    return;
  }

  if (!confirm("과거 의견이 없는 조직원입니다. SharePoint에서도 완전히 삭제할까요?")) return;
  const previousMembers = state.members;
  state.members = state.members.filter(item => item.id !== deleteId);
  saveState();
  render();
  try {
    setMicrosoftStatus("SharePoint 조직원 삭제 중...");
    await deleteSharePointMember(member);
    setMicrosoftStatus("SharePoint 조직원 삭제 성공", "connected");
  } catch (error) {
    console.error(error);
    state.members = previousMembers;
    saveState();
    render();
    setMicrosoftStatus(`SharePoint 조직원 삭제 실패: ${shortErrorMessage(error)}`, "error");
    alert(`SharePoint 조직원 삭제에 실패해 로컬 삭제를 되돌렸습니다.\n\n${error.message}`);
  }
}, true);

els.memberTable.addEventListener("click", async event => {
  const editId = event.target.dataset.editMember;
  const deleteId = event.target.dataset.deleteMember;
  if (editId) {
    const member = state.members.find(item => item.id === editId);
    if (!member) return;
    els.memberForm.elements.memberId.value = member.id;
    els.memberForm.elements.name.value = member.name;
    els.memberForm.elements.team.value = member.team;
    els.memberForm.elements.sortOrder.value = member.sortOrder || 10;
    els.memberForm.elements.active.value = String(member.active !== false);
  }
  if (deleteId) {
    const member = state.members.find(item => item.id === deleteId);
    if (!member) return;
    if (!confirm("해당 조직원을 비활성화할까요? SharePoint에는 삭제 대신 비활성화 상태로 저장합니다.")) return;
    member.active = false;
    saveState();
    render();
    try {
      setMicrosoftStatus("SharePoint 조직원 비활성화 저장 중...");
      await saveSharePointMember(member);
      setMicrosoftStatus("SharePoint 조직원 비활성화 성공", "connected");
    } catch (error) {
      console.error(error);
      setMicrosoftStatus(`조직원 로컬 비활성화됨, SharePoint 실패: ${shortErrorMessage(error)}`, "error");
      alert(`조직원은 이 브라우저에서 비활성화됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
    }
    return;
    const hasOpinions = state.weeklyOpinions.some(row => row.memberId === deleteId);
    if (hasOpinions) {
      if (!confirm("과거 의견이 있어 완전 삭제 대신 비활성 처리합니다. 계속할까요?")) return;
      const member = state.members.find(item => item.id === deleteId);
      if (member) member.active = false;
    } else {
      if (!confirm("이 본부원을 삭제할까요?")) return;
      state.members = state.members.filter(item => item.id !== deleteId);
    }
    saveState();
    render();
  }
});

els.clearOpinionButton.addEventListener("click", () => {
  els.opinionForm.reset();
  renderFormsDefaults();
});

els.clearWeeklySummaryButton?.addEventListener("click", () => {
  if (!els.weeklySummaryForm) return;
  els.weeklySummaryForm.elements.overallSummaryText.value = "";
  els.weeklySummaryForm.elements.rateDuration.value = "";
  els.weeklySummaryForm.elements.ktb3yRangeHigh.value = "";
  els.weeklySummaryForm.elements.ktb3yRangeLow.value = "";
  els.weeklySummaryForm.elements.rateConfidence.value = "3";
  els.weeklySummaryForm.elements.rateSummaryText.value = "";
  els.weeklySummaryForm.elements.curveDuration.value = "";
  els.weeklySummaryForm.elements.curveSpreadRangeHigh.value = "";
  els.weeklySummaryForm.elements.curveSpreadRangeLow.value = "";
  els.weeklySummaryForm.elements.curveConfidence.value = "3";
  els.weeklySummaryForm.elements.curveSummaryText.value = "";
  els.weeklySummaryForm.elements.creditDuration.value = "";
  els.weeklySummaryForm.elements.creditSpreadRangeHigh.value = "";
  els.weeklySummaryForm.elements.creditSpreadRangeLow.value = "";
  els.weeklySummaryForm.elements.creditConfidence.value = "3";
  els.weeklySummaryForm.elements.creditSummaryText.value = "";
});

els.clearMarketButton.addEventListener("click", () => {
  els.marketForm.reset();
  renderFormsDefaults();
  renderMarketDatePreview();
});

els.archiveWeekButton?.addEventListener("click", async () => {
  const week = selectedWeek;
  if (!opinionsForWeek(week).length) {
    alert("Archive needs at least one opinion for the selected week.");
    return;
  }
  const existing = activeArchiveForWeek(week);
  const promptText = existing
    ? `Re-archive ${weekDisplayLabel(week)}? Previous active archive will be superseded locally.`
    : `Archive ${weekDisplayLabel(week)}?`;
  if (!confirm(promptText)) return;
  const snapshot = buildWeeklyArchiveSnapshot(week, els.archiveMemo?.value || "");
  upsertLocalArchiveSnapshot(snapshot);
  selectedArchiveId = snapshot.header.archiveId;
  saveState();
  render();
  try {
    setMicrosoftStatus("SharePoint archive storage checking...");
    const result = await saveSharePointArchiveSnapshot(snapshot);
    const savedCount = result.results.reduce((sum, item) => sum + item.total, 0);
    const skippedText = result.skipped.length ? `, skipped ${result.skipped.length} missing lists` : "";
    setMicrosoftStatus(`SharePoint archive saved: ${savedCount} rows${skippedText}`, "connected");
  } catch (error) {
    console.error(error);
    setMicrosoftStatus(`Archive local saved, SharePoint failed: ${shortErrorMessage(error)}`, "error");
    alert(`Archive was saved locally, but SharePoint save failed.\n\n${error.message}`);
  }
});

els.marketWeek.addEventListener("change", event => {
  const mapping = state.marketWeekMappings.find(item => item.week === event.target.value);
  if (mapping) {
    els.marketBaseDate.value = mapping.baseDate;
    els.marketForm.elements.memo.value = mapping.memo || "";
  }
  renderMarketDatePreview();
});

els.marketBaseDate.addEventListener("change", renderMarketDatePreview);

els.rateFileInput.addEventListener("change", async event => {
  const [file] = event.target.files;
  if (!file) return;
  els.rateUploadStatus.textContent = "엑셀 파일을 읽는 중입니다...";
  try {
    const dailyRates = await parseRatesWorkbook(file);
    const uploadedFile = {
      id: id("rates-file"),
      fileName: file.name,
      uploadedAt: new Date().toISOString(),
      rowCount: dailyRates.length,
      minDate: dailyRates[0]?.date || "",
      maxDate: dailyRates.at(-1)?.date || ""
    };
    state.dailyRates = dailyRates;
    state.uploadedRateFiles.push(uploadedFile);
    syncMappedMarketsFromDailyRates();
    saveState();
    render();
    try {
      setMicrosoftStatus("SharePoint daily rates 저장 중...");
      const [rateResult, fileResult] = await Promise.all([
        saveSharePointDailyRates(dailyRates),
        saveSharePointUploadedRateFile(uploadedFile)
      ]);
      const mappedMarkets = state.marketWeekMappings
        .map(mapping => state.marketData.find(row => row.week === mapping.week))
        .filter(Boolean);
      for (const market of mappedMarkets) {
        await saveSharePointMarket(market);
      }
      setMicrosoftStatus(
        `SharePoint rates ${rateResult.total} saved, uploaded file ${fileResult.mode}, mapped markets ${mappedMarkets.length}`,
        "connected"
      );
    } catch (error) {
      console.error(error);
      setMicrosoftStatus(`rate upload local saved, SharePoint failed: ${shortErrorMessage(error)}`, "error");
      alert(`금리 파일은 이 브라우저에 저장됐지만 SharePoint 저장은 실패했습니다.\n\n${error.message}`);
    }
  } catch (error) {
    console.error(error);
    alert("엑셀 파일을 읽지 못했습니다. Raw 시트와 필요한 금리 컬럼을 확인해 주세요.");
    renderRateUploadState();
  } finally {
    event.target.value = "";
  }
});

els.clearMemberButton.addEventListener("click", () => {
  els.memberForm.reset();
  els.memberForm.elements.memberId.value = "";
});

els.exportJsonButton.addEventListener("click", () => {
  download(`daol-fi-strategy-v2-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(state, null, 2), "application/json");
});

els.exportCsvButton.addEventListener("click", () => {
  download(`daol-fi-strategy-v2-${new Date().toISOString().slice(0, 10)}.csv`, buildCsv(), "text/csv;charset=utf-8");
});

els.performanceMetricSelect.addEventListener("change", event => {
  selectedPerformanceMetric = event.target.value;
  renderPerformance();
});

els.topRankingMetricSelect?.addEventListener("change", event => {
  selectedTopRankingMetric = event.target.value;
  renderOverview();
});

els.importJsonInput.addEventListener("change", async event => {
  const [file] = event.target.files;
  if (!file) return;
  try {
    state = normalizeState(JSON.parse(await file.text()));
    saveState();
    selectedWeek = latestWeek();
    render();
  } catch {
    alert("JSON 파일을 읽을 수 없습니다.");
  } finally {
    event.target.value = "";
  }
});

window.addEventListener("resize", () => {
  STRATEGIES.forEach(strategy => {
    const durationSvg = document.querySelector(`#duration-${strategy.key}`);
    if (durationSvg) drawStrategyChart(durationSvg, strategy, strategyChartModes[strategy.key]);
  });
  if (els.weeklyPerformanceChart) drawWeeklyPerformanceChart();
  if (els.performanceChart) drawPerformanceChart();
});

render();
loadBundledRatesIfAvailable();
