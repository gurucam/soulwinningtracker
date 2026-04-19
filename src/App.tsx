import React, { useCallback, useEffect, useMemo, useRef, useState } from "react"
import "./App.css"
import { THEME_DEFS } from "./themes"
import { isSupabaseConfigured, supabase } from "./lib/supabaseClient"

type SessionType = "session" | "standalone"
type SalvationRole = "presenter" | "partner"

type GroupBy = "year" | "month" | "week" | "day"
type TagMatchMode = "any" | "all"
type GoalMetric = "sessions" | "salvations" | "doors"
type GoalPeriod =
  | "week"
  | "month"
  | "year"
  | "overall"
  | "last_week"
  | "last_month"
  | "last_year"
  | "custom"

type SessionRow = {
  id: string
  date: string
  name: string
  type: SessionType
  dataset_id: string
  saved_count: number
  doors_knocked: number | null
  tagIds: string[]
  start_time: string | null
  end_time: string | null
  notes: string
  created_at: string
}

type PersonRow = {
  id: string
  session_id: string
  name: string
  tagIds: string[]
  role: SalvationRole
  time_spent_minutes: number | null
  notes: string
  created_at: string
}

type TagRow = {
  id: string
  name: string
  color: string | null
  created_at: string
}

type DatasetRow = {
  id: string
  name: string
  created_at: string
}

type DraftPerson = {
  id: string
  name: string
  tagIds: string[]
  role: SalvationRole
  notes: string
}

type TagUsageEntry = {
  person: PersonRow
  session?: SessionRow
  dataset?: DatasetRow
}

type StatsSection =
  | "totals"
  | "charts"
  | "weekday"
  | "insights"
  | "dataset_breakdown"
  | "top_events"
  | "salvation_search"
  | "tag_breakdown"
  | "period_breakdown"
  | "time_metrics"

type StatsView = {
  id: string
  name: string
  sections: StatsSection[]
  created_at: string
}

type GoalDefinition = {
  id: string
  name: string
  metric: GoalMetric
  period: GoalPeriod
  target: number
  datasetId: string | "all"
  tagIds: string[]
  tagMatchMode: TagMatchMode
  startDate?: string
  endDate?: string
  created_at: string
}

type AuthViewMode = "signin" | "signup"
type FeedbackCategory = "general" | "bug" | "idea"

type StatsDetail =
  | {
      kind: "event"
      session: SessionRow
    }
  | {
      kind: "salvation"
      person: PersonRow
      session?: SessionRow
    }

type UpdateStatusPayload = {
  message?: string
  ready?: boolean
  progress?: number
}

type SoulwinningBridge = {
  version?: string
  checkForUpdates?: () => Promise<{ ok: boolean; error?: string } | void>
  installUpdate?: () => void
  getVersion?: () => Promise<string>
  openExternal?: (url: string) => void
  getUserDataPath?: () => Promise<string>
  onUpdateStatus?: (handler: (payload: UpdateStatusPayload) => void) => (() => void) | void
}

declare global {
  interface Window {
    soulwinning?: SoulwinningBridge
  }
}

const TAG_COLORS = [
  "#C56B3C",
  "#6A8D7C",
  "#D6A94C",
  "#7C5D7C",
  "#3F7D9D",
  "#C95B7C",
  "#8C6A4F",
  "#4E9E7F",
]

const TOP_TAG_LIMIT = 8
const TAG_SEARCH_LIMIT = 12
const TAG_DETAIL_PREVIEW_LIMIT = 3
const STATS_SEARCH_LIMIT = 6
const DEFAULT_GOALS: GoalDefinition[] = []
const SUPPORT_URL = "https://ko-fi.com/rrcam"
const DEFAULT_STATS_VIEWS: StatsView[] = [
  {
    id: "default-overview",
    name: "Overview",
    sections: ["totals", "time_metrics", "insights", "charts"],
    created_at: new Date().toISOString(),
  },
  {
    id: "default-detail",
    name: "Full breakdown",
    sections: [
      "totals",
      "time_metrics",
      "charts",
      "weekday",
      "dataset_breakdown",
      "top_events",
      "period_breakdown",
      "tag_breakdown",
      "salvation_search",
      "insights",
    ],
    created_at: new Date().toISOString(),
  },
]

const STATS_SECTION_OPTIONS: Array<{
  id: StatsSection
  label: string
  description: string
}> = [
  { id: "totals", label: "Totals", description: "Overall counts and averages." },
  { id: "time_metrics", label: "Time metrics", description: "Session duration and salvations per hour." },
  { id: "charts", label: "Charts", description: "Salvations/events by period and ratios." },
  { id: "weekday", label: "Weekday breakdown", description: "Events and salvations by weekday." },
  { id: "insights", label: "Insights", description: "Highlights and breakdowns." },
  { id: "dataset_breakdown", label: "Data set breakdown", description: "Totals by data set." },
  { id: "top_events", label: "Top events", description: "Highest-salvation events." },
  { id: "period_breakdown", label: "Period breakdown", description: "Table by the selected grouping." },
  { id: "tag_breakdown", label: "Tag breakdown", description: "Salvations by tag." },
  { id: "salvation_search", label: "Salvation details", description: "Searchable list of salvations." },
]

const numberFormatter = new Intl.NumberFormat("en-US")
const dateFormatter = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric",
  year: "numeric",
})
const timeFormatter = new Intl.DateTimeFormat("en-US", {
  hour: "numeric",
  minute: "2-digit",
})
const monthFormatter = new Intl.DateTimeFormat("en-US", {
  month: "short",
  year: "numeric",
})
const dateTimeFormatter = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric",
  year: "numeric",
  hour: "numeric",
  minute: "2-digit",
})

const normalizeTimeValue = (value: unknown) => {
  if (typeof value !== "string") return null
  const trimmed = value.trim()
  return /^\d{2}:\d{2}$/.test(trimmed) ? trimmed : null
}

const normalizeMinutesValue = (value: unknown) => {
  if (value === null || value === undefined) return null
  const parsed = typeof value === "number" ? value : Number.parseInt(String(value), 10)
  if (!Number.isFinite(parsed) || parsed < 0) return null
  return Math.floor(parsed)
}

const formatTimeValue = (time: string | null | undefined) => {
  if (!time) return ""
  const [hours, minutes] = time.split(":").map((part) => Number.parseInt(part, 10))
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return time
  const date = new Date()
  date.setHours(hours, minutes, 0, 0)
  return timeFormatter.format(date)
}

const formatDurationMinutes = (minutes: number | null | undefined) => {
  if (minutes === null || minutes === undefined) return ""
  const safeMinutes = Math.max(0, Math.floor(minutes))
  const hours = Math.floor(safeMinutes / 60)
  const remaining = safeMinutes % 60
  if (hours > 0 && remaining > 0) return `${hours}h ${remaining}m`
  if (hours > 0) return `${hours}h`
  return `${remaining}m`
}

const parseTimeToMinutes = (time: string | null | undefined) => {
  if (!time) return null
  const [hours, minutes] = time.split(":").map((part) => Number.parseInt(part, 10))
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return null
  return hours * 60 + minutes
}

const getSessionDurationMinutes = (startTime: string | null, endTime: string | null) => {
  const startMinutes = parseTimeToMinutes(startTime)
  const endMinutes = parseTimeToMinutes(endTime)
  if (startMinutes === null || endMinutes === null) return null
  if (endMinutes < startMinutes) return null
  return endMinutes - startMinutes
}

const formatCount = (count: number, singular: string, plural: string) => {
  const label = count === 1 ? singular : plural
  return `${numberFormatter.format(count)} ${label}`
}

const formatCountLabel = (count: number, singular: string, plural: string) => {
  const label = count === 1 ? singular : plural
  return `${label}: ${numberFormatter.format(count)}`
}

const formatSalvationCount = (count: number) => formatCount(count, "salvation", "salvations")
const formatEventCount = (count: number) => formatCount(count, "event", "events")
const formatSalvationCountLabel = (count: number) =>
  formatCountLabel(count, "Salvation", "Salvations")

const DEFAULT_SALVATION_ROLE: SalvationRole = "presenter"

const normalizeSalvationRole = (value: unknown): SalvationRole => {
  const normalized = typeof value === "string" ? value.trim().toLowerCase() : ""
  if (
    normalized === "partner" ||
    normalized === "silent_partner" ||
    normalized === "silent partner" ||
    normalized === "silent"
  ) {
    return "partner"
  }
  return "presenter"
}

const formatSalvationRoleLabel = (role: SalvationRole | null | undefined) =>
  role === "partner" ? "Silent partner" : "Preacher"

const createLocalId = () => {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID()
  }
  return `id-${Math.random().toString(36).slice(2, 10)}`
}

const slugify = (value: string) => {
  const slug = value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
  return slug || "dataset"
}

const createDefaultDataset = () => ({
  id: createLocalId(),
  name: "Personal",
  created_at: new Date().toISOString(),
})

const normalizeSessionType = (value: unknown): SessionType => {
  if (value === "standalone") return "standalone"
  return "session"
}

const normalizeDatasets = (raw: unknown) => {
  const datasets = Array.isArray(raw)
    ? raw.map((dataset) => ({
        id: typeof dataset.id === "string" ? dataset.id : createLocalId(),
        name: typeof dataset.name === "string" && dataset.name.trim() ? dataset.name.trim() : "Unnamed",
        created_at: typeof dataset.created_at === "string" ? dataset.created_at : new Date().toISOString(),
      }))
    : []
  if (!datasets.length) {
    datasets.push(createDefaultDataset())
  }
  return datasets
}

type LegacyGoalSettings = {
  weeklySessions?: number
  weeklySalvations?: number
  monthlySessions?: number
  monthlySalvations?: number
  yearlySessions?: number
  yearlySalvations?: number
}

const normalizeGoalTarget = (value: unknown) => {
  const parsed = typeof value === "number" ? value : Number.parseInt(String(value), 10)
  if (!Number.isFinite(parsed) || parsed < 0) return 0
  return Math.floor(parsed)
}

const normalizeGoalMetric = (value: unknown): GoalMetric => {
  if (value === "salvations") return "salvations"
  if (value === "doors" || value === "doors_knocked") return "doors"
  return "sessions"
}

const normalizeGoalPeriod = (value: unknown): GoalPeriod => {
  if (
    value === "month" ||
    value === "year" ||
    value === "overall" ||
    value === "last_week" ||
    value === "last_month" ||
    value === "last_year" ||
    value === "custom"
  ) {
    return value
  }
  return "week"
}

const normalizeGoalDate = (value: unknown) => {
  if (typeof value !== "string") return ""
  const trimmed = value.trim()
  return /^\d{4}-\d{2}-\d{2}$/.test(trimmed) ? trimmed : ""
}

const normalizeTagMatchMode = (value: unknown): TagMatchMode => {
  if (value === "all") return "all"
  return "any"
}

const normalizeGoals = (raw: unknown): GoalDefinition[] => {
  if (Array.isArray(raw)) {
    return raw.map((goal) => {
      const data = goal as Partial<GoalDefinition>
      return {
        id: typeof data.id === "string" ? data.id : createLocalId(),
        name: typeof data.name === "string" ? data.name.trim() : "",
        metric: normalizeGoalMetric(data.metric),
        period: normalizeGoalPeriod(data.period),
        target: normalizeGoalTarget(data.target),
        datasetId: typeof data.datasetId === "string" ? data.datasetId : "all",
        tagIds: Array.isArray(data.tagIds) ? data.tagIds.filter((tag) => typeof tag === "string") : [],
        tagMatchMode: normalizeTagMatchMode(data.tagMatchMode),
        startDate: normalizeGoalDate(data.startDate) || undefined,
        endDate: normalizeGoalDate(data.endDate) || undefined,
        created_at:
          typeof data.created_at === "string" ? data.created_at : new Date().toISOString(),
      }
    })
  }

  if (raw && typeof raw === "object") {
    const data = raw as LegacyGoalSettings
    const legacyGoals: GoalDefinition[] = []
    const pushLegacy = (value: unknown, metric: GoalMetric, period: GoalPeriod, name: string) => {
      const target = normalizeGoalTarget(value)
      if (target <= 0) return
      legacyGoals.push({
        id: createLocalId(),
        name,
        metric,
        period,
        target,
        datasetId: "all",
        tagIds: [],
        tagMatchMode: "any",
        created_at: new Date().toISOString(),
      })
    }
    pushLegacy(data.weeklySessions, "sessions", "week", "Weekly events")
    pushLegacy(data.weeklySalvations, "salvations", "week", "Weekly salvations")
    pushLegacy(data.monthlySessions, "sessions", "month", "Monthly events")
    pushLegacy(data.monthlySalvations, "salvations", "month", "Monthly salvations")
    pushLegacy(data.yearlySessions, "sessions", "year", "Yearly events")
    pushLegacy(data.yearlySalvations, "salvations", "year", "Yearly salvations")
    return legacyGoals
  }

  return []
}

const normalizeStatsSection = (value: unknown): StatsSection | null => {
  if (
    value === "totals" ||
    value === "charts" ||
    value === "weekday" ||
    value === "insights" ||
    value === "dataset_breakdown" ||
    value === "top_events" ||
    value === "salvation_search" ||
    value === "tag_breakdown" ||
    value === "period_breakdown" ||
    value === "time_metrics"
  ) {
    return value
  }
  return null
}

const normalizeStatsViews = (raw: unknown): StatsView[] => {
  if (Array.isArray(raw)) {
    const fallbackSections: StatsSection[] = ["totals", "charts"]
    const views = raw
      .map((view) => {
        const data = view as Partial<StatsView>
        const sections = Array.isArray(data.sections)
          ? data.sections
              .map((section) => normalizeStatsSection(section))
              .filter((section): section is StatsSection => Boolean(section))
          : []
        return {
          id: typeof data.id === "string" ? data.id : createLocalId(),
          name: typeof data.name === "string" ? data.name.trim() : "",
          sections: sections.length ? sections : fallbackSections,
          created_at:
            typeof data.created_at === "string" ? data.created_at : new Date().toISOString(),
        }
      })
      .filter((view) => view.name)
    if (views.length) return views
  }
  return [...DEFAULT_STATS_VIEWS]
}

const getDefaultDatasetId = (datasets: DatasetRow[]) => {
  return datasets[0]?.id ?? ""
}

const normalizeSessions = (raw: unknown, defaultDatasetId: string) => {
  return Array.isArray(raw)
    ? raw.map((session) => {
        const data = session as Partial<SessionRow> & { tracker_id?: string }
        const datasetId =
          typeof data.dataset_id === "string"
            ? data.dataset_id
            : typeof data.tracker_id === "string"
              ? data.tracker_id
              : defaultDatasetId
        return {
          id: typeof data.id === "string" ? data.id : createLocalId(),
          date: typeof data.date === "string" ? data.date : toDateKey(new Date()),
          name: typeof data.name === "string" ? data.name : "",
          type: normalizeSessionType(data.type),
          dataset_id: datasetId,
          saved_count: typeof data.saved_count === "number" ? data.saved_count : 0,
          doors_knocked: typeof data.doors_knocked === "number" ? data.doors_knocked : null,
          tagIds: Array.isArray(data.tagIds)
            ? data.tagIds.filter((tagId) => typeof tagId === "string")
            : [],
          start_time: normalizeTimeValue(data.start_time),
          end_time: normalizeTimeValue(data.end_time),
          notes: typeof data.notes === "string" ? data.notes : "",
          created_at:
            typeof data.created_at === "string" ? data.created_at : new Date().toISOString(),
        }
      })
    : []
}

const toDateKey = (date: Date) => {
  const year = date.getFullYear()
  const month = `${date.getMonth() + 1}`.padStart(2, "0")
  const day = `${date.getDate()}`.padStart(2, "0")
  return `${year}-${month}-${day}`
}

const startOfWeek = (date: Date, weekStartsOn: number) => {
  const day = date.getDay()
  const diff = (day - weekStartsOn + 7) % 7
  const start = new Date(date)
  start.setDate(date.getDate() - diff)
  start.setHours(0, 0, 0, 0)
  return start
}

const getGroupInfo = (dateValue: string, groupBy: GroupBy, weekStartsOn: number) => {
  const date = new Date(`${dateValue}T00:00:00`)
  const year = date.getFullYear()
  const month = date.getMonth()

  if (groupBy === "year") {
    return {
      key: `${year}`,
      label: `${year}`,
      sortValue: new Date(year, 0, 1).getTime(),
    }
  }

  if (groupBy === "month") {
    return {
      key: `${year}-${`${month + 1}`.padStart(2, "0")}`,
      label: monthFormatter.format(date),
      sortValue: new Date(year, month, 1).getTime(),
    }
  }

  if (groupBy === "week") {
    const start = startOfWeek(date, weekStartsOn)
    return {
      key: `week-${toDateKey(start)}`,
      label: `Week of ${dateFormatter.format(start)}`,
      sortValue: start.getTime(),
    }
  }

  return {
    key: dateValue,
    label: dateFormatter.format(date),
    sortValue: date.getTime(),
  }
}

const isWithinRange = (dateValue: string, start: string, end: string) => {
  if (start && dateValue < start) return false
  if (end && dateValue > end) return false
  return true
}

const escapeCsvValue = (value: string | number | null | undefined) => {
  if (value === null || value === undefined) return ""
  const stringValue = String(value)
  if (/[",\n]/.test(stringValue)) {
    return `"${stringValue.replace(/"/g, "\"\"")}"`
  }
  return stringValue
}

const downloadCsv = (filename: string, rows: Array<Record<string, string | number | null | undefined>>) => {
  if (!rows.length) return
  const headerSet = new Set<string>()
  rows.forEach((row) => {
    Object.keys(row).forEach((key) => headerSet.add(key))
  })
  const headers = Array.from(headerSet)
  const lines = [headers.join(",")]

  rows.forEach((row) => {
    const line = headers.map((header) => escapeCsvValue(row[header])).join(",")
    lines.push(line)
  })

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" })
  const link = document.createElement("a")
  link.href = URL.createObjectURL(blob)
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(link.href)
}

type AppData = {
  version: 1
  sessions: SessionRow[]
  people: PersonRow[]
  tags: TagRow[]
  datasets: DatasetRow[]
  goals: GoalDefinition[]
  statsViews: StatsView[]
}

type UserSyncState = {
  signature: string
  synced_at: string
  uploaded_at: string
  downloaded_at: string
}

type SyncStateStore = Record<string, UserSyncState>

type RollbackSnapshot = {
  saved_at: string
  reason: string
  data: AppData
}

const STORAGE_KEY = "soulwinning-tracker-data-v1"
const SYNC_STATE_KEY = "soulwinning-tracker-sync-state-v1"
const ROLLBACK_KEY = "soulwinning-tracker-rollback-v1"
const THEME_MODE_KEY = "soulwinning_theme"
const THEME_NAME_KEY = "soulwinning_theme_name"
const SUPABASE_SNAPSHOT_TABLE = "user_snapshots"
const SUPABASE_FEEDBACK_TABLE = "feedback_submissions"
const AUTO_SYNC_DEBOUNCE_MS = 1500
const EXCLUDED_THEME_IDS = new Set([
  "dark-souls-2",
  "eva-00",
  "eva-01",
  "eva-02",
  "gp",
  "steven",
  "matt",
])
const AVAILABLE_THEMES = THEME_DEFS.filter((themeDef) => !EXCLUDED_THEME_IDS.has(themeDef.id))
const DEFAULT_THEME_ID =
  AVAILABLE_THEMES.find((themeDef) => themeDef.id === "default")?.id ??
  AVAILABLE_THEMES[0]?.id ??
  "default"

const getStoredThemeMode = () => {
  try {
    const stored = localStorage.getItem(THEME_MODE_KEY)
    return stored === "dark" || stored === "light" ? stored : "light"
  } catch {
    return "light"
  }
}

const getStoredThemeName = () => {
  try {
    const stored = localStorage.getItem(THEME_NAME_KEY)
    if (AVAILABLE_THEMES.some((themeDef) => themeDef.id === stored)) {
      return stored ?? DEFAULT_THEME_ID
    }
  } catch {
    // ignore
  }
  return DEFAULT_THEME_ID
}

const buildEmptyData = (): AppData => ({
  version: 1,
  sessions: [],
  people: [],
  tags: [],
  datasets: [createDefaultDataset()],
  goals: [...DEFAULT_GOALS],
  statsViews: [...DEFAULT_STATS_VIEWS],
})

const normalizeTimestampValue = (value: unknown) => {
  if (typeof value !== "string") return ""
  const trimmed = value.trim()
  if (!trimmed) return ""
  const parsed = Date.parse(trimmed)
  if (!Number.isFinite(parsed)) return ""
  return new Date(parsed).toISOString()
}

const formatTimestampLabel = (value: string) => {
  if (!value) return ""
  const parsed = Date.parse(value)
  if (!Number.isFinite(parsed)) return ""
  return dateTimeFormatter.format(new Date(parsed))
}

const buildCloudSignature = (userId: string, data: AppData) => `${userId}:${JSON.stringify(data)}`

const hasMeaningfulData = (data: AppData) =>
  data.sessions.length > 0 || data.people.length > 0 || data.tags.length > 0 || data.goals.length > 0

const loadSyncStateStore = (): SyncStateStore => {
  if (typeof localStorage === "undefined") return {}
  try {
    const raw = localStorage.getItem(SYNC_STATE_KEY)
    if (!raw) return {}
    const parsed = JSON.parse(raw)
    if (!parsed || typeof parsed !== "object") return {}
    return parsed as SyncStateStore
  } catch {
    return {}
  }
}

const saveSyncStateStore = (store: SyncStateStore) => {
  if (typeof localStorage === "undefined") return
  localStorage.setItem(SYNC_STATE_KEY, JSON.stringify(store))
}

const loadUserSyncState = (userId: string): UserSyncState | null => {
  if (!userId) return null
  const store = loadSyncStateStore()
  const state = store[userId]
  if (!state || typeof state !== "object") return null
  const signature = typeof state.signature === "string" ? state.signature : ""
  if (!signature) return null
  return {
    signature,
    synced_at: normalizeTimestampValue(state.synced_at) || "",
    uploaded_at: normalizeTimestampValue(state.uploaded_at) || "",
    downloaded_at: normalizeTimestampValue(state.downloaded_at) || "",
  }
}

const saveUserSyncState = (userId: string, state: UserSyncState) => {
  if (!userId) return
  const store = loadSyncStateStore()
  store[userId] = state
  saveSyncStateStore(store)
}

const loadRollbackSnapshot = (): RollbackSnapshot | null => {
  if (typeof localStorage === "undefined") return null
  try {
    const raw = localStorage.getItem(ROLLBACK_KEY)
    if (!raw) return null
    const parsed = JSON.parse(raw)
    if (!parsed || typeof parsed !== "object") return null
    const snapshot = parsed as Partial<RollbackSnapshot>
    if (!snapshot.data || typeof snapshot.data !== "object") return null
    const savedAt = normalizeTimestampValue(snapshot.saved_at) || ""
    if (!savedAt) return null
    return {
      saved_at: savedAt,
      reason: typeof snapshot.reason === "string" ? snapshot.reason : "",
      data: normalizeImportedData(snapshot.data),
    }
  } catch {
    return null
  }
}

const saveRollbackSnapshot = (snapshot: RollbackSnapshot) => {
  if (typeof localStorage === "undefined") return
  localStorage.setItem(ROLLBACK_KEY, JSON.stringify(snapshot))
}

const loadAppData = () => {
  if (typeof localStorage === "undefined") return buildEmptyData()
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return buildEmptyData()
    const parsed = JSON.parse(raw) as Partial<AppData> & { trackers?: unknown }
    const datasets = normalizeDatasets(parsed.datasets ?? parsed.trackers)
    const defaultDatasetId = getDefaultDatasetId(datasets)
    const datasetIds = new Set(datasets.map((dataset) => dataset.id))
    const sessions = normalizeSessions(parsed.sessions, defaultDatasetId).map((session) =>
      datasetIds.has(session.dataset_id) ? session : { ...session, dataset_id: defaultDatasetId },
    )
    const goals = normalizeGoals(parsed.goals).map((goal) =>
      goal.datasetId !== "all" && !datasetIds.has(goal.datasetId)
        ? { ...goal, datasetId: "all" }
        : goal,
    )
    const statsViews = normalizeStatsViews(parsed.statsViews)
    return {
      version: 1,
      sessions,
      people: Array.isArray(parsed.people)
        ? parsed.people.map((person) => ({
            ...person,
            tagIds: Array.isArray(person.tagIds) ? person.tagIds : [],
            role: normalizeSalvationRole(person.role),
            time_spent_minutes: normalizeMinutesValue(person.time_spent_minutes),
            notes: typeof person.notes === "string" ? person.notes : "",
          }))
        : [],
      tags: Array.isArray(parsed.tags)
        ? [...parsed.tags].sort((a, b) => a.name.localeCompare(b.name))
        : [],
      datasets,
      goals,
      statsViews,
    }
  } catch {
    return buildEmptyData()
  }
}

const saveAppData = (data: AppData) => {
  if (typeof localStorage === "undefined") return
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
}

const downloadJson = (filename: string, data: unknown) => {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" })
  const link = document.createElement("a")
  link.href = URL.createObjectURL(blob)
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(link.href)
}

const normalizeImportedData = (raw: unknown): AppData => {
  if (!raw || typeof raw !== "object") return buildEmptyData()
  const data = raw as Partial<AppData> & { trackers?: unknown }
  const datasets = normalizeDatasets(data.datasets ?? data.trackers)
  const defaultDatasetId = getDefaultDatasetId(datasets)
  const datasetIds = new Set(datasets.map((dataset) => dataset.id))
  const sessions = normalizeSessions(data.sessions, defaultDatasetId).map((session) =>
    datasetIds.has(session.dataset_id) ? session : { ...session, dataset_id: defaultDatasetId },
  )
  const people = Array.isArray(data.people)
    ? data.people.map((person) => ({
        ...person,
        tagIds: Array.isArray(person.tagIds) ? person.tagIds : [],
        role: normalizeSalvationRole(person.role),
        time_spent_minutes: normalizeMinutesValue(person.time_spent_minutes),
        notes: typeof person.notes === "string" ? person.notes : "",
      }))
    : []
  const tags = Array.isArray(data.tags) ? data.tags : []
  const goals = normalizeGoals(data.goals).map((goal) =>
    goal.datasetId !== "all" && !datasetIds.has(goal.datasetId)
      ? { ...goal, datasetId: "all" }
      : goal,
  )
  const statsViews = normalizeStatsViews(data.statsViews)
  return {
    version: 1,
    sessions,
    people,
    tags,
    datasets,
    goals,
    statsViews,
  }
}

const App = () => {
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const backupCheckInputRef = useRef<HTMLInputElement | null>(null)
  const eventTagInputRef = useRef<HTMLInputElement | null>(null)
  const sessionTagInputRef = useRef<HTMLInputElement | null>(null)
  const standaloneTagInputRef = useRef<HTMLInputElement | null>(null)
  const logEventTagInputRef = useRef<HTMLInputElement | null>(null)
  const logPersonTagInputRef = useRef<HTMLInputElement | null>(null)
  const logAddTagInputRef = useRef<HTMLInputElement | null>(null)
  const settingsDetailRef = useRef<HTMLDivElement | null>(null)
  const initialData = useMemo(() => loadAppData(), [])
  const [view, setView] = useState<"home" | "log" | "goals" | "stats">("home")
  const [homePanel, setHomePanel] = useState<"session" | "standalone" | null>(null)

  const [sessions, setSessions] = useState<SessionRow[]>(initialData.sessions)
  const [people, setPeople] = useState<PersonRow[]>(initialData.people)
  const [tags, setTags] = useState<TagRow[]>(initialData.tags)
  const [datasets, setDatasets] = useState<DatasetRow[]>(initialData.datasets)
  const [goals, setGoals] = useState<GoalDefinition[]>(initialData.goals)
  const [statsViews, setStatsViews] = useState<StatsView[]>(initialData.statsViews)
  const [loading, setLoading] = useState(false)
  const [actionMessage, setActionMessage] = useState("")
  const [actionError, setActionError] = useState("")
  const [authMode, setAuthMode] = useState<AuthViewMode>("signin")
  const [authEmail, setAuthEmail] = useState("")
  const [authPassword, setAuthPassword] = useState("")
  const [authUserId, setAuthUserId] = useState("")
  const [authUserEmail, setAuthUserEmail] = useState("")
  const [authLoading, setAuthLoading] = useState(false)
  const [syncLastUploadedAt, setSyncLastUploadedAt] = useState("")
  const [syncLastDownloadedAt, setSyncLastDownloadedAt] = useState("")
  const [syncLoading, setSyncLoading] = useState(false)
  const [autoSyncing, setAutoSyncing] = useState(false)
  const [autoSyncError, setAutoSyncError] = useState("")
  const [rollbackSavedAt, setRollbackSavedAt] = useState("")
  const [appVersion, setAppVersion] = useState(window.soulwinning?.version ?? "0.0.0")
  const [updateStatus, setUpdateStatus] = useState("")
  const [updateReady, setUpdateReady] = useState(false)
  const [updateProgress, setUpdateProgress] = useState<number | null>(null)
  const [userDataPath, setUserDataPath] = useState("")
  const [backupCheckResult, setBackupCheckResult] = useState<{
    ok: boolean
    message: string
    summary?: string
  } | null>(null)
  const [feedbackCategory, setFeedbackCategory] = useState<FeedbackCategory>("general")
  const [feedbackMessage, setFeedbackMessage] = useState("")
  const [feedbackSubmitting, setFeedbackSubmitting] = useState(false)
  const [statsDetail, setStatsDetail] = useState<StatsDetail | null>(null)
  const [isOffline, setIsOffline] = useState(() => 
    typeof navigator !== "undefined" ? !navigator.onLine : false
  )
  const cloudSyncInFlightRef = useRef(false)
  const cloudSyncRetryQueuedRef = useRef(false)
  const lastCloudUploadSignatureRef = useRef("")
  const cloudSyncBaselineUserIdRef = useRef("")
  const hasAutoPulledUserIdRef = useRef("")

  const defaultDatasetId = useMemo(() => datasets[0]?.id ?? "", [datasets])

  const [draftSessionName, setDraftSessionName] = useState("")
  const [draftDate, setDraftDate] = useState(() => toDateKey(new Date()))
  const [draftSavedCount, setDraftSavedCount] = useState("0")
  const [draftStartTime, setDraftStartTime] = useState("")
  const [draftEndTime, setDraftEndTime] = useState("")
  const [draftDoors, setDraftDoors] = useState("")
  const [draftSessionNotes, setDraftSessionNotes] = useState("")
  const [draftDatasetId, setDraftDatasetId] = useState(() => initialData.datasets[0]?.id ?? "")
  const [draftEventTagIds, setDraftEventTagIds] = useState<string[]>([])
  const [eventTagQuery, setEventTagQuery] = useState("")
  const [sessionPeople, setSessionPeople] = useState<DraftPerson[]>([])
  const [sessionSalvationsOpen, setSessionSalvationsOpen] = useState(false)
  const [editingSessionPersonId, setEditingSessionPersonId] = useState<string | null>(null)
  const [sessionPersonName, setSessionPersonName] = useState("")
  const [sessionPersonTags, setSessionPersonTags] = useState<string[]>([])
  const [sessionPersonRole, setSessionPersonRole] = useState<SalvationRole>(
    DEFAULT_SALVATION_ROLE,
  )
  const [sessionPersonNotes, setSessionPersonNotes] = useState("")
  const [sessionTagQuery, setSessionTagQuery] = useState("")
  const [sessionPersonNotice, setSessionPersonNotice] = useState("")
  const [sessionPersonHighlightId, setSessionPersonHighlightId] = useState<string | null>(null)

  const [standaloneDate, setStandaloneDate] = useState(() => toDateKey(new Date()))
  const [standaloneDatasetId, setStandaloneDatasetId] = useState(
    () => initialData.datasets[0]?.id ?? "",
  )
  const [standalonePersonName, setStandalonePersonName] = useState("")
  const [standalonePersonTags, setStandalonePersonTags] = useState<string[]>([])
  const [standalonePersonRole, setStandalonePersonRole] = useState<SalvationRole>(
    DEFAULT_SALVATION_ROLE,
  )
  const [standaloneTimeSpent, setStandaloneTimeSpent] = useState("")
  const [standalonePersonNotes, setStandalonePersonNotes] = useState("")
  const [standaloneTagQuery, setStandaloneTagQuery] = useState("")

  const [tagName, setTagName] = useState("")
  const [tagColor, setTagColor] = useState(TAG_COLORS[0])
  const [editingTagId, setEditingTagId] = useState<string | null>(null)
  const [editTagName, setEditTagName] = useState("")
  const [editTagColor, setEditTagColor] = useState(TAG_COLORS[0])
  const [datasetName, setDatasetName] = useState("")
  const [editingDatasetId, setEditingDatasetId] = useState<string | null>(null)
  const [datasetRename, setDatasetRename] = useState("")
  const [goalName, setGoalName] = useState("")
  const [goalMetric, setGoalMetric] = useState<GoalMetric>("sessions")
  const [goalPeriod, setGoalPeriod] = useState<GoalPeriod>("week")
  const [goalTarget, setGoalTarget] = useState("0")
  const [goalDatasetId, setGoalDatasetId] = useState<"all" | string>("all")
  const [goalTagIds, setGoalTagIds] = useState<string[]>([])
  const [goalTagMatchMode, setGoalTagMatchMode] = useState<TagMatchMode>("any")
  const [goalTagQuery, setGoalTagQuery] = useState("")
  const [goalStartDate, setGoalStartDate] = useState("")
  const [goalEndDate, setGoalEndDate] = useState("")
  const [editingGoalId, setEditingGoalId] = useState<string | null>(null)

  const [rangeStart, setRangeStart] = useState("")
  const [rangeEnd, setRangeEnd] = useState("")
  const [groupBy, setGroupBy] = useState<GroupBy>("month")
  const [selectedTagIds, setSelectedTagIds] = useState<string[]>([])
  const [selectedDatasetIds, setSelectedDatasetIds] = useState<string[]>([])
  const [tagMatchMode, setTagMatchMode] = useState<TagMatchMode>("any")
  const [weekStartsOn, setWeekStartsOn] = useState<0 | 1>(0)
  const [salvationQuery, setSalvationQuery] = useState("")
  const [logFilter, setLogFilter] = useState<"both" | "sessions" | "salvations">("both")
  const [logQuery, setLogQuery] = useState("")
  const [logDatasetIds, setLogDatasetIds] = useState<string[]>([])
  const [editingLogSessionId, setEditingLogSessionId] = useState<string | null>(null)
  const [editEventName, setEditEventName] = useState("")
  const [editEventDate, setEditEventDate] = useState("")
  const [editEventDatasetId, setEditEventDatasetId] = useState("")
  const [editEventSavedCount, setEditEventSavedCount] = useState("")
  const [editEventDoors, setEditEventDoors] = useState("")
  const [editEventStartTime, setEditEventStartTime] = useState("")
  const [editEventEndTime, setEditEventEndTime] = useState("")
  const [editEventNotes, setEditEventNotes] = useState("")
  const [editEventTagIds, setEditEventTagIds] = useState<string[]>([])
  const [editEventTagQuery, setEditEventTagQuery] = useState("")
  const [editingLogPersonId, setEditingLogPersonId] = useState<string | null>(null)
  const [logPersonName, setLogPersonName] = useState("")
  const [logPersonTags, setLogPersonTags] = useState<string[]>([])
  const [logPersonNotes, setLogPersonNotes] = useState("")
  const [logPersonTagQuery, setLogPersonTagQuery] = useState("")
  const [logPersonRole, setLogPersonRole] = useState<SalvationRole>(DEFAULT_SALVATION_ROLE)
  const [logPersonTimeSpent, setLogPersonTimeSpent] = useState("")
  const [logAddSessionId, setLogAddSessionId] = useState<string | null>(null)
  const [logAddPersonName, setLogAddPersonName] = useState("")
  const [logAddPersonTags, setLogAddPersonTags] = useState<string[]>([])
  const [logAddPersonRole, setLogAddPersonRole] = useState<SalvationRole>(DEFAULT_SALVATION_ROLE)
  const [logAddPersonNotes, setLogAddPersonNotes] = useState("")
  const [logAddPersonTagQuery, setLogAddPersonTagQuery] = useState("")
  const [statsQuery, setStatsQuery] = useState("")
  const [activeStatsViewId, setActiveStatsViewId] = useState(
    () => initialData.statsViews[0]?.id ?? DEFAULT_STATS_VIEWS[0]?.id ?? "",
  )
  const [statsViewName, setStatsViewName] = useState("")
  const [statsViewSections, setStatsViewSections] = useState<StatsSection[]>([])
  const [statsBuilderOpen, setStatsBuilderOpen] = useState(false)
  const [editingStatsViewId, setEditingStatsViewId] = useState<string | null>(null)
  const [themeMode, setThemeMode] = useState<"light" | "dark">(() => getStoredThemeMode())
  const [themeName, setThemeName] = useState(() => getStoredThemeName())
  const [settingsDetail, setSettingsDetail] = useState<
    "sessions" | "salvations" | "tags" | "datasets" | null
  >(null)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [settingsSection, setSettingsSection] = useState<
    | "overview"
    | "summary"
    | "appearance"
    | "tags"
    | "datasets"
    | "feedback"
    | "howto"
    | "exports"
    | "backup"
    | "sync"
    | "update"
    | "danger"
  >("overview")

  const clearMessages = () => {
    setActionMessage("")
    setActionError("")
  }

  const openHomePanel = (panel: "session" | "standalone") => {
    setHomePanel(panel)
    if (panel === "session") {
      setSessionSalvationsOpen(false)
    }
  }

  const closeHomePanel = () => {
    setHomePanel(null)
    setSessionSalvationsOpen(false)
  }

  const resetSessionDraft = useCallback(() => {
    setDraftSessionName("")
    setDraftDate(toDateKey(new Date()))
    setDraftSavedCount("0")
    setDraftStartTime("")
    setDraftEndTime("")
    setDraftDoors("")
    setDraftSessionNotes("")
    setDraftDatasetId(defaultDatasetId)
    setDraftEventTagIds([])
    setEventTagQuery("")
    setSessionPeople([])
    setSessionSalvationsOpen(false)
    setEditingSessionPersonId(null)
    setSessionPersonNotice("")
    setSessionPersonHighlightId(null)
    setSessionPersonName("")
    setSessionPersonTags([])
    setSessionPersonRole(DEFAULT_SALVATION_ROLE)
    setSessionPersonNotes("")
    setSessionTagQuery("")
  }, [defaultDatasetId])

  const resetStandaloneDraft = useCallback(() => {
    setStandaloneDate(toDateKey(new Date()))
    setStandaloneDatasetId(defaultDatasetId)
    setStandalonePersonName("")
    setStandalonePersonTags([])
    setStandalonePersonRole(DEFAULT_SALVATION_ROLE)
    setStandaloneTimeSpent("")
    setStandalonePersonNotes("")
    setStandaloneTagQuery("")
  }, [defaultDatasetId])

  const resetGoalForm = () => {
    setGoalName("")
    setGoalMetric("sessions")
    setGoalPeriod("week")
    setGoalTarget("0")
    setGoalDatasetId("all")
    setGoalTagIds([])
    setGoalTagMatchMode("any")
    setGoalTagQuery("")
    setGoalStartDate("")
    setGoalEndDate("")
    setEditingGoalId(null)
    setGoalStartDate("")
    setGoalEndDate("")
    setEditingGoalId(null)
  }

  const clearGoals = () => {
    clearMessages()
    setGoals([])
    resetGoalForm()
    setActionMessage("Goals cleared.")
  }

  const toggleGoalTag = (tagId: string) => {
    setGoalTagIds((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const handleSaveGoal = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    const target = normalizeGoalTarget(goalTarget)
    if (target <= 0) {
      setActionError("Goal target must be greater than 0.")
      return
    }
    const customStart = goalStartDate.trim()
    const customEnd = goalEndDate.trim()
    if (goalPeriod === "custom") {
      if (!customStart || !customEnd) {
        setActionError("Custom range requires a start and end date.")
        return
      }
      if (customEnd < customStart) {
        setActionError("Custom end date must be on or after the start date.")
        return
      }
    }
    const existingGoal = editingGoalId ? goals.find((goal) => goal.id === editingGoalId) : null
    const newGoal: GoalDefinition = {
      id: editingGoalId ?? createLocalId(),
      name: goalName.trim(),
      metric: goalMetric,
      period: goalPeriod,
      target,
      datasetId: goalDatasetId,
      tagIds: goalMetric === "salvations" ? goalTagIds : [],
      tagMatchMode: goalTagMatchMode,
      startDate: goalPeriod === "custom" ? customStart : undefined,
      endDate: goalPeriod === "custom" ? customEnd : undefined,
      created_at: existingGoal?.created_at ?? new Date().toISOString(),
    }
    if (editingGoalId) {
      setGoals((prev) => prev.map((goal) => (goal.id === editingGoalId ? newGoal : goal)))
      setActionMessage("Goal updated.")
    } else {
      setGoals((prev) => [...prev, newGoal])
      setActionMessage("Goal added.")
    }
    resetGoalForm()
  }

  const handleDeleteGoal = (goalId: string) => {
    if (!window.confirm("Delete this goal?")) return
    clearMessages()
    setGoals((prev) => prev.filter((goal) => goal.id !== goalId))
    if (editingGoalId === goalId) {
      resetGoalForm()
    }
    setActionMessage("Goal deleted.")
  }

  const handleEditGoal = (goalId: string) => {
    const goal = goals.find((item) => item.id === goalId)
    if (!goal) return
    clearMessages()
    setEditingGoalId(goal.id)
    setGoalName(goal.name)
    setGoalMetric(goal.metric)
    setGoalPeriod(goal.period)
    setGoalTarget(String(goal.target))
    setGoalDatasetId(goal.datasetId)
    setGoalTagIds(goal.tagIds)
    setGoalTagMatchMode(goal.tagMatchMode)
    setGoalTagQuery("")
    setGoalStartDate(goal.startDate ?? "")
    setGoalEndDate(goal.endDate ?? "")
  }

  const handleCancelGoalEdit = () => {
    clearMessages()
    resetGoalForm()
  }

  const handleCheckForUpdates = () => {
    setUpdateReady(false)
    setUpdateProgress(null)
    setUpdateStatus("Checking for updates...")
    const api = window.soulwinning
    if (!api?.checkForUpdates) {
      setUpdateStatus("Updates are available in the desktop app.")
      return
    }
    api.checkForUpdates().catch(() => {
      setUpdateStatus("Update check failed.")
    })
  }

  const handleInstallUpdate = () => {
    const api = window.soulwinning
    if (!api?.installUpdate) {
      setUpdateStatus("Updates are available in the desktop app.")
      return
    }
    api.installUpdate()
  }

  const handleSubmitFeedback = async (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()

    const message = feedbackMessage.trim()
    if (message.length < 4) {
      setActionError("Please enter at least 4 characters of feedback.")
      return
    }
    if (!supabase || !isSupabaseConfigured) {
      setActionError("Feedback is not configured for this deployment.")
      return
    }

    setFeedbackSubmitting(true)
    try {
      const { error } = await supabase.from(SUPABASE_FEEDBACK_TABLE).insert({
        category: feedbackCategory,
        message,
        app_version: appVersion || null,
        page_url: typeof window === "undefined" ? null : window.location.href,
        user_agent: typeof navigator === "undefined" ? null : navigator.userAgent,
        user_id: authUserId || null,
      })
      if (error) throw error
      setFeedbackMessage("")
      setFeedbackCategory("general")
      setActionMessage("Thanks. Feedback sent.")
    } catch (error) {
      setActionError(error instanceof Error ? error.message : "Could not send feedback.")
    } finally {
      setFeedbackSubmitting(false)
    }
  }

  const handleOpenSupport = () => {
    const api = window.soulwinning
    if (api?.openExternal) {
      api.openExternal(SUPPORT_URL)
      return
    }
    window.open(SUPPORT_URL, "_blank", "noopener")
  }

  const handleBackupCheckClick = () => {
    backupCheckInputRef.current?.click()
  }

  const handleBackupCheckFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const input = event.target
    const file = input.files?.[0]
    if (!file) return
    setBackupCheckResult(null)
    try {
      const text = await file.text()
      const parsed = JSON.parse(text)
      const normalized = normalizeImportedData(parsed)
      if (!normalized.sessions.length && !normalized.people.length && !normalized.tags.length) {
        setBackupCheckResult({
          ok: false,
          message: "No data found in that file.",
        })
        return
      }
      const summary = [
        `Events: ${numberFormatter.format(normalized.sessions.length)}`,
        `Salvations: ${numberFormatter.format(normalized.people.length)}`,
        `Tags: ${numberFormatter.format(normalized.tags.length)}`,
        `Data sets: ${numberFormatter.format(normalized.datasets.length)}`,
        `Goals: ${numberFormatter.format(normalized.goals.length)}`,
        `Views: ${numberFormatter.format(normalized.statsViews.length)}`,
      ].join(" · ")
      setBackupCheckResult({
        ok: true,
        message: "Backup looks valid.",
        summary,
      })
    } catch {
      setBackupCheckResult({
        ok: false,
        message: "Backup file could not be read.",
      })
    } finally {
      input.value = ""
    }
  }

  const handleStatsSearchKeyDown = (
    event: React.KeyboardEvent,
    action: () => void,
  ) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault()
      action()
    }
  }

  const handleOpenStatsEvent = (session: SessionRow) => {
    setStatsDetail({ kind: "event", session })
  }

  const handleOpenStatsSalvation = (person: PersonRow, session?: SessionRow) => {
    setStatsDetail({ kind: "salvation", person, session })
  }

  const handleViewStatsDetailInLog = () => {
    if (!statsDetail) return
    if (statsDetail.kind === "event") {
      const label = statsDetail.session.name || statsDetail.session.date
      setLogFilter("sessions")
      setLogQuery(label)
    } else {
      const personName = statsDetail.person.name || ""
      const fallback = statsDetail.session?.name || statsDetail.session?.date || ""
      setLogFilter("salvations")
      setLogQuery(personName || fallback)
    }
    setView("log")
    setStatsDetail(null)
  }

  const createTag = (name: string, color?: string) => {
    clearMessages()
    const trimmed = name.trim()
    if (!trimmed) {
      setActionError("Tag name is required.")
      return null
    }
    if (tags.some((tag) => tag.name.toLowerCase() === trimmed.toLowerCase())) {
      setActionError("That tag already exists.")
      return null
    }
    const newTag: TagRow = {
      id: createLocalId(),
      name: trimmed,
      color: color ?? TAG_COLORS[0],
      created_at: new Date().toISOString(),
    }
    setTags((prev) => [...prev, newTag].sort((a, b) => a.name.localeCompare(b.name)))
    setActionMessage("Tag added.")
    return newTag
  }

  const appDataSnapshot = useMemo<AppData>(
    () => ({
      version: 1,
      sessions,
      people,
      tags,
      datasets,
      goals,
      statsViews,
    }),
    [sessions, people, tags, datasets, goals, statsViews],
  )

  useEffect(() => {
    saveAppData(appDataSnapshot)
  }, [appDataSnapshot])

  useEffect(() => {
    const rollback = loadRollbackSnapshot()
    setRollbackSavedAt(rollback?.saved_at ?? "")
  }, [])

  useEffect(() => {
    if (!supabase) return
    let isMounted = true
    supabase.auth
      .getSession()
      .then(({ data }) => {
        if (!isMounted) return
        const session = data.session
        setAuthUserId(session?.user.id ?? "")
        setAuthUserEmail(session?.user.email ?? "")
      })
      .catch(() => {
        if (!isMounted) return
        setAuthUserId("")
        setAuthUserEmail("")
      })

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!isMounted) return
      setAuthUserId(session?.user.id ?? "")
      setAuthUserEmail(session?.user.email ?? "")
      if (!session) {
        setSyncLastUploadedAt("")
        setSyncLastDownloadedAt("")
        setAutoSyncError("")
        setAutoSyncing(false)
        cloudSyncInFlightRef.current = false
        cloudSyncRetryQueuedRef.current = false
        lastCloudUploadSignatureRef.current = ""
        cloudSyncBaselineUserIdRef.current = ""
        hasAutoPulledUserIdRef.current = ""
      }
    })

    return () => {
      isMounted = false
      subscription.unsubscribe()
    }
  }, [])

  useEffect(() => {
    if (!authUserId) {
      cloudSyncBaselineUserIdRef.current = ""
      cloudSyncRetryQueuedRef.current = false
      lastCloudUploadSignatureRef.current = ""
      hasAutoPulledUserIdRef.current = ""
      setAutoSyncError("")
      setAutoSyncing(false)
      return
    }

    if (cloudSyncBaselineUserIdRef.current !== authUserId) {
      cloudSyncBaselineUserIdRef.current = authUserId
      cloudSyncRetryQueuedRef.current = false
      const syncState = loadUserSyncState(authUserId)
      lastCloudUploadSignatureRef.current = syncState?.signature ?? ""
      setSyncLastUploadedAt(syncState?.uploaded_at ?? "")
      setSyncLastDownloadedAt(syncState?.downloaded_at ?? "")
      setAutoSyncError("")
    }
  }, [authUserId])

  useEffect(() => {
    const isValidTheme = AVAILABLE_THEMES.some((theme) => theme.id === themeName)
    const fallbackTheme =
      AVAILABLE_THEMES.find((theme) => theme.id === DEFAULT_THEME_ID) ??
      AVAILABLE_THEMES[0] ??
      THEME_DEFS[0]
    const themeDef =
      (isValidTheme ? AVAILABLE_THEMES.find((theme) => theme.id === themeName) : null) ??
      fallbackTheme
    const vars = themeMode === "dark" ? themeDef.dark : themeDef.light
    const root = document.documentElement
    const resolve = (key: string, fallback: string) => vars[key] ?? fallback

    root.dataset.theme = themeMode
    root.dataset.themeName = themeDef.id

    root.style.setProperty("--theme-bg", resolve("--bg", "#f6efe3"))
    root.style.setProperty("--theme-bg-2", resolve("--bg-2", resolve("--bg", "#f6efe3")))
    root.style.setProperty(
      "--theme-panel",
      resolve("--panel-bg", resolve("--panel-solid", "rgba(255, 255, 255, 0.86)")),
    )
    root.style.setProperty(
      "--theme-panel-solid",
      resolve("--panel-solid", resolve("--panel-bg", "#ffffff")),
    )
    root.style.setProperty("--theme-text", resolve("--text", "#201812"))
    root.style.setProperty("--theme-muted", resolve("--muted", "#6f6256"))
    root.style.setProperty("--theme-border", resolve("--border", "rgba(51, 35, 24, 0.14)"))
    root.style.setProperty(
      "--theme-border-strong",
      resolve("--border-strong", resolve("--border", "rgba(51, 35, 24, 0.22)")),
    )
    root.style.setProperty("--theme-accent", resolve("--accent", "#c56b3c"))
    root.style.setProperty("--theme-accent-2", resolve("--accent-2", resolve("--accent", "#c56b3c")))
    root.style.setProperty("--theme-shadow", resolve("--shadow", "0 18px 50px rgba(63, 44, 30, 0.18)"))
    root.style.setProperty("--theme-toggle-bg", resolve("--toggle-bg", "rgba(32, 24, 18, 0.2)"))
    root.style.setProperty("--theme-toggle-knob", resolve("--toggle-knob", "#ffffff"))
    root.style.setProperty("--theme-nav-hover", resolve("--nav-hover", "rgba(32, 24, 18, 0.08)"))
    root.style.setProperty("--theme-nav-active", resolve("--nav-active", "rgba(32, 24, 18, 0.14)"))
    root.style.setProperty(
      "--theme-input-bg",
      resolve("--input-bg", resolve("--panel-bg", "rgba(255, 255, 255, 0.82)")),
    )
    root.style.setProperty("--radius-lg", resolve("--radius-l", "24px"))
    root.style.setProperty("--radius-md", resolve("--radius-m", "16px"))
    root.style.setProperty("--radius-sm", resolve("--radius-s", "12px"))

    if (vars["--font-sans"]) {
      root.style.setProperty("--theme-font-sans", vars["--font-sans"])
    } else {
      root.style.removeProperty("--theme-font-sans")
    }

    if (!isValidTheme && themeDef && themeDef.id !== themeName) {
      setThemeName(themeDef.id)
    }

    try {
      localStorage.setItem(THEME_MODE_KEY, themeMode)
      localStorage.setItem(THEME_NAME_KEY, themeDef.id)
    } catch {
      // ignore
    }
  }, [themeMode, themeName])

  useEffect(() => {
    if (!actionMessage) return
    const timer = window.setTimeout(() => setActionMessage(""), 3200)
    return () => window.clearTimeout(timer)
  }, [actionMessage])

  useEffect(() => {
    const api = window.soulwinning
    if (!api?.getUserDataPath) return
    api
      .getUserDataPath()
      .then((path) => {
        if (typeof path === "string") {
          setUserDataPath(path)
        }
      })
      .catch(() => {
        // ignore
      })
  }, [])

  useEffect(() => {
    const api = window.soulwinning
    if (!api?.getVersion) return
    api
      .getVersion()
      .then((version) => {
        if (typeof version === "string" && version.trim()) {
          setAppVersion(version)
        }
      })
      .catch(() => {
        // ignore
      })
  }, [])

  useEffect(() => {
    const api = window.soulwinning
    if (!api?.onUpdateStatus) return
    const unsubscribe = api.onUpdateStatus((payload) => {
      if (!payload) return
      if (payload.message) setUpdateStatus(payload.message)
      if (typeof payload.progress === "number" && Number.isFinite(payload.progress)) {
        const clamped = Math.min(100, Math.max(0, payload.progress))
        setUpdateProgress(clamped)
      } else if (!payload.ready) {
        setUpdateProgress(null)
      }
      setUpdateReady(Boolean(payload.ready))
    })
    return () => {
      if (typeof unsubscribe === "function") {
        unsubscribe()
      }
    }
  }, [])

  useEffect(() => {
    if (!sessionPersonNotice) return
    const timer = window.setTimeout(() => setSessionPersonNotice(""), 2200)
    return () => window.clearTimeout(timer)
  }, [sessionPersonNotice])

  useEffect(() => {
    if (!sessionPersonHighlightId) return
    const timer = window.setTimeout(() => setSessionPersonHighlightId(null), 2200)
    return () => window.clearTimeout(timer)
  }, [sessionPersonHighlightId])

  useEffect(() => {
    if (settingsDetail && settingsDetailRef.current) {
      settingsDetailRef.current.scrollIntoView({ behavior: "smooth", block: "start" })
    }
  }, [settingsDetail])

  useEffect(() => {
    if (!settingsOpen) {
      setSettingsSection("overview")
      setSettingsDetail(null)
    }
  }, [settingsOpen])

  useEffect(() => {
    if (settingsSection !== "summary" && settingsDetail) {
      setSettingsDetail(null)
    }
  }, [settingsDetail, settingsSection])

  useEffect(() => {
    if (view !== "home" && homePanel) {
      setHomePanel(null)
      setSessionSalvationsOpen(false)
    }
  }, [view, homePanel])

  useEffect(() => {
    if (view !== "stats" && statsDetail) {
      setStatsDetail(null)
    }
  }, [view, statsDetail])

  useEffect(() => {
    if (!statsViews.length) {
      setStatsViews([...DEFAULT_STATS_VIEWS])
      setActiveStatsViewId(DEFAULT_STATS_VIEWS[0]?.id ?? "")
      return
    }
    if (!statsViews.some((view) => view.id === activeStatsViewId)) {
      setActiveStatsViewId(statsViews[0]?.id ?? "")
    }
  }, [statsViews, activeStatsViewId])

  useEffect(() => {
    if (!defaultDatasetId) return
    if (!datasets.some((dataset) => dataset.id === draftDatasetId)) {
      setDraftDatasetId(defaultDatasetId)
    }
    if (!datasets.some((dataset) => dataset.id === standaloneDatasetId)) {
      setStandaloneDatasetId(defaultDatasetId)
    }
    if (
      editingLogSessionId &&
      (!editEventDatasetId || !datasets.some((dataset) => dataset.id === editEventDatasetId))
    ) {
      setEditEventDatasetId(defaultDatasetId)
    }
    setSelectedDatasetIds((prev) =>
      prev.filter((datasetId) => datasets.some((dataset) => dataset.id === datasetId)),
    )
    setLogDatasetIds((prev) =>
      prev.filter((datasetId) => datasets.some((dataset) => dataset.id === datasetId)),
    )
    if (goalDatasetId !== "all" && !datasets.some((dataset) => dataset.id === goalDatasetId)) {
      setGoalDatasetId("all")
    }
  }, [
    defaultDatasetId,
    draftDatasetId,
    standaloneDatasetId,
    datasets,
    goalDatasetId,
    editingLogSessionId,
    editEventDatasetId,
  ])

  useEffect(() => {
    if (goalMetric === "salvations") return
    if (goalTagIds.length) {
      setGoalTagIds([])
    }
    if (goalTagQuery) {
      setGoalTagQuery("")
    }
  }, [goalMetric, goalTagIds, goalTagQuery])

  useEffect(() => {
    if (goalPeriod === "custom") return
    if (goalStartDate || goalEndDate) {
      setGoalStartDate("")
      setGoalEndDate("")
    }
  }, [goalPeriod, goalStartDate, goalEndDate])

  const peopleWithTags = useMemo(() => {
    return people
  }, [people])

  const tagsById = useMemo(() => {
    const map = new Map<string, TagRow>()
    tags.forEach((tag) => map.set(tag.id, tag))
    return map
  }, [tags])

  const datasetById = useMemo(() => {
    const map = new Map<string, DatasetRow>()
    datasets.forEach((dataset) => map.set(dataset.id, dataset))
    return map
  }, [datasets])

  const sessionsById = useMemo(() => {
    const map = new Map<string, SessionRow>()
    sessions.forEach((session) => map.set(session.id, session))
    return map
  }, [sessions])

  const peopleBySession = useMemo(() => {
    const map = new Map<string, typeof peopleWithTags>()
    peopleWithTags.forEach((person) => {
      const list = map.get(person.session_id)
      if (list) {
        list.push(person)
      } else {
        map.set(person.session_id, [person])
      }
    })
    return map
  }, [peopleWithTags])

  const tagCounts = useMemo(() => {
    const counts = new Map<string, number>()
    peopleWithTags.forEach((person) => {
      person.tagIds.forEach((tagId) => {
        counts.set(tagId, (counts.get(tagId) ?? 0) + 1)
      })
    })
    sessions.forEach((session) => {
      session.tagIds.forEach((tagId) => {
        counts.set(tagId, (counts.get(tagId) ?? 0) + 1)
      })
    })
    return counts
  }, [peopleWithTags, sessions])

  const tagUsageMap = useMemo(() => {
    const map = new Map<string, TagUsageEntry[]>()
    peopleWithTags.forEach((person) => {
      const session = sessionsById.get(person.session_id)
      const dataset = session ? datasetById.get(session.dataset_id) : undefined
      person.tagIds.forEach((tagId) => {
        const list = map.get(tagId)
        const entry = { person, session, dataset }
        if (list) {
          list.push(entry)
        } else {
          map.set(tagId, [entry])
        }
      })
    })
    map.forEach((entries) => {
      entries.sort((a, b) => {
        const aDate = a.session?.date ?? ""
        const bDate = b.session?.date ?? ""
        return bDate.localeCompare(aDate)
      })
    })
    return map
  }, [peopleWithTags, sessionsById, datasetById])

  const tagsSortedByUsage = useMemo(() => {
    return [...tags].sort((a, b) => {
      const countDiff = (tagCounts.get(b.id) ?? 0) - (tagCounts.get(a.id) ?? 0)
      if (countDiff !== 0) return countDiff
      return a.name.localeCompare(b.name)
    })
  }, [tags, tagCounts])

  const datasetCounts = useMemo(() => {
    const counts = new Map<string, { sessions: number; salvations: number }>()
    datasets.forEach((dataset) => counts.set(dataset.id, { sessions: 0, salvations: 0 }))
    sessions.forEach((session) => {
      const entry = counts.get(session.dataset_id)
      if (entry) {
        entry.sessions += 1
        entry.salvations += session.saved_count ?? 0
      }
    })
    return counts
  }, [datasets, sessions])

  const eventTagOptions = useMemo(() => {
    const query = eventTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [eventTagQuery, tags, tagsSortedByUsage])

  const sessionTagOptions = useMemo(() => {
    const query = sessionTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [sessionTagQuery, tags, tagsSortedByUsage])

  const standaloneTagOptions = useMemo(() => {
    const query = standaloneTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [standaloneTagQuery, tags, tagsSortedByUsage])

  const logEventTagOptions = useMemo(() => {
    const query = editEventTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [editEventTagQuery, tags, tagsSortedByUsage])

  const logPersonTagOptions = useMemo(() => {
    const query = logPersonTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [logPersonTagQuery, tags, tagsSortedByUsage])

  const logAddTagOptions = useMemo(() => {
    const query = logAddPersonTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [logAddPersonTagQuery, tags, tagsSortedByUsage])

  const goalTagOptions = useMemo(() => {
    const query = goalTagQuery.trim().toLowerCase()
    if (!query) return tagsSortedByUsage.slice(0, TOP_TAG_LIMIT)
    return tags
      .filter((tag) => tag.name.toLowerCase().includes(query))
      .slice(0, TAG_SEARCH_LIMIT)
  }, [goalTagQuery, tags, tagsSortedByUsage])

  const sessionTagName = sessionTagQuery.trim()
  const eventTagName = eventTagQuery.trim()
  const standaloneTagName = standaloneTagQuery.trim()
  const logEventTagName = editEventTagQuery.trim()
  const logPersonTagName = logPersonTagQuery.trim()
  const logAddTagName = logAddPersonTagQuery.trim()
  const sessionTagMatch = sessionTagName
    ? tags.find((tag) => tag.name.toLowerCase() === sessionTagName.toLowerCase()) ?? null
    : null
  const eventTagMatch = eventTagName
    ? tags.find((tag) => tag.name.toLowerCase() === eventTagName.toLowerCase()) ?? null
    : null
  const standaloneTagMatch = standaloneTagName
    ? tags.find((tag) => tag.name.toLowerCase() === standaloneTagName.toLowerCase()) ?? null
    : null
  const logEventTagMatch = logEventTagName
    ? tags.find((tag) => tag.name.toLowerCase() === logEventTagName.toLowerCase()) ?? null
    : null
  const logPersonTagMatch = logPersonTagName
    ? tags.find((tag) => tag.name.toLowerCase() === logPersonTagName.toLowerCase()) ?? null
    : null
  const logAddTagMatch = logAddTagName
    ? tags.find((tag) => tag.name.toLowerCase() === logAddTagName.toLowerCase()) ?? null
    : null
  const sessionTagActionLabel = sessionTagName
    ? sessionTagMatch
      ? `+ Use tag "${sessionTagMatch.name}"`
      : `+ Add tag "${sessionTagName}"`
    : "+ Add tag"
  const eventTagActionLabel = eventTagName
    ? eventTagMatch
      ? `+ Use tag "${eventTagMatch.name}"`
      : `+ Add tag "${eventTagName}"`
    : "+ Add tag"
  const standaloneTagActionLabel = standaloneTagName
    ? standaloneTagMatch
      ? `+ Use tag "${standaloneTagMatch.name}"`
      : `+ Add tag "${standaloneTagName}"`
    : "+ Add tag"
  const logEventTagActionLabel = logEventTagName
    ? logEventTagMatch
      ? `+ Use tag "${logEventTagMatch.name}"`
      : `+ Add tag "${logEventTagName}"`
    : "+ Add tag"
  const logPersonTagActionLabel = logPersonTagName
    ? logPersonTagMatch
      ? `+ Use tag "${logPersonTagMatch.name}"`
      : `+ Add tag "${logPersonTagName}"`
    : "+ Add tag"
  const logAddTagActionLabel = logAddTagName
    ? logAddTagMatch
      ? `+ Use tag "${logAddTagMatch.name}"`
      : `+ Add tag "${logAddTagName}"`
    : "+ Add tag"

  const sessionsSorted = useMemo(() => {
    return [...sessions].sort((a, b) => b.date.localeCompare(a.date))
  }, [sessions])

  const salvationLogEntries = useMemo(() => {
    return peopleWithTags
      .map((person) => {
        const session = sessionsById.get(person.session_id)
        return {
          person,
          session,
          dataset: session ? datasetById.get(session.dataset_id) : undefined,
        }
      })
      .filter((entry) => Boolean(entry.session))
      .sort((a, b) => {
        const aDate = a.session?.date ?? ""
        const bDate = b.session?.date ?? ""
        if (aDate !== bDate) return bDate.localeCompare(aDate)
        const aCreated = a.person.created_at ?? ""
        const bCreated = b.person.created_at ?? ""
        if (aCreated !== bCreated) return bCreated.localeCompare(aCreated)
        return a.person.name.localeCompare(b.person.name)
      })
  }, [peopleWithTags, sessionsById, datasetById])

  const logSessions = useMemo(() => {
    const query = logQuery.trim().toLowerCase()
    return sessionsSorted.filter((session) => {
      if (logDatasetIds.length && !logDatasetIds.includes(session.dataset_id)) return false
      if (!query) return true
      const dataset = datasetById.get(session.dataset_id)
      const sessionLabel =
        session.name ||
        (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
      const matchesSession =
        sessionLabel.toLowerCase().includes(query) ||
        (session.notes ?? "").toLowerCase().includes(query) ||
        (dataset?.name ?? "").toLowerCase().includes(query)
      if (matchesSession) return true
      const peopleInSession = peopleBySession.get(session.id) ?? []
      return peopleInSession.some((person) => {
        if (person.name.toLowerCase().includes(query)) return true
        return person.notes ? person.notes.toLowerCase().includes(query) : false
      })
    })
  }, [sessionsSorted, logQuery, logDatasetIds, datasetById, peopleBySession])

  const logSalvations = useMemo(() => {
    const query = logQuery.trim().toLowerCase()
    return salvationLogEntries.filter(({ person, session, dataset }) => {
      if (!session) return false
      if (logDatasetIds.length && !logDatasetIds.includes(session.dataset_id)) return false
      if (!query) return true
      const sessionLabel =
        session.name ||
        (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
      return (
        person.name.toLowerCase().includes(query) ||
        (person.notes ? person.notes.toLowerCase().includes(query) : false) ||
        sessionLabel.toLowerCase().includes(query) ||
        (dataset?.name ?? "").toLowerCase().includes(query)
      )
    })
  }, [salvationLogEntries, logQuery, logDatasetIds])

  const logUnnamedSalvations = useMemo(() => {
    return logSessions
      .map((session) => {
        const peopleCount = peopleBySession.get(session.id)?.length ?? 0
        const unnamedCount = (session.saved_count ?? 0) - peopleCount
        if (unnamedCount <= 0) return null
        return {
          session,
          dataset: datasetById.get(session.dataset_id),
          unnamedCount,
        }
      })
      .filter(
        (
          entry,
        ): entry is { session: SessionRow; dataset: DatasetRow | undefined; unnamedCount: number } =>
          Boolean(entry),
      )
  }, [logSessions, peopleBySession, datasetById])

  const homeUnnamedSalvations = useMemo(() => {
    return sessionsSorted
      .map((session) => {
        const peopleCount = peopleBySession.get(session.id)?.length ?? 0
        const unnamedCount = (session.saved_count ?? 0) - peopleCount
        if (unnamedCount <= 0) return null
        return {
          session,
          dataset: datasetById.get(session.dataset_id),
          unnamedCount,
        }
      })
      .filter(
        (
          entry,
        ): entry is { session: SessionRow; dataset: DatasetRow | undefined; unnamedCount: number } =>
          Boolean(entry),
      )
  }, [sessionsSorted, peopleBySession, datasetById])

  const statsQueryValue = statsQuery.trim().toLowerCase()

  const sessionsInRange = useMemo(() => {
    return sessions.filter((session) => {
      if (!isWithinRange(session.date, rangeStart, rangeEnd)) return false
      if (selectedDatasetIds.length && !selectedDatasetIds.includes(session.dataset_id)) return false
      return true
    })
  }, [sessions, rangeStart, rangeEnd, selectedDatasetIds])

  const sessionIdsInRange = useMemo(() => {
    return new Set(sessionsInRange.map((session) => session.id))
  }, [sessionsInRange])

  const peopleInRange = useMemo(() => {
    return peopleWithTags.filter((person) => sessionIdsInRange.has(person.session_id))
  }, [peopleWithTags, sessionIdsInRange])

  const peopleMatchingQuery = useMemo(() => {
    if (!statsQueryValue) return peopleInRange
    return peopleInRange.filter((person) => {
      const session = sessionsById.get(person.session_id)
      const dataset = session ? datasetById.get(session.dataset_id) : undefined
      const sessionLabel = session
        ? session.name || (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
        : ""
      const personTagNames = person.tagIds
        .map((id) => tagsById.get(id)?.name ?? "")
        .join(" ")
      const eventTagNames = session?.tagIds
        ? session.tagIds.map((id) => tagsById.get(id)?.name ?? "").join(" ")
        : ""
      return (
        person.name.toLowerCase().includes(statsQueryValue) ||
        (person.notes ? person.notes.toLowerCase().includes(statsQueryValue) : false) ||
        sessionLabel.toLowerCase().includes(statsQueryValue) ||
        (session?.notes ? session.notes.toLowerCase().includes(statsQueryValue) : false) ||
        (dataset?.name ?? "").toLowerCase().includes(statsQueryValue) ||
        personTagNames.toLowerCase().includes(statsQueryValue) ||
        eventTagNames.toLowerCase().includes(statsQueryValue)
      )
    })
  }, [peopleInRange, statsQueryValue, sessionsById, datasetById, tagsById])

  const peopleMatchingTags = useMemo(() => {
    if (!selectedTagIds.length) return peopleMatchingQuery
    return peopleMatchingQuery.filter((person) => {
      const session = sessionsById.get(person.session_id)
      if (session?.tagIds?.length) {
        const matchesEventTags =
          tagMatchMode === "all"
            ? selectedTagIds.every((tagId) => session.tagIds.includes(tagId))
            : selectedTagIds.some((tagId) => session.tagIds.includes(tagId))
        if (matchesEventTags) return true
      }
      if (tagMatchMode === "all") {
        return selectedTagIds.every((tagId) => person.tagIds.includes(tagId))
      }
      return selectedTagIds.some((tagId) => person.tagIds.includes(tagId))
    })
  }, [peopleMatchingQuery, selectedTagIds, tagMatchMode, sessionsById])

  const sessionsMatchingTags = useMemo(() => {
    const sessionIds = new Set(peopleMatchingTags.map((person) => person.session_id))
    const query = statsQueryValue
    return sessionsInRange.filter((session) => {
      const dataset = datasetById.get(session.dataset_id)
      const sessionLabel =
        session.name || (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
      const eventTagNames = session.tagIds
        .map((id) => tagsById.get(id)?.name ?? "")
        .join(" ")
      const matchesQuery =
        !query ||
        sessionLabel.toLowerCase().includes(query) ||
        (session.notes ? session.notes.toLowerCase().includes(query) : false) ||
        (dataset?.name ?? "").toLowerCase().includes(query) ||
        eventTagNames.toLowerCase().includes(query)

      if (!selectedTagIds.length) {
        return matchesQuery || sessionIds.has(session.id)
      }
      if (sessionIds.has(session.id)) return true
      if (!matchesQuery) return false
      if (!session.tagIds.length) return false
      if (tagMatchMode === "all") {
        return selectedTagIds.every((tagId) => session.tagIds.includes(tagId))
      }
      return selectedTagIds.some((tagId) => session.tagIds.includes(tagId))
    })
  }, [sessionsInRange, selectedTagIds, peopleMatchingTags, tagMatchMode, statsQueryValue, datasetById, tagsById])

  const statsSearchResults = useMemo(() => {
    if (!statsQueryValue) {
      return {
        events: [],
        salvations: [],
        eventPreview: [],
        salvationPreview: [],
        eventOverflow: 0,
        salvationOverflow: 0,
      }
    }

    const events = sessionsMatchingTags
      .map((session) => ({
        session,
        dataset: datasetById.get(session.dataset_id),
        label:
          session.name || (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event"),
      }))
      .sort((a, b) => b.session.date.localeCompare(a.session.date))

    const salvations = peopleMatchingTags
      .map((person) => {
        const session = sessionsById.get(person.session_id)
        const sessionLabel = session
          ? session.name ||
            (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
          : "Unknown event"
        return {
          person,
          session,
          dataset: session ? datasetById.get(session.dataset_id) : undefined,
          sessionLabel,
        }
      })
      .sort((a, b) => {
        const aDate = a.session?.date ?? ""
        const bDate = b.session?.date ?? ""
        if (aDate !== bDate) return bDate.localeCompare(aDate)
        return a.person.name.localeCompare(b.person.name)
      })

    const eventPreview = events.slice(0, STATS_SEARCH_LIMIT)
    const salvationPreview = salvations.slice(0, STATS_SEARCH_LIMIT)
    return {
      events,
      salvations,
      eventPreview,
      salvationPreview,
      eventOverflow: Math.max(0, events.length - eventPreview.length),
      salvationOverflow: Math.max(0, salvations.length - salvationPreview.length),
    }
  }, [statsQueryValue, sessionsMatchingTags, peopleMatchingTags, sessionsById, datasetById])

  const totals = useMemo(() => {
    const totalSessions = sessionsMatchingTags.length
    const totalSavedReported = sessionsMatchingTags.reduce(
      (sum, session) => sum + (session.saved_count ?? 0),
      0,
    )
    const totalDoors = sessionsMatchingTags.reduce(
      (sum, session) => sum + (session.doors_knocked ?? 0),
      0,
    )
    const namedSaved = peopleMatchingTags.length
    const avgSaved = totalSessions ? totalSavedReported / totalSessions : 0
    return {
      totalSessions,
      totalSavedReported,
      totalDoors,
      namedSaved,
      avgSaved,
    }
  }, [sessionsMatchingTags, peopleMatchingTags])

  const groupRows = useMemo(() => {
    type GroupRow = {
      key: string
      label: string
      sortValue: number
      sessions: number
      savedReported: number
      namedSaved: number
      doors: number
    }
    const map = new Map<string, GroupRow>()

    sessionsMatchingTags.forEach((session) => {
      const info = getGroupInfo(session.date, groupBy, weekStartsOn)
      const existing = map.get(info.key)
      if (existing) {
        existing.sessions += 1
        existing.savedReported += session.saved_count ?? 0
        existing.doors += session.doors_knocked ?? 0
      } else {
        map.set(info.key, {
          key: info.key,
          label: info.label,
          sortValue: info.sortValue,
          sessions: 1,
          savedReported: session.saved_count ?? 0,
          namedSaved: 0,
          doors: session.doors_knocked ?? 0,
        })
      }
    })

    peopleMatchingTags.forEach((person) => {
      const session = sessionsById.get(person.session_id)
      if (!session) return
      if (!isWithinRange(session.date, rangeStart, rangeEnd)) return
      const info = getGroupInfo(session.date, groupBy, weekStartsOn)
      const existing = map.get(info.key)
      if (existing) {
        existing.namedSaved += 1
      } else {
        map.set(info.key, {
          key: info.key,
          label: info.label,
          sortValue: info.sortValue,
          sessions: 0,
          savedReported: 0,
          namedSaved: 1,
          doors: 0,
        })
      }
    })

    return Array.from(map.values()).sort((a, b) => b.sortValue - a.sortValue)
  }, [sessionsMatchingTags, peopleMatchingTags, sessionsById, groupBy, weekStartsOn, rangeStart, rangeEnd])

  const tagBreakdown = useMemo(() => {
    const counts = new Map<string, number>()
    peopleMatchingTags.forEach((person) => {
      person.tagIds.forEach((tagId) => {
        counts.set(tagId, (counts.get(tagId) ?? 0) + 1)
      })
    })
    return tags
      .map((tag) => ({
        tag,
        count: counts.get(tag.id) ?? 0,
      }))
      .filter((item) => item.count > 0)
      .sort((a, b) => b.count - a.count)
  }, [peopleMatchingTags, tags])

  const datasetBreakdown = useMemo(() => {
    const map = new Map<
      string,
      {
        dataset: DatasetRow
        sessions: number
        savedReported: number
        namedSaved: number
        doors: number
      }
    >()
    datasets.forEach((dataset) => {
      map.set(dataset.id, {
        dataset,
        sessions: 0,
        savedReported: 0,
        namedSaved: 0,
        doors: 0,
      })
    })
    sessionsMatchingTags.forEach((session) => {
      const entry = map.get(session.dataset_id)
      if (!entry) return
      entry.sessions += 1
      entry.savedReported += session.saved_count ?? 0
      entry.doors += session.doors_knocked ?? 0
    })
    peopleMatchingTags.forEach((person) => {
      const session = sessionsById.get(person.session_id)
      if (!session) return
      const entry = map.get(session.dataset_id)
      if (!entry) return
      entry.namedSaved += 1
    })
    return Array.from(map.values()).filter(
      (entry) => entry.sessions || entry.savedReported || entry.namedSaved,
    )
  }, [datasets, sessionsMatchingTags, peopleMatchingTags, sessionsById])

  const topSessions = useMemo(() => {
    return sessionsMatchingTags
      .map((session) => ({
        session,
        dataset: datasetById.get(session.dataset_id),
      }))
      .sort((a, b) => b.session.saved_count - a.session.saved_count)
      .slice(0, 5)
  }, [sessionsMatchingTags, datasetById])

  const maxSavedInGroup = useMemo(() => {
    if (!groupRows.length) return 1
    return Math.max(...groupRows.map((row) => row.savedReported), 1)
  }, [groupRows])

  const maxSessionsInGroup = useMemo(() => {
    if (!groupRows.length) return 1
    return Math.max(...groupRows.map((row) => row.sessions), 1)
  }, [groupRows])

  const maxTagCount = useMemo(() => {
    if (!tagBreakdown.length) return 1
    return Math.max(...tagBreakdown.map((item) => item.count), 1)
  }, [tagBreakdown])

  const ratioMetrics = useMemo(() => {
    const doorsPerSession = totals.totalSessions ? totals.totalDoors / totals.totalSessions : 0
    const doorsPerSalvation = totals.totalSavedReported
      ? totals.totalDoors / totals.totalSavedReported
      : 0
    return {
      doorsPerSession,
      doorsPerSalvation,
    }
  }, [totals])

  const timeMetrics = useMemo(() => {
    let totalEventMinutes = 0
    let eventsWithTime = 0
    sessionsMatchingTags.forEach((session) => {
      const duration = getSessionDurationMinutes(session.start_time, session.end_time)
      if (duration === null) return
      totalEventMinutes += duration
      eventsWithTime += 1
    })

    let totalStandaloneMinutes = 0
    let standaloneWithTime = 0
    peopleMatchingTags.forEach((person) => {
      const session = sessionsById.get(person.session_id)
      if (!session || session.type !== "standalone") return
      const minutes = person.time_spent_minutes
      if (minutes === null || minutes === undefined) return
      totalStandaloneMinutes += minutes
      standaloneWithTime += 1
    })

    const totalMinutes = totalEventMinutes + totalStandaloneMinutes
    const salvationsPerHour = totalMinutes
      ? totals.totalSavedReported / (totalMinutes / 60)
      : 0
    const minutesPerSalvation = totals.totalSavedReported
      ? totalMinutes / totals.totalSavedReported
      : 0
    const avgEventMinutes = eventsWithTime ? totalEventMinutes / eventsWithTime : 0
    const avgStandaloneMinutes = standaloneWithTime
      ? totalStandaloneMinutes / standaloneWithTime
      : 0

    return {
      totalMinutes,
      totalEventMinutes,
      totalStandaloneMinutes,
      eventsWithTime,
      standaloneWithTime,
      avgEventMinutes,
      avgStandaloneMinutes,
      salvationsPerHour,
      minutesPerSalvation,
    }
  }, [sessionsMatchingTags, peopleMatchingTags, sessionsById, totals.totalSavedReported])

  const weekdayRows = useMemo(() => {
    const labels = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    const rows = labels.map((label) => ({ label, sessions: 0, salvations: 0 }))
    sessionsMatchingTags.forEach((session) => {
      const dayIndex = new Date(`${session.date}T00:00:00`).getDay()
      const row = rows[dayIndex]
      if (!row) return
      row.sessions += 1
      row.salvations += session.saved_count ?? 0
    })
    return rows
  }, [sessionsMatchingTags])

  const maxWeekdaySalvations = useMemo(() => {
    if (!weekdayRows.length) return 1
    return Math.max(...weekdayRows.map((row) => row.salvations), 1)
  }, [weekdayRows])

  const maxWeekdaySessions = useMemo(() => {
    if (!weekdayRows.length) return 1
    return Math.max(...weekdayRows.map((row) => row.sessions), 1)
  }, [weekdayRows])

  const mostActiveDay = useMemo(() => {
    if (!weekdayRows.length) return null
    const maxSalvations = Math.max(...weekdayRows.map((row) => row.salvations))
    if (maxSalvations > 0) {
      const match = weekdayRows.find((row) => row.salvations === maxSalvations)
      return match ? { label: match.label, value: match.salvations, metric: "salvations" } : null
    }
    const maxSessions = Math.max(...weekdayRows.map((row) => row.sessions))
    if (maxSessions > 0) {
      const match = weekdayRows.find((row) => row.sessions === maxSessions)
      return match ? { label: match.label, value: match.sessions, metric: "events" } : null
    }
    return null
  }, [weekdayRows])

  const sessionTypeBreakdown = useMemo(() => {
    let sessionCount = 0
    let standaloneCount = 0
    sessionsMatchingTags.forEach((session) => {
      if (session.type === "standalone") {
        standaloneCount += 1
      } else {
        sessionCount += 1
      }
    })
    return {
      sessionCount,
      standaloneCount,
    }
  }, [sessionsMatchingTags])

  const salvationRoleCounts = useMemo(() => {
    return peopleMatchingTags.reduce(
      (acc, person) => {
        const role = normalizeSalvationRole(person.role)
        if (role === "partner") {
          acc.partner += 1
        } else {
          acc.presenter += 1
        }
        return acc
      },
      { presenter: 0, partner: 0 },
    )
  }, [peopleMatchingTags])

  const goalWindows = useMemo(() => {
    const today = new Date()
    const todayKey = toDateKey(today)
    const weekStart = startOfWeek(today, weekStartsOn)
    const weekEnd = new Date(weekStart)
    weekEnd.setDate(weekStart.getDate() + 6)
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1)
    const monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0)
    const yearStart = new Date(today.getFullYear(), 0, 1)
    const yearEnd = new Date(today.getFullYear(), 11, 31)
    const lastWeekStart = new Date(today)
    lastWeekStart.setDate(today.getDate() - 6)
    const lastMonthStart = new Date(today)
    lastMonthStart.setDate(today.getDate() - 29)
    const lastYearStart = new Date(today)
    lastYearStart.setDate(today.getDate() - 364)
    return {
      week: {
        label: `Week of ${dateFormatter.format(weekStart)}`,
        startKey: toDateKey(weekStart),
        endKey: toDateKey(weekEnd),
      },
      month: {
        label: monthFormatter.format(monthStart),
        startKey: toDateKey(monthStart),
        endKey: toDateKey(monthEnd),
      },
      year: {
        label: `${today.getFullYear()}`,
        startKey: toDateKey(yearStart),
        endKey: toDateKey(yearEnd),
      },
      overall: {
        label: "All time",
        startKey: "",
        endKey: "",
      },
      last_week: {
        label: "Last week",
        startKey: toDateKey(lastWeekStart),
        endKey: todayKey,
      },
      last_month: {
        label: "Last month",
        startKey: toDateKey(lastMonthStart),
        endKey: todayKey,
      },
      last_year: {
        label: "Last year",
        startKey: toDateKey(lastYearStart),
        endKey: todayKey,
      },
    }
  }, [weekStartsOn])

  const formatGoalRangeLabel = (startKey: string, endKey: string) => {
    const formatKey = (value: string) => {
      if (!value) return ""
      const parsed = new Date(`${value}T00:00:00`)
      return Number.isNaN(parsed.getTime()) ? value : dateFormatter.format(parsed)
    }
    const startLabel = formatKey(startKey)
    const endLabel = formatKey(endKey)
    if (startLabel && endLabel) return `${startLabel} - ${endLabel}`
    if (startLabel) return `Since ${startLabel}`
    if (endLabel) return `Through ${endLabel}`
    return "All time"
  }

  const resolveGoalWindow = (goal: GoalDefinition) => {
    if (goal.period === "custom") {
      const startKey = goal.startDate ?? ""
      const endKey = goal.endDate ?? ""
      return {
        startKey,
        endKey,
        label: startKey || endKey ? formatGoalRangeLabel(startKey, endKey) : "Custom range",
      }
    }
    return (
      goalWindows[goal.period as Exclude<GoalPeriod, "custom">] ?? goalWindows.week
    )
  }

  const goalCards = useMemo(() => {
    if (!goals.length) return []
    return goals.map((goal) => {
      const window = resolveGoalWindow(goal)
      const startKey = window.startKey
      const endKey = window.endKey
      const periodLabel = window.label
      const matchesDataset = (session: SessionRow) =>
        goal.datasetId === "all" || session.dataset_id === goal.datasetId
      let current = 0
      if (goal.metric === "sessions") {
        current = sessions.filter(
          (session) =>
            matchesDataset(session) && isWithinRange(session.date, startKey, endKey),
        ).length
      } else if (goal.metric === "doors") {
        current = sessions.reduce((sum, session) => {
          if (!matchesDataset(session)) return sum
          if (!isWithinRange(session.date, startKey, endKey)) return sum
          const doors =
            typeof session.doors_knocked === "number" && Number.isFinite(session.doors_knocked)
              ? session.doors_knocked
              : 0
          return sum + Math.max(0, doors)
        }, 0)
      } else if (goal.metric === "salvations" && goal.tagIds.length) {
        current = peopleWithTags.reduce((sum, person) => {
          const session = sessionsById.get(person.session_id)
          if (!session) return sum
          if (!matchesDataset(session)) return sum
          if (!isWithinRange(session.date, startKey, endKey)) return sum
          const matchesTags =
            goal.tagMatchMode === "all"
              ? goal.tagIds.every((tagId) => person.tagIds.includes(tagId))
              : goal.tagIds.some((tagId) => person.tagIds.includes(tagId))
          return matchesTags ? sum + 1 : sum
        }, 0)
      } else {
        current = sessions.reduce((sum, session) => {
          if (!matchesDataset(session)) return sum
          if (!isWithinRange(session.date, startKey, endKey)) return sum
          return sum + (session.saved_count ?? 0)
        }, 0)
      }
      const periodLabelPrefix =
        goal.period === "month"
          ? "Monthly"
          : goal.period === "year"
            ? "Yearly"
            : goal.period === "overall"
              ? "Overall"
              : goal.period === "last_week"
                ? "Last week"
                : goal.period === "last_month"
                  ? "Last month"
                  : goal.period === "last_year"
                    ? "Last year"
                    : goal.period === "custom"
                      ? "Custom range"
                      : "Weekly"
      const metricLabel =
        goal.metric === "sessions"
          ? "events"
          : goal.metric === "doors"
            ? "doors knocked"
            : "salvations"
      const label = goal.name.trim() || `${periodLabelPrefix} ${metricLabel}`
      const detailParts: string[] = []
      const datasetLabel =
        goal.datasetId === "all"
          ? "All data sets"
          : datasetById.get(goal.datasetId)?.name ?? "Unknown data set"
      detailParts.push(datasetLabel)
      if (goal.metric === "salvations" && goal.tagIds.length) {
        const tagLabels = goal.tagIds
          .map((tagId) => tagsById.get(tagId)?.name)
          .filter(Boolean)
        const matchLabel = goal.tagMatchMode === "all" ? "All tags" : "Any tag"
        if (tagLabels.length) {
          detailParts.push(`${matchLabel}: ${tagLabels.join(", ")}`)
        } else {
          detailParts.push(`${matchLabel}: (tags missing)`)
        }
      }
      return {
        id: goal.id,
        label,
        period: periodLabel,
        details: detailParts.join(" • "),
        current,
        goal: goal.target,
        metric: goal.metric,
      }
    })
  }, [
    datasetById,
    goalWindows,
    goals,
    peopleWithTags,
    resolveGoalWindow,
    sessions,
    sessionsById,
    tagsById,
  ])

  const activeGoalCards = useMemo(
    () => goalCards.filter((card) => card.goal > 0),
    [goalCards],
  )

  const hasGoals = useMemo(() => activeGoalCards.length > 0, [activeGoalCards])

  const updateStatusTone =
    updateStatus &&
    (updateStatus.toLowerCase().includes("error") ||
      updateStatus.toLowerCase().includes("failed"))
      ? "status--error"
      : "status--success"
  const syncLastUploadedLabel = formatTimestampLabel(syncLastUploadedAt)
  const syncLastDownloadedLabel = formatTimestampLabel(syncLastDownloadedAt)
  const rollbackSavedLabel = formatTimestampLabel(rollbackSavedAt)

  const settingsDetailLabel =
    settingsDetail === "sessions"
      ? "Events"
      : settingsDetail === "salvations"
        ? "Salvations"
        : settingsDetail === "tags"
          ? "Tags"
          : settingsDetail === "datasets"
            ? "Data sets"
            : ""

  const activeStatsView = useMemo(() => {
    if (!statsViews.length) return DEFAULT_STATS_VIEWS[0]
    return statsViews.find((view) => view.id === activeStatsViewId) ?? statsViews[0]
  }, [statsViews, activeStatsViewId])

  const activeStatsSections = useMemo(
    () => new Set(activeStatsView?.sections ?? []),
    [activeStatsView],
  )

  const statsSectionLabelMap = useMemo(() => {
    return new Map<StatsSection, string>(STATS_SECTION_OPTIONS.map((section) => [section.id, section.label]))
  }, [])

  const buildTagUsagePreview = (tagId: string, limit = TAG_DETAIL_PREVIEW_LIMIT) => {
    const entries = tagUsageMap.get(tagId) ?? []
    if (!entries.length) {
      return {
        hasEntries: false,
        label: "No salvations tagged yet.",
        remaining: 0,
      }
    }
    const preview = entries.slice(0, limit).map((entry) => {
      const name = entry.person.name || "Unknown Name"
      const dateLabel = entry.session
        ? dateFormatter.format(new Date(`${entry.session.date}T00:00:00`))
        : "Date unknown"
      return `${name} (${dateLabel})`
    })
    return {
      hasEntries: true,
      label: preview.join(", "),
      remaining: entries.length - preview.length,
    }
  }

  const salvationResults = useMemo(() => {
    const query = salvationQuery.trim().toLowerCase()
    const filtered = peopleMatchingTags.filter((person) => {
      if (!query) return true
      return (
        person.name.toLowerCase().includes(query) ||
        (person.notes ? person.notes.toLowerCase().includes(query) : false)
      )
    })
    return filtered
      .map((person) => {
        const session = sessionsById.get(person.session_id)
        return {
          person,
          session,
          dataset: session ? datasetById.get(session.dataset_id) : undefined,
        }
      })
      .sort((a, b) => {
        const aDate = a.session?.date ?? ""
        const bDate = b.session?.date ?? ""
        if (aDate !== bDate) return bDate.localeCompare(aDate)
        return a.person.name.localeCompare(b.person.name)
      })
  }, [peopleMatchingTags, salvationQuery, sessionsById, datasetById])
  const handleImportClick = () => {
    fileInputRef.current?.click()
  }

  const applyImportedData = useCallback((normalized: AppData, message?: string) => {
    setSessions(normalized.sessions)
    setPeople(normalized.people)
    setTags([...normalized.tags].sort((a, b) => a.name.localeCompare(b.name)))
    setDatasets(normalized.datasets)
    setGoals(normalized.goals)
    setStatsViews(normalized.statsViews)
    setActiveStatsViewId(normalized.statsViews[0]?.id ?? "")
    setSelectedTagIds([])
    setSelectedDatasetIds([])
    setLogDatasetIds([])
    setLogQuery("")
    setStatsQuery("")
    setStatsViewName("")
    setStatsViewSections([])
    setStatsBuilderOpen(false)
    setEditingStatsViewId(null)
    resetSessionDraft()
    resetStandaloneDraft()
    setEditingLogSessionId(null)
    setEditEventName("")
    setEditEventDate("")
    setEditEventDatasetId("")
    setEditEventSavedCount("")
    setEditEventDoors("")
    setEditEventStartTime("")
    setEditEventEndTime("")
    setEditEventNotes("")
    setEditEventTagIds([])
    setEditEventTagQuery("")
    setEditingLogPersonId(null)
    setLogPersonName("")
    setLogPersonTags([])
    setLogPersonNotes("")
    setLogPersonTagQuery("")
    setLogPersonTimeSpent("")
    setLogAddSessionId(null)
    setLogAddPersonName("")
    setLogAddPersonTags([])
    setLogAddPersonNotes("")
    setLogAddPersonTagQuery("")
    if (message) {
      setActionMessage(message)
    }
  }, [resetSessionDraft, resetStandaloneDraft])

  const persistUserSyncState = useCallback(
    ({
      signature,
      uploadedAt,
      downloadedAt,
    }: {
      signature: string
      uploadedAt: string
      downloadedAt: string
    }) => {
      if (!authUserId) return
      const syncedAt = new Date().toISOString()
      saveUserSyncState(authUserId, {
        signature,
        synced_at: syncedAt,
        uploaded_at: uploadedAt,
        downloaded_at: downloadedAt,
      })
      lastCloudUploadSignatureRef.current = signature
      setSyncLastUploadedAt(uploadedAt)
      setSyncLastDownloadedAt(downloadedAt)
    },
    [authUserId],
  )

  const saveRollbackGuardSnapshot = useCallback(
    (reason: string) => {
      const savedAt = new Date().toISOString()
      saveRollbackSnapshot({
        saved_at: savedAt,
        reason,
        data: appDataSnapshot,
      })
      setRollbackSavedAt(savedAt)
    },
    [appDataSnapshot],
  )

  const handleRestoreRollbackSnapshot = () => {
    clearMessages()
    const snapshot = loadRollbackSnapshot()
    if (!snapshot) {
      setActionError("No rollback snapshot found.")
      return
    }
    if (!window.confirm("Restore the previous local snapshot? This replaces your current local data.")) {
      return
    }
    const normalized = normalizeImportedData(snapshot.data)
    applyImportedData(normalized, "Rollback restored from local safety snapshot.")
    if (authUserId) {
      lastCloudUploadSignatureRef.current = ""
      hasAutoPulledUserIdRef.current = authUserId
    }
    setAutoSyncError("")
  }

  const uploadCloudSnapshot = useCallback(
    async ({
      force = false,
      showBusyState = false,
      showSuccessMessage = false,
      showErrorMessage = true,
      successMessage = "Cloud backup uploaded.",
    }: {
      force?: boolean
      showBusyState?: boolean
      showSuccessMessage?: boolean
      showErrorMessage?: boolean
      successMessage?: string
    } = {}) => {
      if (!navigator.onLine) {
        if (showErrorMessage) {
          setActionError("You are offline. Changes will sync when connection is restored.")
        }
        return false
      }
      if (!supabase || !isSupabaseConfigured) {
        if (showErrorMessage) {
          setActionError("Supabase is not configured for this deployment.")
        }
        return false
      }
      if (!authUserId) {
        if (showErrorMessage) {
          setActionError("Sign in first to upload a cloud backup.")
        }
        return false
      }

      if (cloudSyncInFlightRef.current) {
        if (!showBusyState && !showErrorMessage) {
          cloudSyncRetryQueuedRef.current = true
        } else if (showErrorMessage) {
          setActionError("Cloud sync already in progress. Please wait a moment and try again.")
        }
        return false
      }

      const payload = appDataSnapshot
      const signature = buildCloudSignature(authUserId, payload)
      if (!force && signature === lastCloudUploadSignatureRef.current) {
        return true
      }

      cloudSyncInFlightRef.current = true
      if (showBusyState) {
        setSyncLoading(true)
      } else {
        setAutoSyncing(true)
      }

      try {
        const { data: existingSnapshot, error: existingSnapshotError } = await supabase
          .from(SUPABASE_SNAPSHOT_TABLE)
          .select("payload")
          .eq("user_id", authUserId)
          .maybeSingle()
        if (existingSnapshotError) throw existingSnapshotError

        let remoteSignature = ""
        if (
          existingSnapshot &&
          typeof existingSnapshot.payload === "object" &&
          existingSnapshot.payload !== null
        ) {
          const normalizedRemote = normalizeImportedData(existingSnapshot.payload)
          remoteSignature = buildCloudSignature(authUserId, normalizedRemote)
        }

        const knownSyncedSignature = lastCloudUploadSignatureRef.current
        const remoteChangedSinceLastSync =
          Boolean(remoteSignature) &&
          Boolean(knownSyncedSignature) &&
          remoteSignature !== knownSyncedSignature
        const unknownSyncHistoryWithDifferentRemote =
          !knownSyncedSignature && Boolean(remoteSignature) && remoteSignature !== signature

        if (remoteChangedSinceLastSync || unknownSyncHistoryWithDifferentRemote) {
          const conflictMessage =
            "Cloud backup changed on another device. Download from cloud first, or use Sync now and confirm overwrite."
          if (!force) {
            if (showErrorMessage) {
              setActionError(conflictMessage)
            } else {
              setAutoSyncError(conflictMessage)
            }
            return false
          }
          const confirmed = window.confirm(
            "Cloud backup has changed on another device. Sync now will overwrite cloud data. Continue?",
          )
          if (!confirmed) {
            return false
          }
        }

        const updatedAt = new Date().toISOString()
        const { error } = await supabase.from(SUPABASE_SNAPSHOT_TABLE).upsert(
          {
            user_id: authUserId,
            payload,
            updated_at: updatedAt,
          },
          { onConflict: "user_id" },
        )
        if (error) throw error
        persistUserSyncState({
          signature,
          uploadedAt: updatedAt,
          downloadedAt: syncLastDownloadedAt,
        })
        setAutoSyncError("")
        if (showSuccessMessage) {
          setActionMessage(successMessage)
        }
        return true
      } catch (error) {
        const message = error instanceof Error ? error.message : "Cloud upload failed."
        if (showErrorMessage) {
          setActionError(message)
        } else {
          setAutoSyncError(message)
        }
        return false
      } finally {
        cloudSyncInFlightRef.current = false
        if (showBusyState) {
          setSyncLoading(false)
        } else {
          setAutoSyncing(false)
        }

        if (cloudSyncRetryQueuedRef.current) {
          cloudSyncRetryQueuedRef.current = false
          window.setTimeout(() => {
            void uploadCloudSnapshot({
              force: false,
              showBusyState: false,
              showSuccessMessage: false,
              showErrorMessage: false,
            })
          }, 0)
        }
      }
    },
    [appDataSnapshot, authUserId, persistUserSyncState, syncLastDownloadedAt],
  )

  const downloadCloudSnapshot = useCallback(
    async ({
      force = false,
      showBusyState = false,
      showSuccessMessage = true,
      showErrorMessage = true,
      successMessage = "Cloud backup downloaded.",
    }: {
      force?: boolean
      showBusyState?: boolean
      showSuccessMessage?: boolean
      showErrorMessage?: boolean
      successMessage?: string
    } = {}) => {
      if (!navigator.onLine) {
        if (showErrorMessage) {
          setActionError("You are offline. Cannot download cloud backup.")
        }
        return false
      }
      if (!supabase || !isSupabaseConfigured) {
        if (showErrorMessage) {
          setActionError("Supabase is not configured for this deployment.")
        }
        return false
      }
      if (!authUserId) {
        if (showErrorMessage) {
          setActionError("Sign in first to download a cloud backup.")
        }
        return false
      }

      if (cloudSyncInFlightRef.current) {
        if (showErrorMessage) {
          setActionError("Cloud sync already in progress. Please wait a moment and try again.")
        }
        return false
      }

      cloudSyncInFlightRef.current = true
      if (showBusyState) {
        setSyncLoading(true)
      } else {
        setAutoSyncing(true)
      }

      try {
        const { data, error } = await supabase
          .from(SUPABASE_SNAPSHOT_TABLE)
          .select("payload, updated_at")
          .eq("user_id", authUserId)
          .maybeSingle()
        if (error) throw error
        if (!data || typeof data.payload !== "object" || data.payload === null) {
          if (showErrorMessage) {
            setActionError("No cloud backup found for this account.")
          }
          return false
        }

        const normalized = normalizeImportedData(data.payload)
        const remoteSignature = buildCloudSignature(authUserId, normalized)
        const localSignature = buildCloudSignature(authUserId, appDataSnapshot)
        const knownSyncedSignature = lastCloudUploadSignatureRef.current
        const localUnsyncedSinceLastSync =
          Boolean(knownSyncedSignature) && localSignature !== knownSyncedSignature
        const localUnsyncedWithoutHistory =
          !knownSyncedSignature && hasMeaningfulData(appDataSnapshot)
        const wouldOverwriteLocal = remoteSignature !== localSignature
        const downloadedAt = normalizeTimestampValue(data.updated_at) || new Date().toISOString()
        setAutoSyncError("")

        if (!force && wouldOverwriteLocal && (localUnsyncedSinceLastSync || localUnsyncedWithoutHistory)) {
          const conflictMessage =
            "Conflict detected: local unsynced data was kept. Use Sync now to upload local changes, or Download from cloud to overwrite local."
          if (showErrorMessage) {
            setActionError(conflictMessage)
          } else {
            setAutoSyncError(conflictMessage)
          }
          return false
        }

        if (!force && !wouldOverwriteLocal) {
          persistUserSyncState({
            signature: remoteSignature,
            uploadedAt: syncLastUploadedAt,
            downloadedAt,
          })
          return true
        }

        if (wouldOverwriteLocal) {
          saveRollbackGuardSnapshot("before-cloud-download")
        }
        cloudSyncRetryQueuedRef.current = false
        persistUserSyncState({
          signature: remoteSignature,
          uploadedAt: syncLastUploadedAt,
          downloadedAt,
        })
        applyImportedData(normalized, showSuccessMessage ? successMessage : undefined)
        return true
      } catch (error) {
        const message = error instanceof Error ? error.message : "Cloud download failed."
        if (showErrorMessage) {
          setActionError(message)
        } else {
          setAutoSyncError(message)
        }
        return false
      } finally {
        cloudSyncInFlightRef.current = false
        if (showBusyState) {
          setSyncLoading(false)
        } else {
          setAutoSyncing(false)
        }
      }
    },
    [
      appDataSnapshot,
      authUserId,
      applyImportedData,
      persistUserSyncState,
      saveRollbackGuardSnapshot,
      syncLastUploadedAt,
    ],
  )

  useEffect(() => {
    if (!supabase || !isSupabaseConfigured) return
    if (!authUserId) return
    if (hasAutoPulledUserIdRef.current === authUserId) return

    hasAutoPulledUserIdRef.current = authUserId
    void downloadCloudSnapshot({
      force: false,
      showBusyState: false,
      showSuccessMessage: true,
      showErrorMessage: false,
      successMessage: "Cloud data synced from your account.",
    })
  }, [authUserId, downloadCloudSnapshot])

  useEffect(() => {
    if (!supabase || !isSupabaseConfigured) return
    if (!authUserId) return
    if (authLoading || syncLoading || loading) return
    const timer = window.setTimeout(() => {
      void uploadCloudSnapshot({
        force: false,
        showBusyState: false,
        showSuccessMessage: false,
        showErrorMessage: false,
      })
    }, AUTO_SYNC_DEBOUNCE_MS)
    return () => window.clearTimeout(timer)
  }, [appDataSnapshot, authLoading, authUserId, loading, syncLoading, uploadCloudSnapshot, isOffline])

  useEffect(() => {
    const updateNetworkStatus = () => {
      if (typeof navigator !== "undefined") {
        setIsOffline(!navigator.onLine)
      }
    }

    window.addEventListener("online", updateNetworkStatus)
    window.addEventListener("offline", updateNetworkStatus)

    // Catch any changes that occurred between initial render and effect execution
    updateNetworkStatus()

    // Fallback polling: PWAs loading from Service Workers (and DevTools Offline mode)
    // can sometimes silently flip navigator.onLine to false during boot without 
    // reliably triggering the window 'offline' event.
    const intervalId = setInterval(() => {
      if (typeof navigator !== "undefined") {
        setIsOffline((prev) => {
          const currentOffline = !navigator.onLine
          return prev !== currentOffline ? currentOffline : prev
        })
      }
    }, 1000)

    return () => {
      window.removeEventListener("online", updateNetworkStatus)
      window.removeEventListener("offline", updateNetworkStatus)
      clearInterval(intervalId)
    }
  }, [])
  
  const handleAuthSubmit = async (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    if (!navigator.onLine) {
      setActionError("You are offline. Please connect to the internet to sign in.")
      return
    }
    if (!supabase || !isSupabaseConfigured) {
      setActionError("Supabase is not configured for this deployment.")
      return
    }
    const email = authEmail.trim().toLowerCase()
    const password = authPassword
    if (!email || !password) {
      setActionError("Email and password are required.")
      return
    }
    setAuthLoading(true)
    try {
      if (authMode === "signin") {
        const { error } = await supabase.auth.signInWithPassword({ email, password })
        if (error) throw error
        setActionMessage("Signed in.")
      } else {
        const { data, error } = await supabase.auth.signUp({ email, password })
        if (error) throw error
        if (data.session) {
          setActionMessage("Account created and signed in.")
        } else {
          setActionMessage("Account created. Check your email to confirm, then sign in.")
        }
        setAuthMode("signin")
      }
      setAuthPassword("")
    } catch (error) {
      setActionError(error instanceof Error ? error.message : "Authentication failed.")
    } finally {
      setAuthLoading(false)
    }
  }

  const handleSignOut = async () => {
    clearMessages()
    if (!navigator.onLine) {
      setActionError("You are offline. Please connect to the internet to sign out.")
      return
    }
    if (!supabase || !isSupabaseConfigured) {
      setActionError("Supabase is not configured for this deployment.")
      return
    }
    if (cloudSyncInFlightRef.current) {
      setActionError("Cloud sync already in progress. Please wait a moment and try again.")
      return
    }
    setAuthLoading(true)
    try {
      const { error } = await supabase.auth.signOut()
      if (error) throw error
      setAuthPassword("")
      setActionMessage("Signed out.")
    } catch (error) {
      setActionError(error instanceof Error ? error.message : "Sign-out failed.")
    } finally {
      setAuthLoading(false)
    }
  }

  const handleCloudSyncUpload = async () => {
    clearMessages()
    if (!navigator.onLine) {
      setActionError("You are offline. Changes will sync when connection is restored.")
      return
    }
    if (syncLoading || authLoading) return
    await uploadCloudSnapshot({
      force: true,
      showBusyState: true,
      showSuccessMessage: true,
      showErrorMessage: true,
      successMessage: "Cloud backup uploaded.",
    })
  }

  const handleCloudSyncDownload = async () => {
    clearMessages()
    if (!navigator.onLine) {
      setActionError("You are offline. Cannot download cloud backup.")
      return
    }
    if (syncLoading || authLoading) return
    if (!authUserId) {
      setActionError("Sign in first to download a cloud backup.")
      return
    }
    if (!window.confirm("Download from cloud and replace your local data on this device?")) {
      return
    }
    await downloadCloudSnapshot({
      force: true,
      showBusyState: true,
      showSuccessMessage: true,
      showErrorMessage: true,
      successMessage: "Cloud backup downloaded.",
    })
  }

  const handleImportJson = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const input = event.target
    const file = input.files?.[0]
    if (!file) return
    clearMessages()
    setLoading(true)
    try {
      const text = await file.text()
      const parsed = JSON.parse(text)
      const normalized = normalizeImportedData(parsed)
      if (!normalized.sessions.length && !normalized.people.length && !normalized.tags.length) {
        setActionError("No data found in that file.")
        return
      }
      if (!window.confirm("Importing will replace your current data. Continue?")) {
        return
      }
      applyImportedData(normalized, "Data imported.")
    } catch {
      setActionError("Could not import that file.")
    } finally {
      setLoading(false)
      input.value = ""
    }
  }

  const exportJson = () => {
    downloadJson("soulwinning-backup.json", {
      version: 1,
      exported_at: new Date().toISOString(),
      sessions,
      people,
      tags,
      datasets,
      goals,
      statsViews,
    })
  }

  const exportDatasetBackup = (
    dataset: DatasetRow,
    sessionsForDataset: SessionRow[],
    peopleForDataset: PersonRow[],
  ) => {
    const tagIds = new Set<string>()
    peopleForDataset.forEach((person) => {
      person.tagIds.forEach((tagId) => tagIds.add(tagId))
    })
    sessionsForDataset.forEach((session) => {
      session.tagIds.forEach((tagId) => tagIds.add(tagId))
    })
    const tagsForDataset = tags.filter((tag) => tagIds.has(tag.id))
    const filename = `soulwinning-dataset-${slugify(dataset.name)}.json`
    downloadJson(filename, {
      version: 1,
      dataset_backup: true,
      exported_at: new Date().toISOString(),
      datasets: [dataset],
      sessions: sessionsForDataset,
      people: peopleForDataset,
      tags: tagsForDataset,
    })
  }

  const handleClearAllData = () => {
    if (
      !window.confirm(
        "This will delete all events, salvations, tags, and data sets on this device. Continue?",
      )
    ) {
      return
    }
    clearMessages()
    const empty = buildEmptyData()
    setSessions(empty.sessions)
    setPeople(empty.people)
    setTags(empty.tags)
    setDatasets(empty.datasets)
    setSelectedTagIds([])
    setSelectedDatasetIds([])
    setDraftSessionName("")
    setDraftDate(toDateKey(new Date()))
    setDraftSavedCount("0")
    setDraftStartTime("")
    setDraftEndTime("")
    setDraftDoors("")
    setDraftSessionNotes("")
    setDraftDatasetId(empty.datasets[0]?.id ?? "")
    setDraftEventTagIds([])
    setEventTagQuery("")
    setSessionPeople([])
    setSessionSalvationsOpen(false)
    setEditingSessionPersonId(null)
    setSessionPersonNotice("")
    setSessionPersonHighlightId(null)
    setSessionPersonName("")
    setSessionPersonTags([])
    setSessionPersonNotes("")
    setSessionTagQuery("")
    setStandaloneDate(toDateKey(new Date()))
    setStandaloneDatasetId(empty.datasets[0]?.id ?? "")
    setStandalonePersonName("")
    setStandalonePersonTags([])
    setStandaloneTimeSpent("")
    setStandalonePersonNotes("")
    setStandaloneTagQuery("")
    setEditingLogSessionId(null)
    setEditEventName("")
    setEditEventDate("")
    setEditEventDatasetId("")
    setEditEventSavedCount("")
    setEditEventDoors("")
    setEditEventStartTime("")
    setEditEventEndTime("")
    setEditEventNotes("")
    setEditEventTagIds([])
    setEditEventTagQuery("")
    setEditingLogPersonId(null)
    setLogPersonName("")
    setLogPersonTags([])
    setLogPersonNotes("")
    setLogPersonTagQuery("")
    setLogPersonTimeSpent("")
    setLogAddSessionId(null)
    setLogAddPersonName("")
    setLogAddPersonTags([])
    setLogAddPersonNotes("")
    setLogAddPersonTagQuery("")
    setTagName("")
    setTagColor(TAG_COLORS[0])
    setDatasetName("")
    setGoalName("")
    setGoalMetric("sessions")
    setGoalPeriod("week")
    setGoalTarget("0")
    setGoalDatasetId("all")
    setGoalTagIds([])
    setGoalTagMatchMode("any")
    setGoalTagQuery("")
    setSalvationQuery("")
    setLogFilter("both")
    setLogQuery("")
    setLogDatasetIds([])
    setStatsQuery("")
    setStatsViewName("")
    setStatsViewSections([])
    setStatsBuilderOpen(false)
    setEditingStatsViewId(null)
    setSettingsDetail(null)
    setGoals(empty.goals)
    setStatsViews(empty.statsViews)
    setActiveStatsViewId(empty.statsViews[0]?.id ?? "")
    setActionMessage("All data cleared.")
  }

  const handleCreateTag = (event: React.FormEvent) => {
    event.preventDefault()
    const created = createTag(tagName, tagColor)
    if (!created) return
    setTagName("")
    setTagColor(TAG_COLORS[0])
  }

  const handleStartTagEdit = (tag: TagRow) => {
    clearMessages()
    setEditingTagId(tag.id)
    setEditTagName(tag.name)
    setEditTagColor(tag.color ?? TAG_COLORS[0])
  }

  const handleCancelTagEdit = () => {
    clearMessages()
    setEditingTagId(null)
    setEditTagName("")
    setEditTagColor(TAG_COLORS[0])
  }

  const handleSaveTagEdit = (event: React.FormEvent, tagId: string) => {
    event.preventDefault()
    clearMessages()
    const name = editTagName.trim()
    if (!name) {
      setActionError("Tag name is required.")
      return
    }
    if (
      tags.some((tag) => tag.id !== tagId && tag.name.toLowerCase() === name.toLowerCase())
    ) {
      setActionError("A tag with that name already exists.")
      return
    }
    setTags((prev) =>
      prev
        .map((tag) => (tag.id === tagId ? { ...tag, name, color: editTagColor } : tag))
        .sort((a, b) => a.name.localeCompare(b.name)),
    )
    setEditingTagId(null)
    setEditTagName("")
    setEditTagColor(TAG_COLORS[0])
    setActionMessage("Tag updated.")
  }

  const handleDeleteTag = (tagId: string) => {
    if (
      !window.confirm("Delete this tag? It will be removed from all salvations and events.")
    ) {
      return
    }
    clearMessages()
    setTags((prev) => prev.filter((tag) => tag.id !== tagId))
    setSelectedTagIds((prev) => prev.filter((id) => id !== tagId))
    setPeople((prev) =>
      prev.map((person) => ({
        ...person,
        tagIds: person.tagIds.filter((id) => id !== tagId),
      })),
    )
    setSessions((prev) =>
      prev.map((session) => ({
        ...session,
        tagIds: session.tagIds.filter((id) => id !== tagId),
      })),
    )
    setSessionPeople((prev) =>
      prev.map((person) => ({
        ...person,
        tagIds: person.tagIds.filter((id) => id !== tagId),
      })),
    )
    setDraftEventTagIds((prev) => prev.filter((id) => id !== tagId))
    setSessionPersonTags((prev) => prev.filter((id) => id !== tagId))
    setStandalonePersonTags((prev) => prev.filter((id) => id !== tagId))
    setEditEventTagIds((prev) => prev.filter((id) => id !== tagId))
    setLogPersonTags((prev) => prev.filter((id) => id !== tagId))
    setLogAddPersonTags((prev) => prev.filter((id) => id !== tagId))
    setGoalTagIds((prev) => prev.filter((id) => id !== tagId))
    setGoals((prev) =>
      prev.map((goal) => ({
        ...goal,
        tagIds: goal.tagIds.filter((id) => id !== tagId),
      })),
    )
    if (editingTagId === tagId) {
      setEditingTagId(null)
      setEditTagName("")
      setEditTagColor(TAG_COLORS[0])
    }
    setActionMessage("Tag deleted.")
  }

  const handleCreateDataset = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    const name = datasetName.trim()
    if (!name) {
      setActionError("Data set name is required.")
      return
    }
    const newDataset: DatasetRow = {
      id: createLocalId(),
      name,
      created_at: new Date().toISOString(),
    }
    setDatasets((prev) => [...prev, newDataset].sort((a, b) => a.name.localeCompare(b.name)))
    setDatasetName("")
    setActionMessage("Data set added.")
  }

  const handleStartDatasetRename = (dataset: DatasetRow) => {
    clearMessages()
    setEditingDatasetId(dataset.id)
    setDatasetRename(dataset.name)
  }

  const handleCancelDatasetRename = () => {
    clearMessages()
    setEditingDatasetId(null)
    setDatasetRename("")
  }

  const handleRenameDataset = (event: React.FormEvent, datasetId: string) => {
    event.preventDefault()
    clearMessages()
    const name = datasetRename.trim()
    if (!name) {
      setActionError("Data set name is required.")
      return
    }
    if (
      datasets.some(
        (dataset) => dataset.id !== datasetId && dataset.name.toLowerCase() === name.toLowerCase(),
      )
    ) {
      setActionError("A data set with that name already exists.")
      return
    }
    setDatasets((prev) =>
      prev
        .map((dataset) => (dataset.id === datasetId ? { ...dataset, name } : dataset))
        .sort((a, b) => a.name.localeCompare(b.name)),
    )
    setEditingDatasetId(null)
    setDatasetRename("")
    setActionMessage("Data set renamed.")
  }

  const handleDeleteDataset = (datasetId: string) => {
    const dataset = datasetById.get(datasetId)
    if (!dataset) return
    const sessionsForDataset = sessions.filter((session) => session.dataset_id === datasetId)
    const sessionIds = new Set(sessionsForDataset.map((session) => session.id))
    const peopleForDataset = people.filter((person) => sessionIds.has(person.session_id))
    const eventCountLabel = formatEventCount(sessionsForDataset.length)
    const salvationCountLabel = formatSalvationCount(peopleForDataset.length)
    const shouldBackup = window.confirm(
      `Back up "${dataset.name}" before deleting? Click OK to download a backup file, or Cancel to delete without a backup.`,
    )
    if (shouldBackup) {
      exportDatasetBackup(dataset, sessionsForDataset, peopleForDataset)
    }
    const deleteMessage = shouldBackup
      ? `Delete "${dataset.name}" now? This will permanently remove ${eventCountLabel} and ${salvationCountLabel} from this device.`
      : `Delete "${dataset.name}" without a backup? This will permanently remove ${eventCountLabel} and ${salvationCountLabel}. This cannot be undone.`
    if (!window.confirm(deleteMessage)) return
    clearMessages()
    setSessions((prev) => prev.filter((session) => session.dataset_id !== datasetId))
    setPeople((prev) => prev.filter((person) => !sessionIds.has(person.session_id)))
    setDatasets((prev) => {
      const remaining = prev.filter((item) => item.id !== datasetId)
      return remaining.length ? remaining : [createDefaultDataset()]
    })
    setSelectedDatasetIds((prev) => prev.filter((id) => id !== datasetId))
    setGoals((prev) =>
      prev.map((goal) =>
        goal.datasetId === datasetId ? { ...goal, datasetId: "all" } : goal,
      ),
    )
    if (editingDatasetId === datasetId) {
      setEditingDatasetId(null)
      setDatasetRename("")
    }
    if (shouldBackup) {
      setActionMessage(`Data set "${dataset.name}" deleted.`)
    } else {
      setActionMessage(`Data set "${dataset.name}" deleted without backup. Data is permanently lost.`)
    }
  }

  const handleSaveSession = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    const savedCount = Number.parseInt(draftSavedCount, 10)
    const doorsValue = draftDoors.trim()
    const doorsCount = doorsValue ? Number.parseInt(doorsValue, 10) : null
    const startTime = draftStartTime.trim() || null
    const endTime = draftEndTime.trim() || null
    const datasetId = draftDatasetId || defaultDatasetId

    if (!draftDate) {
      setActionError("Date is required.")
      return
    }
    if (!datasetId) {
      setActionError("Data set is required.")
      return
    }
    if (Number.isNaN(savedCount) || savedCount < 0) {
      setActionError("Salvations must be 0 or more.")
      return
    }
    if (doorsCount !== null && (Number.isNaN(doorsCount) || doorsCount < 0)) {
      setActionError("Doors knocked must be 0 or more.")
      return
    }
    if (startTime && !normalizeTimeValue(startTime)) {
      setActionError("Start time must be valid.")
      return
    }
    if (endTime && !normalizeTimeValue(endTime)) {
      setActionError("End time must be valid.")
      return
    }
    if (startTime && endTime && getSessionDurationMinutes(startTime, endTime) === null) {
      setActionError("End time must be after the start time.")
      return
    }

    const sessionId = createLocalId()
    const createdAt = new Date().toISOString()
    const sessionData: SessionRow = {
      id: sessionId,
      date: draftDate,
      name: draftSessionName.trim(),
      type: "session",
      dataset_id: datasetId,
      saved_count: savedCount,
      doors_knocked: doorsCount,
      tagIds: draftEventTagIds,
      start_time: startTime,
      end_time: endTime,
      notes: draftSessionNotes.trim(),
      created_at: createdAt,
    }

    const preparedPeople = editingSessionPersonId
      ? sessionPeople.map((person) =>
          person.id === editingSessionPersonId
            ? {
                ...person,
                name: sessionPersonName.trim() || "Unknown Name",
                tagIds: sessionPersonTags,
                role: sessionPersonRole,
                notes: sessionPersonNotes.trim(),
              }
            : person,
        )
      : sessionPeople

    const peopleToInsert = preparedPeople
      .map((person) => ({
        ...person,
        name: person.name.trim() || "Unknown Name",
        notes: person.notes.trim(),
      }))
      .map((person) => ({
        id: createLocalId(),
        session_id: sessionId,
        name: person.name,
        tagIds: person.tagIds,
        role: person.role,
        time_spent_minutes: null,
        notes: person.notes,
        created_at: createdAt,
      }))

    setSessions((prev) => [sessionData, ...prev])
    if (peopleToInsert.length) {
      setPeople((prev) => [...peopleToInsert, ...prev])
    }

    resetSessionDraft()
    setActionMessage("Event saved successfully!")
  }

  const handleDeleteSession = (sessionId: string) => {
    if (!window.confirm("Delete this event? This removes salvations for that event.")) return
    clearMessages()
    setSessions((prev) => prev.filter((session) => session.id !== sessionId))
    setPeople((prev) => prev.filter((person) => person.session_id !== sessionId))
    if (editingLogSessionId === sessionId) {
      setEditingLogSessionId(null)
      setEditEventName("")
      setEditEventDate("")
      setEditEventDatasetId("")
      setEditEventSavedCount("")
      setEditEventDoors("")
      setEditEventStartTime("")
      setEditEventEndTime("")
      setEditEventNotes("")
      setEditEventTagIds([])
      setEditEventTagQuery("")
    }
    const editingPerson = editingLogPersonId
      ? people.find((person) => person.id === editingLogPersonId)
      : null
    if (editingPerson?.session_id === sessionId) {
      setEditingLogPersonId(null)
      setLogPersonName("")
      setLogPersonTags([])
      setLogPersonNotes("")
      setLogPersonTagQuery("")
      setLogPersonTimeSpent("")
    }
    if (logAddSessionId === sessionId) {
      setLogAddSessionId(null)
      setLogAddPersonName("")
      setLogAddPersonTags([])
      setLogAddPersonNotes("")
      setLogAddPersonTagQuery("")
    }
    setActionMessage("Event deleted.")
  }

  const handleEditLogSession = (sessionId: string) => {
    const session = sessions.find((item) => item.id === sessionId)
    if (!session) return
    clearMessages()
    setEditingLogSessionId(sessionId)
    setEditEventName(session.name)
    setEditEventDate(session.date)
    setEditEventDatasetId(session.dataset_id)
    setEditEventSavedCount(String(session.saved_count ?? 0))
    setEditEventDoors(
      session.doors_knocked === null || session.doors_knocked === undefined
        ? ""
        : String(session.doors_knocked),
    )
    setEditEventStartTime(session.start_time ?? "")
    setEditEventEndTime(session.end_time ?? "")
    setEditEventNotes(session.notes ?? "")
    setEditEventTagIds(session.tagIds ?? [])
    setEditEventTagQuery("")
  }

  const resetLogSessionEdit = () => {
    setEditingLogSessionId(null)
    setEditEventName("")
    setEditEventDate("")
    setEditEventDatasetId("")
    setEditEventSavedCount("")
    setEditEventDoors("")
    setEditEventStartTime("")
    setEditEventEndTime("")
    setEditEventNotes("")
    setEditEventTagIds([])
    setEditEventTagQuery("")
  }

  const handleCancelLogSessionEdit = () => {
    clearMessages()
    resetLogSessionEdit()
  }

  const handleSaveLogSessionEdit = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    if (!editingLogSessionId) return
    const savedCount = Number.parseInt(editEventSavedCount, 10)
    const doorsValue = editEventDoors.trim()
    const doorsCount = doorsValue ? Number.parseInt(doorsValue, 10) : null
    const startTime = editEventStartTime.trim() || null
    const endTime = editEventEndTime.trim() || null
    const datasetId = editEventDatasetId || defaultDatasetId

    if (!editEventDate) {
      setActionError("Date is required.")
      return
    }
    if (!datasetId) {
      setActionError("Data set is required.")
      return
    }
    if (Number.isNaN(savedCount) || savedCount < 0) {
      setActionError("Salvations must be 0 or more.")
      return
    }
    if (doorsCount !== null && (Number.isNaN(doorsCount) || doorsCount < 0)) {
      setActionError("Doors knocked must be 0 or more.")
      return
    }
    if (startTime && !normalizeTimeValue(startTime)) {
      setActionError("Start time must be valid.")
      return
    }
    if (endTime && !normalizeTimeValue(endTime)) {
      setActionError("End time must be valid.")
      return
    }
    if (startTime && endTime && getSessionDurationMinutes(startTime, endTime) === null) {
      setActionError("End time must be after the start time.")
      return
    }

    setSessions((prev) =>
      prev.map((session) =>
        session.id === editingLogSessionId
          ? {
              ...session,
              name: editEventName.trim(),
              date: editEventDate,
              dataset_id: datasetId,
              saved_count: savedCount,
              doors_knocked: doorsCount,
              start_time: startTime,
              end_time: endTime,
              notes: editEventNotes.trim(),
              tagIds: editEventTagIds,
            }
          : session,
      ),
    )
    setActionMessage("Event updated.")
    resetLogSessionEdit()
  }

  const toggleEditEventTag = (tagId: string) => {
    setEditEventTagIds((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const handleQuickAddEditEventTag = () => {
    clearMessages()
    const name = logEventTagName
    if (!name) {
      setActionError("Enter a tag name to add.")
      logEventTagInputRef.current?.focus()
      return
    }
    if (logEventTagMatch) {
      setEditEventTagIds((prev) =>
        prev.includes(logEventTagMatch.id) ? prev : [...prev, logEventTagMatch.id],
      )
      setEditEventTagQuery("")
      return
    }
    const color = TAG_COLORS[tags.length % TAG_COLORS.length]
    const created = createTag(name, color)
    if (!created) return
    setEditEventTagIds((prev) => (prev.includes(created.id) ? prev : [...prev, created.id]))
    setEditEventTagQuery("")
  }

  const handleAddSessionPerson = (event?: React.FormEvent) => {
    event?.preventDefault()
    clearMessages()
    const name = sessionPersonName.trim() || "Unknown Name"
    const notes = sessionPersonNotes.trim()
    let highlightId: string | null = null
    if (editingSessionPersonId) {
      setSessionPeople((prev) =>
        prev.map((person) =>
          person.id === editingSessionPersonId
            ? {
                ...person,
                name,
                tagIds: sessionPersonTags,
                role: sessionPersonRole,
                notes,
              }
            : person,
        ),
      )
      highlightId = editingSessionPersonId
      setEditingSessionPersonId(null)
    } else {
      const newId = createLocalId()
      setSessionPeople((prev) => [
        ...prev,
        {
          id: newId,
          name,
          tagIds: sessionPersonTags,
          role: sessionPersonRole,
          notes,
        },
      ])
      highlightId = newId
    }
    setSessionPersonNotice(
      editingSessionPersonId ? "Salvation updated for this event." : "Salvation added to this event.",
    )
    setSessionPersonHighlightId(highlightId)
    setSessionPersonName("")
    setSessionPersonTags([])
    setSessionPersonRole(DEFAULT_SALVATION_ROLE)
    setSessionPersonNotes("")
    setSessionTagQuery("")
  }

  const clearSessionPeople = () => {
    setSessionPeople([])
    setEditingSessionPersonId(null)
    setSessionPersonNotice("")
    setSessionPersonHighlightId(null)
    setSessionPersonName("")
    setSessionPersonTags([])
    setSessionPersonRole(DEFAULT_SALVATION_ROLE)
    setSessionPersonNotes("")
    setSessionTagQuery("")
  }

  const handleEditSessionPerson = (personId: string) => {
    const person = sessionPeople.find((item) => item.id === personId)
    if (!person) return
    setSessionPersonName(person.name === "Unknown Name" ? "" : person.name)
    setSessionPersonTags(person.tagIds)
    setSessionPersonRole(person.role)
    setSessionPersonNotes(person.notes ?? "")
    setSessionTagQuery("")
    setEditingSessionPersonId(personId)
    setSessionSalvationsOpen(true)
  }

  const handleRemoveSessionPerson = (personId: string) => {
    setSessionPeople((prev) => prev.filter((person) => person.id !== personId))
    if (editingSessionPersonId === personId) {
      setEditingSessionPersonId(null)
      setSessionPersonName("")
      setSessionPersonTags([])
      setSessionPersonRole(DEFAULT_SALVATION_ROLE)
      setSessionPersonNotes("")
    }
  }

  const handleCancelSessionEdit = () => {
    setEditingSessionPersonId(null)
    setSessionPersonName("")
    setSessionPersonTags([])
    setSessionPersonRole(DEFAULT_SALVATION_ROLE)
    setSessionPersonNotes("")
  }

  const toggleSessionPersonTag = (tagId: string) => {
    setSessionPersonTags((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const toggleEventTag = (tagId: string) => {
    setDraftEventTagIds((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const handleQuickAddSessionTag = () => {
    clearMessages()
    const name = sessionTagName
    if (!name) {
      setActionError("Enter a tag name to add.")
      sessionTagInputRef.current?.focus()
      return
    }
    if (sessionTagMatch) {
      setSessionPersonTags((prev) =>
        prev.includes(sessionTagMatch.id) ? prev : [...prev, sessionTagMatch.id],
      )
      setSessionTagQuery("")
      return
    }
    const color = TAG_COLORS[tags.length % TAG_COLORS.length]
    const created = createTag(name, color)
    if (!created) return
    setSessionPersonTags((prev) => (prev.includes(created.id) ? prev : [...prev, created.id]))
    setSessionTagQuery("")
  }

  const handleQuickAddEventTag = () => {
    clearMessages()
    const name = eventTagName
    if (!name) {
      setActionError("Enter a tag name to add.")
      eventTagInputRef.current?.focus()
      return
    }
    if (eventTagMatch) {
      setDraftEventTagIds((prev) =>
        prev.includes(eventTagMatch.id) ? prev : [...prev, eventTagMatch.id],
      )
      setEventTagQuery("")
      return
    }
    const color = TAG_COLORS[tags.length % TAG_COLORS.length]
    const created = createTag(name, color)
    if (!created) return
    setDraftEventTagIds((prev) => (prev.includes(created.id) ? prev : [...prev, created.id]))
    setEventTagQuery("")
  }

  const handleAddStandalonePerson = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    const name = standalonePersonName.trim() || "Unknown Name"
    const datasetId = standaloneDatasetId || defaultDatasetId
    const timeSpentValue = standaloneTimeSpent.trim()
    const timeSpent = timeSpentValue ? normalizeMinutesValue(timeSpentValue) : null
    if (!standaloneDate) {
      setActionError("Date is required.")
      return
    }
    if (!datasetId) {
      setActionError("Data set is required.")
      return
    }
    if (timeSpentValue && timeSpent === null) {
      setActionError("Time spent must be 0 or more.")
      return
    }
    const sessionId = createLocalId()
    const createdAt = new Date().toISOString()
    const personRow: PersonRow = {
      id: createLocalId(),
      session_id: sessionId,
      name,
      tagIds: standalonePersonTags,
      role: standalonePersonRole,
      time_spent_minutes: timeSpent,
      notes: standalonePersonNotes.trim(),
      created_at: createdAt,
    }

    const sessionData: SessionRow = {
      id: sessionId,
      date: standaloneDate,
      name: "Standalone salvations",
      type: "standalone",
      dataset_id: datasetId,
      saved_count: 1,
      doors_knocked: null,
      tagIds: [],
      start_time: null,
      end_time: null,
      notes: "",
      created_at: createdAt,
    }

    setSessions((prev) => [sessionData, ...prev])
    setPeople((prev) => [personRow, ...prev])
    setStandalonePersonName("")
    setStandalonePersonTags([])
    setStandalonePersonRole(DEFAULT_SALVATION_ROLE)
    setStandaloneTimeSpent("")
    setStandalonePersonNotes("")
    setStandaloneTagQuery("")
    setActionMessage("Standalone salvation saved.")
  }

  const toggleStandalonePersonTag = (tagId: string) => {
    setStandalonePersonTags((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const handleQuickAddStandaloneTag = () => {
    clearMessages()
    const name = standaloneTagName
    if (!name) {
      setActionError("Enter a tag name to add.")
      standaloneTagInputRef.current?.focus()
      return
    }
    if (standaloneTagMatch) {
      setStandalonePersonTags((prev) =>
        prev.includes(standaloneTagMatch.id) ? prev : [...prev, standaloneTagMatch.id],
      )
      setStandaloneTagQuery("")
      return
    }
    const color = TAG_COLORS[tags.length % TAG_COLORS.length]
    const created = createTag(name, color)
    if (!created) return
    setStandalonePersonTags((prev) => (prev.includes(created.id) ? prev : [...prev, created.id]))
    setStandaloneTagQuery("")
  }

  const handleEditLogPerson = (personId: string) => {
    const person = people.find((item) => item.id === personId)
    if (!person) return
    clearMessages()
    setEditingLogPersonId(personId)
    setLogPersonName(person.name === "Unknown Name" ? "" : person.name)
    setLogPersonTags(person.tagIds ?? [])
    setLogPersonRole(person.role)
    setLogPersonNotes(person.notes ?? "")
    setLogPersonTimeSpent(
      person.time_spent_minutes === null || person.time_spent_minutes === undefined
        ? ""
        : String(person.time_spent_minutes),
    )
    setLogPersonTagQuery("")
  }

  const resetLogPersonEdit = () => {
    setEditingLogPersonId(null)
    setLogPersonName("")
    setLogPersonTags([])
    setLogPersonNotes("")
    setLogPersonTagQuery("")
    setLogPersonRole(DEFAULT_SALVATION_ROLE)
    setLogPersonTimeSpent("")
  }

  const handleCancelLogPersonEdit = () => {
    clearMessages()
    resetLogPersonEdit()
  }

  const handleSaveLogPersonEdit = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    if (!editingLogPersonId) return
    const person = people.find((item) => item.id === editingLogPersonId)
    if (!person) return
    const name = logPersonName.trim() || "Unknown Name"
    const notes = logPersonNotes.trim()
    const session = sessionsById.get(person.session_id)
    const isStandalone = session?.type === "standalone"
    let timeSpent = person.time_spent_minutes ?? null
    if (isStandalone) {
      const timeValue = logPersonTimeSpent.trim()
      const parsed = timeValue ? normalizeMinutesValue(timeValue) : null
      if (timeValue && parsed === null) {
        setActionError("Time spent must be 0 or more.")
        return
      }
      timeSpent = parsed
    }
    setPeople((prev) =>
      prev.map((item) =>
        item.id === editingLogPersonId
          ? {
              ...item,
              name,
              tagIds: logPersonTags,
              role: logPersonRole,
              notes,
              time_spent_minutes: timeSpent,
            }
          : item,
      ),
    )
    setActionMessage("Salvation updated.")
    resetLogPersonEdit()
  }

  const resetLogAddSalvation = () => {
    setLogAddPersonName("")
    setLogAddPersonTags([])
    setLogAddPersonRole(DEFAULT_SALVATION_ROLE)
    setLogAddPersonNotes("")
    setLogAddPersonTagQuery("")
  }

  const handleStartLogAddSalvation = (sessionId: string) => {
    clearMessages()
    setLogAddSessionId(sessionId)
    resetLogAddSalvation()
  }

  const handleCancelLogAddSalvation = () => {
    clearMessages()
    setLogAddSessionId(null)
    resetLogAddSalvation()
  }

  const handleSaveLogAddSalvation = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    if (!logAddSessionId) return
    const session = sessionsById.get(logAddSessionId)
    if (!session) return
    const existingCount = people.filter((person) => person.session_id === logAddSessionId).length
    const remaining = (session.saved_count ?? 0) - existingCount
    if (remaining <= 0) {
      setActionError("No unnamed salvations remaining for this event.")
      return
    }
    const createdAt = new Date().toISOString()
    const personRow: PersonRow = {
      id: createLocalId(),
      session_id: logAddSessionId,
      name: logAddPersonName.trim() || "Unknown Name",
      tagIds: logAddPersonTags,
      role: logAddPersonRole,
      time_spent_minutes: null,
      notes: logAddPersonNotes.trim(),
      created_at: createdAt,
    }
    setPeople((prev) => [personRow, ...prev])
    setActionMessage("Salvation added to event.")
    if (remaining <= 1) {
      setLogAddSessionId(null)
      resetLogAddSalvation()
      return
    }
    resetLogAddSalvation()
  }

  const toggleLogPersonTag = (tagId: string) => {
    setLogPersonTags((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const handleQuickAddLogPersonTag = () => {
    clearMessages()
    const name = logPersonTagName
    if (!name) {
      setActionError("Enter a tag name to add.")
      logPersonTagInputRef.current?.focus()
      return
    }
    if (logPersonTagMatch) {
      setLogPersonTags((prev) =>
        prev.includes(logPersonTagMatch.id) ? prev : [...prev, logPersonTagMatch.id],
      )
      setLogPersonTagQuery("")
      return
    }
    const color = TAG_COLORS[tags.length % TAG_COLORS.length]
    const created = createTag(name, color)
    if (!created) return
    setLogPersonTags((prev) => (prev.includes(created.id) ? prev : [...prev, created.id]))
    setLogPersonTagQuery("")
  }

  const toggleLogAddPersonTag = (tagId: string) => {
    setLogAddPersonTags((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const handleQuickAddLogAddTag = () => {
    clearMessages()
    const name = logAddTagName
    if (!name) {
      setActionError("Enter a tag name to add.")
      logAddTagInputRef.current?.focus()
      return
    }
    if (logAddTagMatch) {
      setLogAddPersonTags((prev) =>
        prev.includes(logAddTagMatch.id) ? prev : [...prev, logAddTagMatch.id],
      )
      setLogAddPersonTagQuery("")
      return
    }
    const color = TAG_COLORS[tags.length % TAG_COLORS.length]
    const created = createTag(name, color)
    if (!created) return
    setLogAddPersonTags((prev) => (prev.includes(created.id) ? prev : [...prev, created.id]))
    setLogAddPersonTagQuery("")
  }

  const toggleTagFilter = (tagId: string) => {
    setSelectedTagIds((prev) =>
      prev.includes(tagId) ? prev.filter((id) => id !== tagId) : [...prev, tagId],
    )
  }

  const toggleDatasetFilter = (datasetId: string) => {
    setSelectedDatasetIds((prev) =>
      prev.includes(datasetId) ? prev.filter((id) => id !== datasetId) : [...prev, datasetId],
    )
  }

  const toggleLogDatasetFilter = (datasetId: string) => {
    setLogDatasetIds((prev) =>
      prev.includes(datasetId) ? prev.filter((id) => id !== datasetId) : [...prev, datasetId],
    )
  }

  const clearLogFilters = () => {
    setLogQuery("")
    setLogDatasetIds([])
  }

  const toggleStatsViewSection = (section: StatsSection) => {
    setStatsViewSections((prev) =>
      prev.includes(section) ? prev.filter((item) => item !== section) : [...prev, section],
    )
  }

  const handleStartStatsViewCreate = () => {
    clearMessages()
    setEditingStatsViewId(null)
    setStatsViewName("")
    setStatsViewSections([])
    setStatsBuilderOpen(true)
  }

  const handleStartStatsViewEdit = (viewId: string) => {
    const viewItem = statsViews.find((view) => view.id === viewId)
    if (!viewItem) return
    clearMessages()
    setEditingStatsViewId(viewItem.id)
    setStatsViewName(viewItem.name)
    setStatsViewSections(viewItem.sections)
    setStatsBuilderOpen(true)
  }

  const handleCancelStatsViewEdit = () => {
    clearMessages()
    setEditingStatsViewId(null)
    setStatsViewName("")
    setStatsViewSections([])
    setStatsBuilderOpen(false)
  }

  const handleSaveStatsView = (event: React.FormEvent) => {
    event.preventDefault()
    clearMessages()
    const name = statsViewName.trim()
    if (!name) {
      setActionError("View name is required.")
      return
    }
    if (!statsViewSections.length) {
      setActionError("Select at least one section.")
      return
    }
    if (
      statsViews.some(
        (view) =>
          view.name.toLowerCase() === name.toLowerCase() && view.id !== editingStatsViewId,
      )
    ) {
      setActionError("A view with that name already exists.")
      return
    }
    if (editingStatsViewId) {
      setStatsViews((prev) =>
        prev.map((view) =>
          view.id === editingStatsViewId ? { ...view, name, sections: statsViewSections } : view,
        ),
      )
      setActiveStatsViewId(editingStatsViewId)
      setActionMessage("Stats view updated.")
    } else {
      const newView: StatsView = {
        id: createLocalId(),
        name,
        sections: statsViewSections,
        created_at: new Date().toISOString(),
      }
      setStatsViews((prev) => [...prev, newView])
      setActiveStatsViewId(newView.id)
      setActionMessage("Stats view saved.")
    }
    setEditingStatsViewId(null)
    setStatsViewName("")
    setStatsViewSections([])
    setStatsBuilderOpen(false)
  }

  const toggleSettingsDetail = (detail: "sessions" | "salvations" | "tags" | "datasets") => {
    setSettingsDetail((prev) => (prev === detail ? null : detail))
  }

  const setQuickRange = (days: number | "all") => {
    if (days === "all") {
      setRangeStart("")
      setRangeEnd("")
      return
    }
    const end = new Date()
    const start = new Date()
    start.setDate(end.getDate() - days + 1)
    setRangeStart(toDateKey(start))
    setRangeEnd(toDateKey(end))
  }

  const exportSessionsCsv = () => {
    const rows = sessionsSorted.map((session) => {
      const namedCount = peopleBySession.get(session.id)?.length ?? 0
      const dataset = datasetById.get(session.dataset_id)
      const tagNames = session.tagIds
        .map((id) => tagsById.get(id)?.name)
        .filter(Boolean)
        .join("|")
      const durationMinutes = getSessionDurationMinutes(session.start_time, session.end_time)
      return {
        date: session.date,
        event_name: session.name,
        event_type: session.type,
        dataset_name: dataset?.name ?? "",
        salvations_reported: session.saved_count,
        doors_knocked: session.doors_knocked ?? "",
        salvations_listed: namedCount,
        event_tags: tagNames,
        start_time: session.start_time ?? "",
        end_time: session.end_time ?? "",
        duration_minutes: durationMinutes ?? "",
        event_notes: session.notes,
      }
    })
    downloadCsv("soulwinning-events.csv", rows)
  }

  const exportPeopleCsv = () => {
    const rows = peopleWithTags.map((person) => {
      const session = sessionsById.get(person.session_id)
      const dataset = session ? datasetById.get(session.dataset_id) : undefined
      const tagNames = person.tagIds
        .map((id) => tagsById.get(id)?.name)
        .filter(Boolean)
        .join("|")
      return {
        event_date: session?.date ?? "",
        event_name: session?.name ?? "",
        event_type: session?.type ?? "",
        dataset_name: dataset?.name ?? "",
        person_name: person.name || "Unknown Name",
        role: formatSalvationRoleLabel(person.role),
        tags: tagNames,
        time_spent_minutes: person.time_spent_minutes ?? "",
        salvation_notes: person.notes,
        event_id: person.session_id,
      }
    })
    downloadCsv("soulwinning-salvations.csv", rows)
  }

  const handleOpenCloudSyncSettings = () => {
    setSettingsDetail(null)
    setSettingsSection("sync")
    setSettingsOpen(true)
  }

  return (
    <div className={`app ${isOffline ? "is-offline" : ""}`}>
      {isOffline && (
        <div className="offline-banner">
          You are currently offline. Changes are saved locally and will sync when connection is restored.
        </div>
      )}
      <header className="app-header">
        <div className="brand">
          <div className="brand__title">Soulwinning Tracker</div>
        </div>
        <div className="header-actions">
          <div className="view-tabs">
            <button
              className={`tab-button ${view === "home" ? "is-active" : ""}`}
              type="button"
              onClick={() => setView("home")}
            >
              Home
            </button>
            <button
              className={`tab-button ${view === "log" ? "is-active" : ""}`}
              type="button"
              onClick={() => setView("log")}
            >
              Log
            </button>
            <button
              className={`tab-button ${view === "goals" ? "is-active" : ""}`}
              type="button"
              onClick={() => setView("goals")}
            >
              Goals
            </button>
            <button
              className={`tab-button ${view === "stats" ? "is-active" : ""}`}
              type="button"
              onClick={() => setView("stats")}
            >
              More Stats
            </button>
          </div>
          <button
            className={`btn btn--compact ${authUserId ? "btn--soft" : "btn--primary"}`}
            type="button"
            onClick={handleOpenCloudSyncSettings}
          >
            {authUserId ? "Account" : "Sign in / Create account"}
          </button>
          <button
            className="btn btn--ghost btn--compact"
            type="button"
            onClick={() => setSettingsOpen((prev) => !prev)}
          >
            Settings
          </button>
        </div>
      </header>

      <main className="app-main">
        {loading && <div className="status status--loading">Working...</div>}
        {actionError && <div className="status status--error">{actionError}</div>}
        {actionMessage && <div className="status status--success">{actionMessage}</div>}

        {view === "home" ? (
          <>
          <div className="content-grid">
                {hasGoals ? (
                  <section className="panel panel--full panel--goal-strip">
                    <div className="goal-strip-header">
                      <span className="goal-strip-title">Goal progress</span>
                      <button
                        className="btn btn--ghost btn--compact"
                        type="button"
                        onClick={() => setView("goals")}
                      >
                        View goals
                      </button>
                    </div>
                    <div className="goal-strip-grid">
                      {activeGoalCards.map((card) => {
                        const percent = Math.min(100, (card.current / card.goal) * 100)
                        return (
                          <div key={card.id} className="goal-strip-row">
                            <div className="goal-strip-label">
                              <div>
                                <span>{card.label}</span>
                                {card.details && card.details !== "All data sets" ? (
                                  <span className="goal-strip-detail">{card.details}</span>
                                ) : null}
                              </div>
                              <span className="goal-strip-period">{card.period}</span>
                            </div>
                            <div className="goal-strip-bar">
                              <div
                                className={`goal-strip-fill ${percent >= 100 ? "is-complete" : ""}`}
                                style={{ width: `${percent}%` }}
                              />
                            </div>
                            <div className="goal-strip-meta">
                              {numberFormatter.format(card.current)} / {numberFormatter.format(card.goal)} (
                              {percent.toFixed(0)}%)
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </section>
                ) : null}
                <section className="panel panel--full">
                  <div className="home-action-grid">
                    <button
                      className="home-action"
                      type="button"
                      onClick={() => openHomePanel("session")}
                    >
                      <div className="home-action__title">Add Soulwinning Time</div>
                      <div className="home-action__desc">
                        Use this for normal soulwinning times. You can add or edit names later in Log.
                      </div>
                    </button>
                    <button
                      className="home-action"
                      type="button"
                      onClick={() => openHomePanel("standalone")}
                    >
                      <div className="home-action__title">Add Other Salvations</div>
                      <div className="home-action__desc">
                        Use for salvations outside a standard soulwinning time (e.g. grocery store). Time
                        spent is optional for time stats.
                      </div>
                    </button>
                  </div>
                </section>
                <section className="panel panel--full">
                  <div className="panel-header">
                    <h2>Recent log</h2>
                    <div className="panel-header__actions">
                      <button
                        className="btn btn--ghost btn--compact"
                        type="button"
                        onClick={() => setView("log")}
                      >
                        Open Log
                      </button>
                    </div>
                  </div>
                  <div className="note">Newest to oldest.</div>
                  <div className="home-log-section">
                    <h3>Logged events</h3>
                    <div className="session-list">
                      {sessionsSorted.length ? (
                        sessionsSorted.map((session) => {
                          const dataset = datasetById.get(session.dataset_id)
                          const datasetLabel = dataset ? dataset.name : "Personal"
                          const sessionLabel =
                            session.name ||
                            (session.type === "standalone"
                              ? "Standalone salvations"
                              : "Soulwinning event")
                          const sessionDate = dateFormatter.format(
                            new Date(`${session.date}T00:00:00`),
                          )
                          const doorsCount =
                            typeof session.doors_knocked === "number" &&
                            Number.isFinite(session.doors_knocked)
                              ? session.doors_knocked
                              : null
                          const showDoors = doorsCount !== null && doorsCount > 0
                          const doorsLabel = showDoors
                            ? numberFormatter.format(doorsCount ?? 0)
                            : ""
                          return (
                            <div key={session.id} className="session-card">
                              <div className="session-summary">
                                <div className="session-summary__title">
                                  <strong>{sessionLabel}</strong>
                                  <span className="muted">{sessionDate}</span>
                                </div>
                                <div className="session-summary__meta">
                                  <span>{formatSalvationCount(session.saved_count)}</span>
                                  {showDoors ? (
                                    <span className="meta-pill meta-pill--compact">
                                      Doors knocked: {doorsLabel}
                                    </span>
                                  ) : null}
                                  <span>{datasetLabel}</span>
                                </div>
                              </div>
                            </div>
                          )
                        })
                      ) : (
                        <div className="empty-state">No events logged yet.</div>
                      )}
                    </div>
                  </div>
                  <div className="home-log-section">
                    <h3>Salvations</h3>
                    <div className="salvation-results">
                      {salvationLogEntries.length || homeUnnamedSalvations.length ? (
                        <>
                          {salvationLogEntries.map(({ person, session, dataset }) => {
                            if (!session) return null
                            const sessionLabel =
                              session.name ||
                              (session.type === "standalone"
                                ? "Standalone salvations"
                                : "Soulwinning event")
                            const sessionDate = dateFormatter.format(
                              new Date(`${session.date}T00:00:00`),
                            )
                            const showTimeSpent =
                              session.type === "standalone" &&
                              person.time_spent_minutes !== null &&
                              person.time_spent_minutes !== undefined
                            return (
                              <div key={person.id} className="salvation-row">
                                <div className="salvation-row__header">
                                  <div>
                                    <strong>{person.name || "Unknown Name"}</strong>
                                    <div className="muted">
                                      {sessionLabel} - {sessionDate}
                                    </div>
                                  </div>
                                  <div className="salvation-row__meta">
                                    <span className="muted">{dataset ? dataset.name : "Personal"}</span>
                                    <span className="meta-pill meta-pill--compact">
                                      {formatSalvationRoleLabel(person.role)}
                                    </span>
                                  </div>
                                </div>
                                <div className="tag-grid">
                                  {person.tagIds.length ? (
                                    person.tagIds.map((tagId) => {
                                      const tag = tagsById.get(tagId)
                                      if (!tag) return null
                                      return (
                                        <span
                                          key={tagId}
                                          className="tag-chip"
                                          style={{
                                            "--tag-color": tag.color ?? TAG_COLORS[0],
                                          } as React.CSSProperties}
                                        >
                                          {tag.name}
                                        </span>
                                      )
                                    })
                                  ) : (
                                    <span className="muted">No tags for this salvation.</span>
                                  )}
                                </div>
                                {person.notes ? <div className="note-text">{person.notes}</div> : null}
                                {showTimeSpent ? (
                                  <div className="muted">
                                    Time spent: {formatDurationMinutes(person.time_spent_minutes)}
                                  </div>
                                ) : null}
                              </div>
                            )
                          })}
                          {homeUnnamedSalvations.map(({ session, dataset, unnamedCount }) => {
                            const sessionLabel =
                              session.name ||
                              (session.type === "standalone"
                                ? "Standalone salvations"
                                : "Soulwinning event")
                            const sessionDate = dateFormatter.format(
                              new Date(`${session.date}T00:00:00`),
                            )
                            const unnamedLabel = `${formatCount(
                              unnamedCount,
                              "unnamed salvation",
                              "unnamed salvations",
                            )} - ${sessionDate} - ${sessionLabel}`
                            return (
                              <div
                                key={`unnamed-home-${session.id}`}
                                className="salvation-row salvation-row--unnamed"
                              >
                                <div className="salvation-row__header">
                                  <div>
                                    <strong>{unnamedLabel}</strong>
                                  </div>
                                  <div className="salvation-row__meta">
                                    <span className="muted">{dataset ? dataset.name : "Personal"}</span>
                                  </div>
                                </div>
                              </div>
                            )
                          })}
                        </>
                      ) : (
                        <div className="empty-state">No salvations logged yet.</div>
                      )}
                    </div>
                  </div>
                </section>
                <section className="panel panel--soft panel--full">
                  <div className="note">Like the app? Send some support!</div>
                  <div className="form-actions">
                    <button className="btn btn--soft btn--compact" type="button" onClick={handleOpenSupport}>
                      Support on Ko-fi
                    </button>
                  </div>
                </section>
              </div>
              {homePanel ? (
                <>
                  <button
                    className="modal-backdrop"
                    type="button"
                    aria-label="Close"
                    onClick={closeHomePanel}
                  />
                  <div className="modal-panel">
                    {homePanel === "session" ? (
                      <section
                        className="panel panel--hero modal-panel__content"
                        role="dialog"
                        aria-modal="true"
                        aria-label="Add Soulwinning Time"
                      >
                        <div className="modal-header">
                          <h2>Add Soulwinning Time</h2>
                          <button
                            className="btn btn--ghost btn--compact"
                            type="button"
                            onClick={closeHomePanel}
                          >
                            Close
                          </button>
                        </div>
                        <div className="note">
                          Use this for normal soulwinning times. You can add or edit names later in Log.
                        </div>
                        <form className="form-grid two-col" id="session-form" onSubmit={handleSaveSession}>
                          <label>
                            Name (optional)
                            <input
                              type="text"
                              value={draftSessionName}
                              onChange={(event) => setDraftSessionName(event.target.value)}
                              placeholder="Example: Sunday Soulwinning"
                            />
                          </label>
                          <label>
                            Data set
                            <select
                              value={draftDatasetId}
                              onChange={(event) => setDraftDatasetId(event.target.value)}
                            >
                              {datasets.map((dataset) => (
                                <option key={dataset.id} value={dataset.id}>
                                  {dataset.name}
                                </option>
                              ))}
                            </select>
                          </label>
                          <label>
                            Date
                            <input
                              type="date"
                              value={draftDate}
                              onChange={(event) => setDraftDate(event.target.value)}
                              required
                            />
                          </label>
                          <label>
                            Start time (optional)
                            <input
                              type="time"
                              value={draftStartTime}
                              onChange={(event) => setDraftStartTime(event.target.value)}
                            />
                          </label>
                          <label>
                            End time (optional)
                            <input
                              type="time"
                              value={draftEndTime}
                              onChange={(event) => setDraftEndTime(event.target.value)}
                            />
                          </label>
                          <label>
                            Salvations
                            <input
                              type="number"
                              min="0"
                              value={draftSavedCount}
                              onChange={(event) => setDraftSavedCount(event.target.value)}
                              required
                            />
                          </label>
                          <label>
                            Doors knocked (optional)
                            <input
                              type="number"
                              min="0"
                              value={draftDoors}
                              onChange={(event) => setDraftDoors(event.target.value)}
                            />
                          </label>
                          <label className="form-span">
                            Notes (optional)
                            <textarea
                              className="textarea-compact"
                              value={draftSessionNotes}
                              onChange={(event) => setDraftSessionNotes(event.target.value)}
                              placeholder="Add any details you want to remember about this event"
                            />
                          </label>
                          <div className="form-span">
                            <div className="field-block">
                              <span>Event tags (optional)</span>
                              <div className="tag-grid">
                                {eventTagOptions.length ? (
                                  eventTagOptions.map((tag) => (
                                    <label
                                      key={tag.id}
                                      className={`tag-pill ${draftEventTagIds.includes(tag.id) ? "is-active" : ""}`}
                                      style={{ "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                                    >
                                      <input
                                        type="checkbox"
                                        checked={draftEventTagIds.includes(tag.id)}
                                        onChange={() => toggleEventTag(tag.id)}
                                      />
                                      {tag.name}
                                    </label>
                                  ))
                                ) : (
                                  <span className="muted">
                                    {tags.length ? "No tags match that search." : "Create tags to assign them here."}
                                  </span>
                                )}
                              </div>
                              {tags.length > TOP_TAG_LIMIT && !eventTagQuery.trim() ? (
                                <div className="note tag-hint">Showing top tags. Search to find more.</div>
                              ) : null}
                            </div>
                            <label style={{ marginTop: "12px" }}>
                              Tag search
                              <input
                                ref={eventTagInputRef}
                                type="text"
                                value={eventTagQuery}
                                onChange={(event) => setEventTagQuery(event.target.value)}
                                placeholder="Search tags"
                              />
                            </label>
                            <button
                              className="btn btn--soft btn--compact"
                              type="button"
                              onClick={handleQuickAddEventTag}
                            >
                              {eventTagActionLabel}
                            </button>
                          </div>
                        </form>
                        <div className="form-actions" style={{ marginTop: "16px" }}>
                          <button
                            className="btn btn--soft"
                            type="button"
                            onClick={() => setSessionSalvationsOpen((prev) => !prev)}
                          >
                            {sessionSalvationsOpen
                              ? "Hide Salvations for This Event"
                              : "Add Salvations to This Event"}
                          </button>
                        </div>
                        {sessionSalvationsOpen ? (
                          <div className="subpanel">
                            <div className="note">
                              Add names and tags for this event. They attach when you save the event.
                            </div>
                            <form className="form-grid two-col" onSubmit={handleAddSessionPerson}>
                              <label className="form-span">
                                Name (optional)
                                <input
                                  type="text"
                                  value={sessionPersonName}
                                  onChange={(event) => setSessionPersonName(event.target.value)}
                                  placeholder="Name of person saved (if known)"
                                />
                              </label>
                              <label>
                                Role
                                <select
                                  value={sessionPersonRole}
                                  onChange={(event) =>
                                    setSessionPersonRole(event.target.value as SalvationRole)
                                  }
                                >
                                  <option value="presenter">Preacher</option>
                                  <option value="partner">Silent partner</option>
                                </select>
                              </label>
                              <label>
                                Tag search
                                <input
                                  ref={sessionTagInputRef}
                                  type="text"
                                  value={sessionTagQuery}
                                  onChange={(event) => setSessionTagQuery(event.target.value)}
                                  placeholder="Search tags"
                                />
                              </label>
                              <div className="field-block">
                                <span>Tags</span>
                                <div className="tag-grid">
                                  {sessionTagOptions.length ? (
                                    sessionTagOptions.map((tag) => (
                                      <label
                                        key={tag.id}
                                        className={`tag-pill ${
                                          sessionPersonTags.includes(tag.id) ? "is-active" : ""
                                        }`}
                                        style={{ "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                                      >
                                        <input
                                          type="checkbox"
                                          checked={sessionPersonTags.includes(tag.id)}
                                          onChange={() => toggleSessionPersonTag(tag.id)}
                                        />
                                        {tag.name}
                                      </label>
                                    ))
                                  ) : (
                                    <span className="muted">
                                      {tags.length ? "No tags match that search." : "Create tags to assign them here."}
                                    </span>
                                  )}
                                  <button
                                    className="tag-pill tag-pill--add"
                                    type="button"
                                    onClick={handleQuickAddSessionTag}
                                  >
                                    {sessionTagActionLabel}
                                  </button>
                                </div>
                                {tags.length > TOP_TAG_LIMIT && !sessionTagQuery.trim() ? (
                                  <div className="note tag-hint">Showing top tags. Search to find more.</div>
                                ) : null}
                              </div>
                              <label className="form-span">
                                Notes (optional)
                                <textarea
                                  className="textarea-compact"
                                  value={sessionPersonNotes}
                                  onChange={(event) => setSessionPersonNotes(event.target.value)}
                                  placeholder="Optional details about this salvation"
                                />
                              </label>
                              <div className="form-actions">
                                <button className="btn btn--primary" type="submit">
                                  {editingSessionPersonId ? "Save changes" : "Add salvation"}
                                </button>
                                {editingSessionPersonId ? (
                                  <button className="btn btn--ghost" type="button" onClick={handleCancelSessionEdit}>
                                    Cancel edit
                                  </button>
                                ) : null}
                              </div>
                            </form>
                          </div>
                        ) : null}
                        <div className="pending-list">
                          <div className="pending-header">
                            <strong>Salvations during this event</strong>
                            <span className="muted">{sessionPeople.length} total</span>
                          </div>
                          <div className="note">These salvations attach when you save the event.</div>
                          {sessionPersonNotice ? (
                            <div className="status status--success status--inline">{sessionPersonNotice}</div>
                          ) : null}
                          {sessionPeople.length ? (
                            <div className="pending-grid">
                              {sessionPeople.map((person) => (
                                <div
                                  key={person.id}
                                  className={`pending-card ${
                                    editingSessionPersonId === person.id ? "is-editing" : ""
                                  } ${sessionPersonHighlightId === person.id ? "is-highlighted" : ""}`}
                                >
                                  <div className="pending-card__header">
                                    <div className="pending-card__title">
                                      <strong>{person.name || "Unknown Name"}</strong>
                                      <span className="meta-pill meta-pill--compact">
                                        {formatSalvationRoleLabel(person.role)}
                                      </span>
                                    </div>
                                    <div className="pending-card__actions">
                                      <button
                                        className="btn btn--ghost btn--compact"
                                        type="button"
                                        onClick={() => handleEditSessionPerson(person.id)}
                                      >
                                        Edit
                                      </button>
                                      <button
                                        className="btn btn--danger btn--compact"
                                        type="button"
                                        onClick={() => handleRemoveSessionPerson(person.id)}
                                      >
                                        Remove
                                      </button>
                                    </div>
                                  </div>
                                  {person.notes ? <div className="note-text">{person.notes}</div> : null}
                                  <div className="tag-grid">
                                    {person.tagIds.length ? (
                                      person.tagIds.map((tagId) => {
                                        const tag = tagsById.get(tagId)
                                        if (!tag) return null
                                        return (
                                          <span
                                            key={tagId}
                                            className="tag-chip"
                                            style={{
                                              "--tag-color": tag.color ?? TAG_COLORS[0],
                                            } as React.CSSProperties}
                                          >
                                            {tag.name}
                                          </span>
                                        )
                                      })
                                    ) : (
                                      <span className="muted">No tags</span>
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          ) : (
                            <div className="empty-state">
                              You have not logged any salvations for this event.
                            </div>
                          )}
                          {sessionPeople.length ? (
                            <div className="form-actions">
                              <button className="btn btn--ghost" type="button" onClick={clearSessionPeople}>
                                Clear all salvations
                              </button>
                            </div>
                          ) : null}
                        </div>
                        <div className="form-actions" style={{ marginTop: "16px" }}>
                          <button className="btn btn--primary" type="submit" form="session-form">
                            Save event
                          </button>
                          <button className="btn btn--ghost" type="button" onClick={resetSessionDraft}>
                            Reset form
                          </button>
                        </div>
                      </section>
                    ) : (
                      <section
                        className="panel modal-panel__content"
                        role="dialog"
                        aria-modal="true"
                        aria-label="Add Other Salvations"
                      >
                        <div className="modal-header">
                          <h2>Add Other Salvations</h2>
                          <button
                            className="btn btn--ghost btn--compact"
                            type="button"
                            onClick={closeHomePanel}
                          >
                            Close
                          </button>
                        </div>
                        <div className="note">
                          Use for salvations outside a standard soulwinning time (e.g. grocery store). Time
                          spent is optional for time stats.
                        </div>
                        <form className="form-grid" onSubmit={handleAddStandalonePerson}>
                          <label>
                            Date
                            <input
                              type="date"
                              value={standaloneDate}
                              onChange={(event) => setStandaloneDate(event.target.value)}
                              required
                            />
                          </label>
                          <label>
                            Data set
                            <select
                              value={standaloneDatasetId}
                              onChange={(event) => setStandaloneDatasetId(event.target.value)}
                            >
                              {datasets.map((dataset) => (
                                <option key={dataset.id} value={dataset.id}>
                                  {dataset.name}
                                </option>
                              ))}
                            </select>
                          </label>
                          <label>
                            Name (optional)
                            <input
                              type="text"
                              value={standalonePersonName}
                              onChange={(event) => setStandalonePersonName(event.target.value)}
                              placeholder="Name of person saved (if known)"
                            />
                          </label>
                          <label>
                            Role
                            <select
                              value={standalonePersonRole}
                              onChange={(event) =>
                                setStandalonePersonRole(event.target.value as SalvationRole)
                              }
                            >
                              <option value="presenter">Preacher</option>
                              <option value="partner">Silent partner</option>
                            </select>
                          </label>
                          <label>
                            Time spent (minutes, optional)
                            <input
                              type="number"
                              min="0"
                              value={standaloneTimeSpent}
                              onChange={(event) => setStandaloneTimeSpent(event.target.value)}
                              placeholder="e.g. 15"
                            />
                          </label>
                          <label>
                            Tag search
                            <input
                              ref={standaloneTagInputRef}
                              type="text"
                              value={standaloneTagQuery}
                              onChange={(event) => setStandaloneTagQuery(event.target.value)}
                              placeholder="Search tags"
                            />
                          </label>
                          <label className="form-span">
                            Notes (optional)
                            <textarea
                              className="textarea-compact"
                              value={standalonePersonNotes}
                              onChange={(event) => setStandalonePersonNotes(event.target.value)}
                              placeholder="Optional details about this salvation"
                            />
                          </label>
                          <div className="form-span">
                            <div className="muted">Tags</div>
                            <div className="tag-grid">
                              {standaloneTagOptions.length ? (
                                standaloneTagOptions.map((tag) => (
                                  <label
                                    key={tag.id}
                                    className={`tag-pill ${
                                      standalonePersonTags.includes(tag.id) ? "is-active" : ""
                                    }`}
                                    style={{ "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                                  >
                                    <input
                                      type="checkbox"
                                      checked={standalonePersonTags.includes(tag.id)}
                                      onChange={() => toggleStandalonePersonTag(tag.id)}
                                    />
                                    {tag.name}
                                  </label>
                                ))
                              ) : (
                                <span className="muted">
                                  {tags.length ? "No tags match that search." : "Create tags to assign them here."}
                                </span>
                              )}
                              <button
                                className="tag-pill tag-pill--add"
                                type="button"
                                onClick={handleQuickAddStandaloneTag}
                              >
                                {standaloneTagActionLabel}
                              </button>
                            </div>
                            {tags.length > TOP_TAG_LIMIT && !standaloneTagQuery.trim() ? (
                              <div className="note tag-hint">Showing top tags. Search to find more.</div>
                            ) : null}
                          </div>
                          <div className="form-actions">
                            <button className="btn btn--primary" type="submit">
                              Save salvation
                            </button>
                          </div>
                        </form>
                      </section>
                    )}
                  </div>
                </>
              ) : null}
            </>
        ) : view === "goals" ? (
          <div className="content-grid">
                <section className="panel panel--full">
                  <div className="panel-header">
                    <h2>Your goals</h2>
                    <div className="panel-header__actions">
                      {goals.length ? (
                        <button
                          className="btn btn--ghost btn--compact"
                          type="button"
                          onClick={clearGoals}
                        >
                          Clear all goals
                        </button>
                      ) : null}
                    </div>
                  </div>
                  <div className="note">
                    Build custom goals with criteria and track progress here and on Home. Edit any goal after
                    saving.
                  </div>
                  {goalCards.length ? (
                    <div className="goal-grid">
                      {goalCards.map((card) => {
                        const goalDisabled = card.goal <= 0
                        const percent = goalDisabled
                          ? 0
                          : Math.min(100, (card.current / card.goal) * 100)
                        return (
                          <div
                            key={card.id}
                            className={`goal-card ${goalDisabled ? "goal-card--disabled" : ""} ${
                              editingGoalId === card.id ? "is-editing" : ""
                            }`}
                          >
                            <div className="goal-card__header">
                              <strong>{card.label}</strong>
                              <span className="muted">{card.period}</span>
                            </div>
                            {card.details ? (
                              <div className="goal-card__details">{card.details}</div>
                            ) : null}
                            <div className="goal-bar">
                              <div
                                className={`goal-bar__fill ${percent >= 100 ? "is-complete" : ""}`}
                                style={{ width: `${percent}%` }}
                              />
                            </div>
                            <div className="goal-meta">
                              {goalDisabled ? (
                                <span>Goal not set</span>
                              ) : (
                                <span>
                                  {numberFormatter.format(card.current)} /{" "}
                                  {numberFormatter.format(card.goal)}
                                </span>
                              )}
                              <span>{goalDisabled ? "n/a" : `${percent.toFixed(0)}%`}</span>
                            </div>
                            <div className="goal-card__actions">
                              <button
                                className="btn btn--ghost btn--compact"
                                type="button"
                                onClick={() => handleEditGoal(card.id)}
                              >
                                Edit
                              </button>
                              <button
                                className="btn btn--ghost btn--compact"
                                type="button"
                                onClick={() => handleDeleteGoal(card.id)}
                              >
                                Delete
                              </button>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  ) : (
                    <div className="empty-state">No goals yet. Create one below.</div>
                  )}
                </section>

                <section className="panel">
                  <h2>Create a goal</h2>
                  <div className="note">Choose a metric, timeframe, and optional filters.</div>
                  <form className="form-grid two-col" onSubmit={handleSaveGoal}>
                    <label className="form-span">
                      Goal name (optional)
                      <input
                        type="text"
                        value={goalName}
                        onChange={(event) => setGoalName(event.target.value)}
                        placeholder="Example: Salvations this year"
                      />
                    </label>
                    <label>
                      Metric
                      <select
                        value={goalMetric}
                        onChange={(event) => setGoalMetric(event.target.value as GoalMetric)}
                      >
                        <option value="sessions">Events</option>
                        <option value="salvations">Salvations</option>
                        <option value="doors">Doors knocked</option>
                      </select>
                    </label>
                    <label>
                      Time period
                      <select
                        value={goalPeriod}
                        onChange={(event) => setGoalPeriod(event.target.value as GoalPeriod)}
                      >
                        <option value="week">Week</option>
                        <option value="month">Month</option>
                        <option value="year">Year</option>
                        <option value="overall">Overall (all time)</option>
                        <option value="last_week">Last week</option>
                        <option value="last_month">Last month</option>
                        <option value="last_year">Last year</option>
                        <option value="custom">Custom range</option>
                      </select>
                    </label>
                    {goalPeriod === "custom" ? (
                      <>
                        <label>
                          Start date
                          <input
                            type="date"
                            value={goalStartDate}
                            onChange={(event) => setGoalStartDate(event.target.value)}
                            required
                          />
                        </label>
                        <label>
                          End date
                          <input
                            type="date"
                            value={goalEndDate}
                            onChange={(event) => setGoalEndDate(event.target.value)}
                            required
                          />
                        </label>
                      </>
                    ) : null}
                    <label>
                      Target
                      <input
                        type="number"
                        min="1"
                        value={goalTarget}
                        onChange={(event) => setGoalTarget(event.target.value)}
                      />
                    </label>
                    <label>
                      Data set
                      <select
                        value={goalDatasetId}
                        onChange={(event) => setGoalDatasetId(event.target.value)}
                      >
                        <option value="all">All data sets</option>
                        {datasets.map((dataset) => (
                          <option key={dataset.id} value={dataset.id}>
                            {dataset.name}
                          </option>
                        ))}
                      </select>
                    </label>
                    {goalMetric === "salvations" ? (
                      <div className="form-span">
                        <div className="goal-tags">
                          <label>
                            Tag search
                            <input
                              type="text"
                              value={goalTagQuery}
                              onChange={(event) => setGoalTagQuery(event.target.value)}
                              placeholder="Search tags"
                            />
                          </label>
                          <div className="field-block">
                            <span>Tags</span>
                            <div className="tag-grid">
                              {goalTagOptions.length ? (
                                goalTagOptions.map((tag) => (
                                  <label
                                    key={tag.id}
                                    className={`tag-pill ${goalTagIds.includes(tag.id) ? "is-active" : ""}`}
                                    style={{ "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                                  >
                                    <input
                                      type="checkbox"
                                      checked={goalTagIds.includes(tag.id)}
                                      onChange={() => toggleGoalTag(tag.id)}
                                    />
                                    {tag.name}
                                  </label>
                                ))
                              ) : (
                                <span className="muted">
                                  {tags.length ? "No tags match that search." : "Create tags to assign them here."}
                                </span>
                              )}
                            </div>
                            {tags.length > TOP_TAG_LIMIT && !goalTagQuery.trim() ? (
                              <div className="note tag-hint">Showing top tags. Search to find more.</div>
                            ) : null}
                          </div>
                        </div>
                        {goalTagIds.length ? (
                          <label style={{ marginTop: "12px" }}>
                            Tag match
                            <select
                              value={goalTagMatchMode}
                              onChange={(event) => setGoalTagMatchMode(event.target.value as TagMatchMode)}
                            >
                              <option value="any">Any selected tags</option>
                              <option value="all">All selected tags</option>
                            </select>
                          </label>
                        ) : null}
                      </div>
                    ) : (
                      <div className="form-span muted">Tags are available for salvation goals.</div>
                    )}
                    <div className="form-actions">
                      <button className="btn btn--primary" type="submit">
                        {editingGoalId ? "Save changes" : "Add goal"}
                      </button>
                      {editingGoalId ? (
                        <button
                          className="btn btn--ghost"
                          type="button"
                          onClick={handleCancelGoalEdit}
                        >
                          Cancel edit
                        </button>
                      ) : null}
                    </div>
                  </form>
                </section>
              </div>
        ) : view === "log" ? (
          <div className="content-grid">
                <section className="panel panel--full">
                  <h2>Filters</h2>
                  <div className="note">
                    View every event and salvation. Edit events, add names to unnamed salvations, or search by
                    name or notes.
                  </div>
                  <div className="tag-grid" style={{ marginTop: "12px" }}>
                    <button
                      type="button"
                      className={`tag-pill ${logFilter === "sessions" ? "is-active" : ""}`}
                      onClick={() => setLogFilter("sessions")}
                    >
                      Events
                    </button>
                    <button
                      type="button"
                      className={`tag-pill ${logFilter === "salvations" ? "is-active" : ""}`}
                      onClick={() => setLogFilter("salvations")}
                    >
                      Salvations
                    </button>
                    <button
                      type="button"
                      className={`tag-pill ${logFilter === "both" ? "is-active" : ""}`}
                      onClick={() => setLogFilter("both")}
                    >
                      Both
                    </button>
                  </div>
                  <div className="form-grid" style={{ marginTop: "12px" }}>
                    <label>
                      Search
                      <input
                        type="text"
                        value={logQuery}
                        onChange={(event) => setLogQuery(event.target.value)}
                        placeholder="Search events or salvations"
                      />
                    </label>
                    <div className="form-span">
                      <div className="muted">Filter by data set</div>
                      <div className="tag-grid" style={{ marginTop: "8px" }}>
                        {datasets.length ? (
                          datasets.map((dataset) => (
                            <button
                              key={dataset.id}
                              type="button"
                              className={`tag-pill ${logDatasetIds.includes(dataset.id) ? "is-active" : ""}`}
                              onClick={() => toggleLogDatasetFilter(dataset.id)}
                            >
                              {dataset.name}
                            </button>
                          ))
                        ) : (
                          <span className="muted">Create data sets to filter the log.</span>
                        )}
                      </div>
                    </div>
                    <div className="form-actions">
                      <button
                        className="btn btn--ghost"
                        type="button"
                        onClick={clearLogFilters}
                        disabled={!logQuery.trim() && !logDatasetIds.length}
                      >
                        Clear log filters
                      </button>
                    </div>
                  </div>
                </section>

                {(logFilter === "sessions" || logFilter === "both") && (
                  <section className="panel panel--full">
                    <h2>Logged events</h2>
                    <div className="session-list">
                      {logSessions.length ? (
                        logSessions.map((session) => {
                          const peopleInSession = peopleBySession.get(session.id) ?? []
                          const dataset = datasetById.get(session.dataset_id)
                          const datasetLabel = dataset ? dataset.name : "Personal"
                          const sessionLabel =
                            session.name ||
                            (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
                          const sessionDate = dateFormatter.format(
                            new Date(`${session.date}T00:00:00`),
                          )
                          const doorsCount =
                            typeof session.doors_knocked === "number" &&
                            Number.isFinite(session.doors_knocked)
                              ? session.doors_knocked
                              : null
                          const showDoors = doorsCount !== null && doorsCount > 0
                          const unnamedCount = Math.max(
                            0,
                            (session.saved_count ?? 0) - peopleInSession.length,
                          )
                          const doorsLabel = showDoors
                            ? numberFormatter.format(doorsCount ?? 0)
                            : ""
                          const startLabel = formatTimeValue(session.start_time)
                          const endLabel = formatTimeValue(session.end_time)
                          const durationMinutes = getSessionDurationMinutes(
                            session.start_time,
                            session.end_time,
                          )
                          const durationLabel =
                            durationMinutes !== null ? formatDurationMinutes(durationMinutes) : ""
                          let timeLabel = ""
                          if (startLabel && endLabel) {
                            timeLabel = `${startLabel} - ${endLabel}`
                          } else if (startLabel) {
                            timeLabel = `Starts ${startLabel}`
                          } else if (endLabel) {
                            timeLabel = `Ends ${endLabel}`
                          }
                          const timeMeta = timeLabel
                            ? durationLabel
                              ? `${timeLabel} (${durationLabel})`
                              : timeLabel
                            : ""
                          return (
                            <details
                              key={session.id}
                              className="session-card"
                              open={editingLogSessionId === session.id || logAddSessionId === session.id}
                            >
                              <summary>
                                <div className="session-summary">
                                  <div className="session-summary__title">
                                    <strong>{sessionLabel}</strong>
                                    <span className="muted">{sessionDate}</span>
                                  </div>
                                  <div className="session-summary__meta">
                                    <span>{formatSalvationCount(session.saved_count)}</span>
                                    {showDoors ? (
                                      <span className="meta-pill meta-pill--compact">
                                        Doors knocked: {doorsLabel}
                                      </span>
                                    ) : null}
                                    <span>{datasetLabel}</span>
                                    <span className="muted">View salvations</span>
                                  </div>
                                </div>
                              </summary>
                              <div className="session-meta">
                                <span>{formatSalvationCountLabel(session.saved_count)}</span>
                                {showDoors ? (
                                  <span className="meta-pill meta-pill--compact">
                                    Doors knocked: {doorsLabel}
                                  </span>
                                ) : null}
                                {timeMeta ? (
                                  <span className="meta-pill meta-pill--compact">Time: {timeMeta}</span>
                                ) : null}
                                <span>Data set: {datasetLabel}</span>
                              </div>
                              {session.notes ? (
                                <div className="session-notes">
                                  <strong>Event notes</strong>
                                  <div className="note-text">{session.notes}</div>
                                </div>
                              ) : null}
                              {session.tagIds.length ? (
                                <div className="session-tags">
                                  <div className="muted">Event tags</div>
                                  <div className="tag-grid">
                                    {session.tagIds.map((tagId) => {
                                      const tag = tagsById.get(tagId)
                                      if (!tag) return null
                                      return (
                                        <span
                                          key={tagId}
                                          className="tag-chip"
                                          style={{
                                            "--tag-color": tag.color ?? TAG_COLORS[0],
                                          } as React.CSSProperties}
                                        >
                                          {tag.name}
                                        </span>
                                      )
                                    })}
                                  </div>
                                </div>
                              ) : null}
                              <div className="session-people">
                                {peopleInSession.length ? (
                                  peopleInSession.map((person) => (
                                    <div key={person.id} className="session-person">
                                      <strong>{person.name || "Unknown Name"}</strong>
                                      <span className="meta-pill meta-pill--compact">
                                        {formatSalvationRoleLabel(person.role)}
                                      </span>
                                      {person.tagIds.map((tagId) => {
                                        const tag = tagsById.get(tagId)
                                        if (!tag) return null
                                        return (
                                          <span
                                            key={tagId}
                                            className="tag-chip"
                                            style={{
                                              "--tag-color": tag.color ?? TAG_COLORS[0],
                                            } as React.CSSProperties}
                                          >
                                            {tag.name}
                                          </span>
                                        )
                                      })}
                                      {person.notes ? (
                                        <div className="person-note">{person.notes}</div>
                                      ) : null}
                                    </div>
                                  ))
                                ) : (
                                  <div className="muted">No salvations listed for this event.</div>
                                )}
                              </div>
                              {unnamedCount > 0 ? (
                                <div className="subpanel">
                                  <div className="pending-header">
                                    <strong>Unnamed salvations remaining</strong>
                                    <span className="muted">{numberFormatter.format(unnamedCount)}</span>
                                  </div>
                                  <div className="note">
                                    Add a name to replace one of the unnamed salvations for this event.
                                  </div>
                                  {logAddSessionId === session.id ? (
                                    <form className="form-grid log-edit-form" onSubmit={handleSaveLogAddSalvation}>
                                      <label>
                                        Name (optional)
                                        <input
                                          type="text"
                                          value={logAddPersonName}
                                          onChange={(event) => setLogAddPersonName(event.target.value)}
                                          placeholder="Name of person saved (if known)"
                                        />
                                      </label>
                                      <label>
                                        Role
                                        <select
                                          value={logAddPersonRole}
                                          onChange={(event) =>
                                            setLogAddPersonRole(event.target.value as SalvationRole)
                                          }
                                        >
                                          <option value="presenter">Preacher</option>
                                          <option value="partner">Silent partner</option>
                                        </select>
                                      </label>
                                      <label>
                                        Tag search
                                        <input
                                          ref={logAddTagInputRef}
                                          type="text"
                                          value={logAddPersonTagQuery}
                                          onChange={(event) => setLogAddPersonTagQuery(event.target.value)}
                                          placeholder="Search tags"
                                        />
                                      </label>
                                      <div className="form-span">
                                        <div className="muted">Tags</div>
                                        <div className="tag-grid">
                                          {logAddTagOptions.length ? (
                                            logAddTagOptions.map((tag) => (
                                              <label
                                                key={tag.id}
                                                className={`tag-pill ${
                                                  logAddPersonTags.includes(tag.id) ? "is-active" : ""
                                                }`}
                                                style={
                                                  { "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties
                                                }
                                              >
                                                <input
                                                  type="checkbox"
                                                  checked={logAddPersonTags.includes(tag.id)}
                                                  onChange={() => toggleLogAddPersonTag(tag.id)}
                                                />
                                                {tag.name}
                                              </label>
                                            ))
                                          ) : (
                                            <span className="muted">
                                              {tags.length
                                                ? "No tags match that search."
                                                : "Create tags to assign them here."}
                                            </span>
                                          )}
                                          <button
                                            className="tag-pill tag-pill--add"
                                            type="button"
                                            onClick={handleQuickAddLogAddTag}
                                          >
                                            {logAddTagActionLabel}
                                          </button>
                                        </div>
                                        {tags.length > TOP_TAG_LIMIT && !logAddPersonTagQuery.trim() ? (
                                          <div className="note tag-hint">
                                            Showing top tags. Search to find more.
                                          </div>
                                        ) : null}
                                      </div>
                                      <label className="form-span">
                                        Notes (optional)
                                        <textarea
                                          className="textarea-compact"
                                          value={logAddPersonNotes}
                                          onChange={(event) => setLogAddPersonNotes(event.target.value)}
                                          placeholder="Optional details about this salvation"
                                        />
                                      </label>
                                      <div className="form-actions">
                                        <button className="btn btn--primary" type="submit">
                                          Add name
                                        </button>
                                        <button
                                          className="btn btn--ghost"
                                          type="button"
                                          onClick={handleCancelLogAddSalvation}
                                        >
                                          Cancel
                                        </button>
                                      </div>
                                    </form>
                                  ) : (
                                    <div className="form-actions">
                                      <button
                                        className="btn btn--primary btn--compact"
                                        type="button"
                                        onClick={() => handleStartLogAddSalvation(session.id)}
                                      >
                                        Add name
                                      </button>
                                    </div>
                                  )}
                                </div>
                              ) : null}
                              {editingLogSessionId === session.id ? (
                                <form className="form-grid two-col log-edit-form" onSubmit={handleSaveLogSessionEdit}>
                                  <label>
                                    Name (optional)
                                    <input
                                      type="text"
                                      value={editEventName}
                                      onChange={(event) => setEditEventName(event.target.value)}
                                      placeholder="Example: Sunday Soulwinning"
                                    />
                                  </label>
                                  <label>
                                    Data set
                                    <select
                                      value={editEventDatasetId}
                                      onChange={(event) => setEditEventDatasetId(event.target.value)}
                                    >
                                      {datasets.map((dataset) => (
                                        <option key={dataset.id} value={dataset.id}>
                                          {dataset.name}
                                        </option>
                                      ))}
                                    </select>
                                  </label>
                                  <label>
                                    Date
                                    <input
                                      type="date"
                                      value={editEventDate}
                                      onChange={(event) => setEditEventDate(event.target.value)}
                                      required
                                    />
                                  </label>
                                  <label>
                                    Start time (optional)
                                    <input
                                      type="time"
                                      value={editEventStartTime}
                                      onChange={(event) => setEditEventStartTime(event.target.value)}
                                    />
                                  </label>
                                  <label>
                                    End time (optional)
                                    <input
                                      type="time"
                                      value={editEventEndTime}
                                      onChange={(event) => setEditEventEndTime(event.target.value)}
                                    />
                                  </label>
                                  <label>
                                    Salvations
                                    <input
                                      type="number"
                                      min="0"
                                      value={editEventSavedCount}
                                      onChange={(event) => setEditEventSavedCount(event.target.value)}
                                      required
                                    />
                                  </label>
                                  <label>
                                    Doors knocked (optional)
                                    <input
                                      type="number"
                                      min="0"
                                      value={editEventDoors}
                                      onChange={(event) => setEditEventDoors(event.target.value)}
                                    />
                                  </label>
                                  <label className="form-span">
                                    Notes (optional)
                                    <textarea
                                      className="textarea-compact"
                                      value={editEventNotes}
                                      onChange={(event) => setEditEventNotes(event.target.value)}
                                      placeholder="Add any details you want to remember about this event"
                                    />
                                  </label>
                                  <div className="form-span">
                                    <div className="field-block">
                                      <span>Event tags (optional)</span>
                                      <div className="tag-grid">
                                        {logEventTagOptions.length ? (
                                          logEventTagOptions.map((tag) => (
                                            <label
                                              key={tag.id}
                                              className={`tag-pill ${
                                                editEventTagIds.includes(tag.id) ? "is-active" : ""
                                              }`}
                                              style={
                                                { "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties
                                              }
                                            >
                                              <input
                                                type="checkbox"
                                                checked={editEventTagIds.includes(tag.id)}
                                                onChange={() => toggleEditEventTag(tag.id)}
                                              />
                                              {tag.name}
                                            </label>
                                          ))
                                        ) : (
                                          <span className="muted">
                                            {tags.length
                                              ? "No tags match that search."
                                              : "Create tags to assign them here."}
                                          </span>
                                        )}
                                      </div>
                                      {tags.length > TOP_TAG_LIMIT && !editEventTagQuery.trim() ? (
                                        <div className="note tag-hint">
                                          Showing top tags. Search to find more.
                                        </div>
                                      ) : null}
                                    </div>
                                    <label style={{ marginTop: "12px" }}>
                                      Tag search
                                      <input
                                        ref={logEventTagInputRef}
                                        type="text"
                                        value={editEventTagQuery}
                                        onChange={(event) => setEditEventTagQuery(event.target.value)}
                                        placeholder="Search tags"
                                      />
                                    </label>
                                    <button
                                      className="btn btn--soft btn--compact"
                                      type="button"
                                      onClick={handleQuickAddEditEventTag}
                                    >
                                      {logEventTagActionLabel}
                                    </button>
                                  </div>
                                  <div className="form-actions">
                                    <button className="btn btn--primary" type="submit">
                                      Save changes
                                    </button>
                                    <button
                                      className="btn btn--ghost"
                                      type="button"
                                      onClick={handleCancelLogSessionEdit}
                                    >
                                      Cancel edit
                                    </button>
                                    <button
                                      className="btn btn--danger"
                                      type="button"
                                      onClick={() => handleDeleteSession(session.id)}
                                    >
                                      Delete event
                                    </button>
                                  </div>
                                </form>
                              ) : (
                                <div className="form-actions">
                                  <button
                                    className="btn btn--ghost"
                                    type="button"
                                    onClick={() => handleEditLogSession(session.id)}
                                  >
                                    Edit event
                                  </button>
                                  <button
                                    className="btn btn--danger"
                                    type="button"
                                    onClick={() => handleDeleteSession(session.id)}
                                  >
                                    Delete event
                                  </button>
                                </div>
                              )}
                            </details>
                          )
                        })
                      ) : (
                        <div className="empty-state">No events match those filters.</div>
                      )}
                    </div>
                  </section>
                )}

                {(logFilter === "salvations" || logFilter === "both") && (
                  <section className="panel panel--full">
                    <h2>Salvations</h2>
                    <div className="salvation-results">
                      {logSalvations.length || logUnnamedSalvations.length ? (
                        <>
                          {logSalvations.map(({ person, session, dataset }) => {
                            if (!session) return null
                            const sessionLabel =
                              session.name ||
                              (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
                            const sessionDate = dateFormatter.format(new Date(`${session.date}T00:00:00`))
                            const isEditing = editingLogPersonId === person.id
                            const showTimeSpent =
                              session.type === "standalone" &&
                              person.time_spent_minutes !== null &&
                              person.time_spent_minutes !== undefined
                            return (
                              <div
                                key={person.id}
                                className={`salvation-row ${isEditing ? "is-editing" : ""}`}
                              >
                                <div className="salvation-row__header">
                                  <div>
                                    <strong>{person.name || "Unknown Name"}</strong>
                                    <div className="muted">
                                      {sessionLabel} - {sessionDate}
                                    </div>
                                  </div>
                                  <div className="salvation-row__meta">
                                    <span className="muted">{dataset ? dataset.name : "Personal"}</span>
                                    <span className="meta-pill meta-pill--compact">
                                      {formatSalvationRoleLabel(person.role)}
                                    </span>
                                    {!isEditing ? (
                                      <button
                                        className="btn btn--ghost btn--compact"
                                        type="button"
                                        onClick={() => handleEditLogPerson(person.id)}
                                      >
                                        Edit
                                      </button>
                                    ) : null}
                                  </div>
                                </div>
                                {isEditing ? (
                                  <form className="form-grid log-edit-form" onSubmit={handleSaveLogPersonEdit}>
                                    <label>
                                      Name (optional)
                                      <input
                                        type="text"
                                        value={logPersonName}
                                        onChange={(event) => setLogPersonName(event.target.value)}
                                        placeholder="Name of person saved (if known)"
                                      />
                                    </label>
                                    <label>
                                      Role
                                      <select
                                        value={logPersonRole}
                                        onChange={(event) =>
                                          setLogPersonRole(event.target.value as SalvationRole)
                                        }
                                      >
                                        <option value="presenter">Preacher</option>
                                        <option value="partner">Silent partner</option>
                                      </select>
                                    </label>
                                    {session.type === "standalone" ? (
                                      <label>
                                        Time spent (minutes, optional)
                                        <input
                                          type="number"
                                          min="0"
                                          value={logPersonTimeSpent}
                                          onChange={(event) => setLogPersonTimeSpent(event.target.value)}
                                          placeholder="e.g. 15"
                                        />
                                      </label>
                                    ) : null}
                                    <label>
                                      Tag search
                                      <input
                                        ref={logPersonTagInputRef}
                                        type="text"
                                        value={logPersonTagQuery}
                                        onChange={(event) => setLogPersonTagQuery(event.target.value)}
                                        placeholder="Search tags"
                                      />
                                    </label>
                                    <div className="form-span">
                                      <div className="muted">Tags</div>
                                      <div className="tag-grid">
                                        {logPersonTagOptions.length ? (
                                          logPersonTagOptions.map((tag) => (
                                            <label
                                              key={tag.id}
                                              className={`tag-pill ${
                                                logPersonTags.includes(tag.id) ? "is-active" : ""
                                              }`}
                                              style={
                                                { "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties
                                              }
                                            >
                                              <input
                                                type="checkbox"
                                                checked={logPersonTags.includes(tag.id)}
                                                onChange={() => toggleLogPersonTag(tag.id)}
                                              />
                                              {tag.name}
                                            </label>
                                          ))
                                        ) : (
                                          <span className="muted">
                                            {tags.length
                                              ? "No tags match that search."
                                              : "Create tags to assign them here."}
                                          </span>
                                        )}
                                        <button
                                          className="tag-pill tag-pill--add"
                                          type="button"
                                          onClick={handleQuickAddLogPersonTag}
                                        >
                                          {logPersonTagActionLabel}
                                        </button>
                                      </div>
                                      {tags.length > TOP_TAG_LIMIT && !logPersonTagQuery.trim() ? (
                                        <div className="note tag-hint">
                                          Showing top tags. Search to find more.
                                        </div>
                                      ) : null}
                                    </div>
                                    <label className="form-span">
                                      Notes (optional)
                                      <textarea
                                        className="textarea-compact"
                                        value={logPersonNotes}
                                        onChange={(event) => setLogPersonNotes(event.target.value)}
                                        placeholder="Optional details about this salvation"
                                      />
                                    </label>
                                    <div className="form-actions">
                                      <button className="btn btn--primary" type="submit">
                                        Save changes
                                      </button>
                                      <button
                                        className="btn btn--ghost"
                                        type="button"
                                        onClick={handleCancelLogPersonEdit}
                                      >
                                        Cancel edit
                                      </button>
                                    </div>
                                  </form>
                                ) : (
                                  <>
                                    <div className="tag-grid">
                                      {person.tagIds.length ? (
                                        person.tagIds.map((tagId) => {
                                          const tag = tagsById.get(tagId)
                                          if (!tag) return null
                                          return (
                                            <span
                                              key={tagId}
                                              className="tag-chip"
                                              style={{
                                                "--tag-color": tag.color ?? TAG_COLORS[0],
                                              } as React.CSSProperties}
                                            >
                                              {tag.name}
                                            </span>
                                          )
                                        })
                                      ) : (
                                        <span className="muted">No tags for this salvation.</span>
                                      )}
                                    </div>
                                    {person.notes ? <div className="note-text">{person.notes}</div> : null}
                                    {showTimeSpent ? (
                                      <div className="muted">
                                        Time spent: {formatDurationMinutes(person.time_spent_minutes)}
                                      </div>
                                    ) : null}
                                  </>
                                )}
                              </div>
                            )
                          })}
                          {logUnnamedSalvations.map(({ session, dataset, unnamedCount }) => {
                            const sessionLabel =
                              session.name ||
                              (session.type === "standalone"
                                ? "Standalone salvations"
                                : "Soulwinning event")
                            const sessionDate = dateFormatter.format(new Date(`${session.date}T00:00:00`))
                            const unnamedLabel = `${formatCount(
                              unnamedCount,
                              "unnamed salvation",
                              "unnamed salvations",
                            )} - ${sessionDate} - ${sessionLabel}`
                            return (
                              <div key={`unnamed-${session.id}`} className="salvation-row salvation-row--unnamed">
                                <div className="salvation-row__header">
                                  <div>
                                    <strong>{unnamedLabel}</strong>
                                  </div>
                                  <div className="salvation-row__meta">
                                    <span className="muted">{dataset ? dataset.name : "Personal"}</span>
                                    <button
                                      className="btn btn--ghost btn--compact"
                                      type="button"
                                      onClick={() => {
                                        setLogFilter("sessions")
                                        handleStartLogAddSalvation(session.id)
                                      }}
                                    >
                                      Add name
                                    </button>
                                  </div>
                                </div>
                              </div>
                            )
                          })}
                        </>
                      ) : (
                        <div className="empty-state">No salvations match those filters.</div>
                      )}
                    </div>
                  </section>
                )}
              </div>
        ) : view === "stats" ? (
          <>
          <div className="content-grid">
                <section className="panel panel--full">
                  <div className="panel-header">
                    <h2>Views</h2>
                    <div className="panel-header__actions">
                      <button
                        className="btn btn--ghost btn--compact"
                        type="button"
                        onClick={handleStartStatsViewCreate}
                      >
                        New view
                      </button>
                    </div>
                  </div>
                  <div className="note">
                    Pick a view and adjust filters below. Use search for quick stats lookups.
                  </div>
                  <div className="stats-controls">
                    <label>
                      Active view
                      <select
                        value={activeStatsViewId}
                        onChange={(event) => setActiveStatsViewId(event.target.value)}
                      >
                        {statsViews.map((viewItem) => (
                          <option key={viewItem.id} value={viewItem.id}>
                            {viewItem.name}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      Search stats
                      <input
                        type="text"
                        value={statsQuery}
                        onChange={(event) => setStatsQuery(event.target.value)}
                        placeholder="Search events, salvations, tags, or notes"
                      />
                    </label>
                  </div>
                  {statsQueryValue ? (
                    <div className="stats-search-results">
                      <div className="stats-search-summary">
                        Search results: {formatEventCount(statsSearchResults.events.length)} /{" "}
                        {formatSalvationCount(statsSearchResults.salvations.length)}
                      </div>
                      {statsSearchResults.events.length || statsSearchResults.salvations.length ? (
                        <div className="stats-search-grid">
                          <div className="stats-search-group">
                            <div className="stats-search-title">Events</div>
                            {statsSearchResults.eventPreview.length ? (
                              statsSearchResults.eventPreview.map(({ session, dataset, label }) => (
                                <div
                                  key={session.id}
                                  className="stats-search-item stats-search-item--button"
                                  role="button"
                                  tabIndex={0}
                                  onClick={() => handleOpenStatsEvent(session)}
                                  onKeyDown={(event) =>
                                    handleStatsSearchKeyDown(event, () =>
                                      handleOpenStatsEvent(session),
                                    )
                                  }
                                >
                                  <div className="stats-search-item__title">{label}</div>
                                  <div className="stats-search-item__meta">
                                    <span>{dateFormatter.format(new Date(`${session.date}T00:00:00`))}</span>
                                    <span>{dataset?.name ?? "Unknown data set"}</span>
                                    <span>{formatSalvationCount(session.saved_count ?? 0)}</span>
                                    <span className="stats-search-item__action">View details</span>
                                  </div>
                                </div>
                              ))
                            ) : (
                              <div className="muted">No events matched.</div>
                            )}
                            {statsSearchResults.eventOverflow ? (
                              <div className="muted">
                                + {numberFormatter.format(statsSearchResults.eventOverflow)} more events
                              </div>
                            ) : null}
                          </div>
                          <div className="stats-search-group">
                            <div className="stats-search-title">Salvations</div>
                            {statsSearchResults.salvationPreview.length ? (
                              statsSearchResults.salvationPreview.map(
                                ({ person, session, dataset, sessionLabel }) => (
                                  <div
                                    key={person.id}
                                    className="stats-search-item stats-search-item--button"
                                    role="button"
                                    tabIndex={0}
                                    onClick={() => handleOpenStatsSalvation(person, session)}
                                    onKeyDown={(event) =>
                                      handleStatsSearchKeyDown(event, () =>
                                        handleOpenStatsSalvation(person, session),
                                      )
                                    }
                                  >
                                    <div className="stats-search-item__title">
                                      {person.name || "Unknown Name"}
                                    </div>
                                    <div className="stats-search-item__meta">
                                      <span>
                                        {session
                                          ? dateFormatter.format(new Date(`${session.date}T00:00:00`))
                                          : "Date unknown"}
                                      </span>
                                      <span>{sessionLabel}</span>
                                      <span>{dataset?.name ?? "Unknown data set"}</span>
                                      <span className="meta-pill meta-pill--compact">
                                        {formatSalvationRoleLabel(person.role)}
                                      </span>
                                      <span className="stats-search-item__action">View details</span>
                                    </div>
                                  </div>
                                ),
                              )
                            ) : (
                              <div className="muted">No salvations matched.</div>
                            )}
                            {statsSearchResults.salvationOverflow ? (
                              <div className="muted">
                                + {numberFormatter.format(statsSearchResults.salvationOverflow)} more salvations
                              </div>
                            ) : null}
                          </div>
                        </div>
                      ) : (
                        <div className="muted">No matches yet.</div>
                      )}
                    </div>
                  ) : null}
                  <div className="stats-view-meta">
                    <div>
                      <div className="muted">Sections in this view</div>
                      <div className="tag-grid" style={{ marginTop: "8px" }}>
                        {(activeStatsView?.sections ?? []).length ? (
                          (activeStatsView?.sections ?? []).map((sectionId) => (
                            <span key={sectionId} className="tag-pill tag-pill--static">
                              {statsSectionLabelMap.get(sectionId) ?? sectionId}
                            </span>
                          ))
                        ) : (
                          <span className="muted">No sections selected.</span>
                        )}
                      </div>
                    </div>
                    <div className="stats-view-actions">
                      <button
                        className="btn btn--ghost btn--compact"
                        type="button"
                        onClick={() => handleStartStatsViewEdit(activeStatsViewId)}
                      >
                        Edit view
                      </button>
                    </div>
                  </div>
                  {statsBuilderOpen ? (
                    <div className="stats-view-builder">
                      <form className="form-grid two-col" onSubmit={handleSaveStatsView}>
                        <label className="form-span">
                          View name
                          <input
                            type="text"
                            value={statsViewName}
                            onChange={(event) => setStatsViewName(event.target.value)}
                            placeholder="Example: Quarterly focus"
                          />
                        </label>
                        <div className="form-span">
                          <div className="muted">
                            Sections selected: {numberFormatter.format(statsViewSections.length)}
                          </div>
                          <div className="stats-section-grid" style={{ marginTop: "8px" }}>
                            {STATS_SECTION_OPTIONS.map((section) => (
                              <label
                                key={section.id}
                                className={`stats-section-option ${
                                  statsViewSections.includes(section.id) ? "is-active" : ""
                                }`}
                              >
                                <input
                                  type="checkbox"
                                  checked={statsViewSections.includes(section.id)}
                                  onChange={() => toggleStatsViewSection(section.id)}
                                />
                                <div>
                                  <strong>{section.label}</strong>
                                  <div className="muted">{section.description}</div>
                                </div>
                              </label>
                            ))}
                          </div>
                        </div>
                        <div className="form-actions">
                          <button className="btn btn--primary" type="submit">
                            {editingStatsViewId ? "Save changes" : "Save view"}
                          </button>
                          <button
                            className="btn btn--ghost"
                            type="button"
                            onClick={handleCancelStatsViewEdit}
                          >
                            Cancel
                          </button>
                        </div>
                      </form>
                    </div>
                  ) : null}
                </section>

                <section className="panel panel--hero">
                  <h2>Filters</h2>
                  <div className="form-grid two-col">
                    <label>
                      Start date
                      <input
                        type="date"
                        value={rangeStart}
                        onChange={(event) => setRangeStart(event.target.value)}
                      />
                    </label>
                    <label>
                      End date
                      <input
                        type="date"
                        value={rangeEnd}
                        onChange={(event) => setRangeEnd(event.target.value)}
                      />
                    </label>
                    <label>
                      Group by
                      <select value={groupBy} onChange={(event) => setGroupBy(event.target.value as GroupBy)}>
                        <option value="year">Year</option>
                        <option value="month">Month</option>
                        <option value="week">Week</option>
                        <option value="day">Day</option>
                      </select>
                    </label>
                    <label>
                      Week starts on
                      <select
                        value={weekStartsOn}
                        onChange={(event) => setWeekStartsOn(Number(event.target.value) as 0 | 1)}
                      >
                        <option value={0}>Sunday</option>
                        <option value={1}>Monday</option>
                      </select>
                    </label>
                  </div>
                  <div className="quick-range">
                    <button className="btn btn--soft" type="button" onClick={() => setQuickRange(30)}>
                      Last 30 days
                    </button>
                    <button className="btn btn--soft" type="button" onClick={() => setQuickRange(90)}>
                      Last 90 days
                    </button>
                    <button className="btn btn--soft" type="button" onClick={() => setQuickRange(365)}>
                      Last year
                    </button>
                    <button className="btn btn--ghost" type="button" onClick={() => setQuickRange("all")}>
                      All time
                    </button>
                  </div>
                  <details className="filter-block" open={selectedTagIds.length > 0}>
                    <summary>Tag filters</summary>
                    <div className="tag-grid" style={{ marginTop: "8px" }}>
                      {tags.length ? (
                        tags.map((tag) => (
                          <button
                            key={tag.id}
                            type="button"
                            className={`tag-pill ${selectedTagIds.includes(tag.id) ? "is-active" : ""}`}
                            style={{ "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                            onClick={() => toggleTagFilter(tag.id)}
                          >
                            {tag.name}
                          </button>
                        ))
                      ) : (
                        <span className="muted">Create tags to filter stats.</span>
                      )}
                    </div>
                    {selectedTagIds.length ? (
                      <label style={{ marginTop: "12px" }}>
                        Tag match
                        <select
                          value={tagMatchMode}
                          onChange={(event) => setTagMatchMode(event.target.value as TagMatchMode)}
                        >
                          <option value="any">Any selected tags</option>
                          <option value="all">All selected tags</option>
                        </select>
                      </label>
                    ) : null}
                  </details>
                  <details className="filter-block" open={selectedDatasetIds.length > 0}>
                    <summary>Data set filters</summary>
                    <div className="tag-grid" style={{ marginTop: "8px" }}>
                      {datasets.length ? (
                        datasets.map((dataset) => (
                          <button
                            key={dataset.id}
                            type="button"
                            className={`tag-pill ${
                              selectedDatasetIds.includes(dataset.id) ? "is-active" : ""
                            }`}
                            onClick={() => toggleDatasetFilter(dataset.id)}
                          >
                            {dataset.name}
                          </button>
                        ))
                      ) : (
                        <span className="muted">Create data sets to filter stats.</span>
                      )}
                    </div>
                  </details>
                </section>

                {activeStatsSections.has("totals") ? (
                  <section className="panel">
                    <h2>Totals</h2>
                    <div className="stats-grid">
                      <div className="metric-card">
                        <h4>Events</h4>
                        <div className="metric-value">{numberFormatter.format(totals.totalSessions)}</div>
                      </div>
                      <div className="metric-card">
                        <h4>Salvations</h4>
                        <div className="metric-value">
                          {numberFormatter.format(totals.totalSavedReported)}
                        </div>
                      </div>
                      <div className="metric-card">
                        <h4>Personal salvations</h4>
                        <div className="metric-value">
                          {numberFormatter.format(salvationRoleCounts.presenter)}
                        </div>
                      </div>
                      <div className="metric-card">
                        <h4>Silent partner salvations</h4>
                        <div className="metric-value">
                          {numberFormatter.format(salvationRoleCounts.partner)}
                        </div>
                      </div>
                      <div className="metric-card">
                        <h4>Doors knocked</h4>
                        <div className="metric-value">{numberFormatter.format(totals.totalDoors)}</div>
                      </div>
                      <div className="metric-card">
                        <h4>Avg salvations per event</h4>
                        <div className="metric-value">{totals.avgSaved.toFixed(1)}</div>
                      </div>
                    </div>
                    {selectedTagIds.length ? (
                      <div className="note">
                        Tag filters match salvations or event tags, then include events tied to those matches.
                      </div>
                    ) : null}
                  </section>
                ) : null}

                {activeStatsSections.has("time_metrics") ? (
                  <section className="panel">
                    <h2>Time metrics</h2>
                    <div className="stats-grid">
                      <div className="metric-card">
                        <h4>Total time tracked</h4>
                        <div className="metric-value">
                          {timeMetrics.totalMinutes
                            ? formatDurationMinutes(timeMetrics.totalMinutes)
                            : "n/a"}
                        </div>
                      </div>
                      <div className="metric-card">
                        <h4>Avg event duration</h4>
                        <div className="metric-value">
                          {timeMetrics.eventsWithTime
                            ? formatDurationMinutes(timeMetrics.avgEventMinutes)
                            : "n/a"}
                        </div>
                      </div>
                      <div className="metric-card">
                        <h4>Salvations per hour</h4>
                        <div className="metric-value">
                          {timeMetrics.totalMinutes ? timeMetrics.salvationsPerHour.toFixed(1) : "n/a"}
                        </div>
                      </div>
                      <div className="metric-card">
                        <h4>Avg minutes per salvation</h4>
                        <div className="metric-value">
                          {timeMetrics.totalMinutes && totals.totalSavedReported
                            ? timeMetrics.minutesPerSalvation.toFixed(1)
                            : "n/a"}
                        </div>
                      </div>
                    </div>
                    <div className="note">
                      Events with time: {numberFormatter.format(timeMetrics.eventsWithTime)} /{" "}
                      {numberFormatter.format(totals.totalSessions)} | Standalone time entries:{" "}
                      {numberFormatter.format(timeMetrics.standaloneWithTime)}
                    </div>
                  </section>
                ) : null}

                {activeStatsSections.has("charts") ? (
                  <section className="panel">
                  <h2>Charts</h2>
                  <div className="chart-grid">
                    <div className="chart-card">
                      <h3>Salvations by period</h3>
                      <div className="bar-chart">
                        {groupRows.length ? (
                          groupRows.map((row) => (
                            <div key={row.key} className="bar-row">
                              <span className="bar-label">{row.label}</span>
                              <div className="bar-track">
                                <div
                                  className="bar-fill"
                                  style={{
                                    width: `${(row.savedReported / maxSavedInGroup) * 100}%`,
                                  }}
                                />
                              </div>
                              <span className="bar-value">
                                {numberFormatter.format(row.savedReported)}
                              </span>
                            </div>
                          ))
                        ) : (
                          <div className="empty-state">No data in this range.</div>
                        )}
                      </div>
                    </div>
                    <div className="chart-card">
                      <h3>Events by period</h3>
                      <div className="bar-chart">
                        {groupRows.length ? (
                          groupRows.map((row) => (
                            <div key={row.key} className="bar-row">
                              <span className="bar-label">{row.label}</span>
                              <div className="bar-track">
                                <div
                                  className="bar-fill bar-fill--soft"
                                  style={{
                                    width: `${(row.sessions / maxSessionsInGroup) * 100}%`,
                                  }}
                                />
                              </div>
                              <span className="bar-value">{numberFormatter.format(row.sessions)}</span>
                            </div>
                          ))
                        ) : (
                          <div className="empty-state">No data in this range.</div>
                        )}
                      </div>
                    </div>
                  </div>
                  <div className="chart-grid" style={{ marginTop: "16px" }}>
                    <div className="chart-card">
                      <h3>Ratios</h3>
                      <div className="ratio-grid">
                        <div>
                          <strong>{ratioMetrics.doorsPerSession.toFixed(1)}</strong>
                          <span className="muted">Doors per event</span>
                        </div>
                        <div>
                          <strong>{ratioMetrics.doorsPerSalvation.toFixed(1)}</strong>
                          <span className="muted">Doors per salvation</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>
                ) : null}

                {activeStatsSections.has("weekday") ? (
                  <section className="panel">
                  <h2>Weekday breakdown</h2>
                  <div className="chart-grid">
                    <div className="chart-card">
                      <h3>Salvations by weekday</h3>
                      <div className="bar-chart">
                        {sessionsMatchingTags.length ? (
                          weekdayRows.map((row) => (
                            <div key={row.label} className="bar-row">
                              <span className="bar-label">{row.label}</span>
                              <div className="bar-track">
                                <div
                                  className="bar-fill"
                                  style={{
                                    width: `${(row.salvations / maxWeekdaySalvations) * 100}%`,
                                  }}
                                />
                              </div>
                              <span className="bar-value">
                                {numberFormatter.format(row.salvations)}
                              </span>
                            </div>
                          ))
                        ) : (
                          <div className="empty-state">No data in this range.</div>
                        )}
                      </div>
                    </div>
                    <div className="chart-card">
                      <h3>Events by weekday</h3>
                      <div className="bar-chart">
                        {sessionsMatchingTags.length ? (
                          weekdayRows.map((row) => (
                            <div key={row.label} className="bar-row">
                              <span className="bar-label">{row.label}</span>
                              <div className="bar-track">
                                <div
                                  className="bar-fill bar-fill--soft"
                                  style={{
                                    width: `${(row.sessions / maxWeekdaySessions) * 100}%`,
                                  }}
                                />
                              </div>
                              <span className="bar-value">
                                {numberFormatter.format(row.sessions)}
                              </span>
                            </div>
                          ))
                        ) : (
                          <div className="empty-state">No data in this range.</div>
                        )}
                      </div>
                    </div>
                  </div>
                </section>
                ) : null}

                {activeStatsSections.has("insights") ? (
                  <section className="panel">
                  <h2>Insights</h2>
                  <div className="insight-grid">
                    <div className="insight-card">
                      <div className="insight-label">Most active weekday</div>
                      <div className="insight-value">
                        {mostActiveDay
                          ? `${mostActiveDay.label}`
                          : "n/a"}
                      </div>
                      <div className="muted">
                        {mostActiveDay
                          ? mostActiveDay.metric === "salvations"
                            ? formatSalvationCount(mostActiveDay.value)
                            : formatEventCount(mostActiveDay.value)
                          : "No events in this range"}
                      </div>
                    </div>
                    <div className="insight-card">
                      <div className="insight-label">Salvations breakdown</div>
                      <div className="insight-value">
                        {formatCount(
                          sessionTypeBreakdown.sessionCount,
                          "soulwinning event",
                          "soulwinning events",
                        )}{" "}
                        /{" "}
                        {formatCount(
                          sessionTypeBreakdown.standaloneCount,
                          "standalone entry",
                          "standalone entries",
                        )}
                      </div>
                      <div className="muted">
                        {totals.totalSessions
                          ? `${(
                              (sessionTypeBreakdown.standaloneCount / totals.totalSessions) *
                              100
                            ).toFixed(0)}% standalone`
                          : "No events in this range"}
                      </div>
                    </div>
                    <div className="insight-card">
                      <div className="insight-label">Salvation roles</div>
                      <div className="insight-value">
                        {formatCount(
                          salvationRoleCounts.presenter,
                          "preacher salvation",
                          "preacher salvations",
                        )}{" "}
                        /{" "}
                        {formatCount(
                          salvationRoleCounts.partner,
                          "silent partner salvation",
                          "silent partner salvations",
                        )}
                      </div>
                      <div className="muted">
                        {totals.namedSaved
                          ? `${(
                              (salvationRoleCounts.presenter / totals.namedSaved) *
                              100
                            ).toFixed(0)}% preacher`
                          : "No named salvations in this range"}
                      </div>
                    </div>
                  </div>
                </section>
                ) : null}

                {activeStatsSections.has("dataset_breakdown") ? (
                  <section className="panel">
                  <h2>Data set breakdown</h2>
                  <div className="tracker-breakdown">
                    {datasetBreakdown.length ? (
                      datasetBreakdown.map((entry) => (
                        <div key={entry.dataset.id} className="tracker-breakdown-card">
                          <div className="tracker-card__header">
                            <strong>{entry.dataset.name}</strong>
                          </div>
                          <div className="tracker-card__meta">
                            <span>{formatEventCount(entry.sessions)}</span>
                            <span>{formatSalvationCount(entry.savedReported)}</span>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="empty-state">No data set data in this range.</div>
                    )}
                  </div>
                </section>
                ) : null}

                {activeStatsSections.has("top_events") ? (
                  <section className="panel">
                  <h2>Top events</h2>
                  <div className="session-highlights">
                    {topSessions.length ? (
                      topSessions.map(({ session, dataset }) => (
                        <div key={session.id} className="session-highlight">
                          <div>
                            <strong>
                              {session.name ||
                                (session.type === "standalone"
                                  ? "Standalone salvations"
                                  : "Soulwinning event")}
                            </strong>
                            <div className="muted">
                              {dateFormatter.format(new Date(`${session.date}T00:00:00`))}
                              {dataset ? ` - ${dataset.name}` : ""}
                            </div>
                          </div>
                          <div className="session-highlight__stats">
                            <span>{formatSalvationCount(session.saved_count)}</span>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="empty-state">No events to highlight.</div>
                    )}
                  </div>
                </section>
                ) : null}

                {activeStatsSections.has("period_breakdown") ? (
                  <section className="panel">
                  <h2>Period breakdown</h2>
                  <table className="stats-table">
                    <thead>
                      <tr>
                        <th>Period</th>
                        <th>Events</th>
                        <th>Salvations</th>
                        <th>Doors</th>
                      </tr>
                    </thead>
                    <tbody>
                      {groupRows.length ? (
                        groupRows.map((row) => (
                          <tr key={row.key}>
                            <td>{row.label}</td>
                            <td>{numberFormatter.format(row.sessions)}</td>
                            <td>{numberFormatter.format(row.savedReported)}</td>
                            <td>{numberFormatter.format(row.doors)}</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={4} className="muted">
                            No data in this range.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </section>
                ) : null}

                {activeStatsSections.has("tag_breakdown") ? (
                  <section className="panel">
                  <h2>Tag breakdown</h2>
                  <div className="tag-breakdown">
                    {tagBreakdown.length ? (
                      tagBreakdown.map((item) => (
                        <div key={item.tag.id} className="tag-breakdown-item">
                          <span
                            className="tag-chip"
                            style={{ "--tag-color": item.tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                          >
                            {item.tag.name}
                          </span>
                          <div className="tag-breakdown-bar">
                            <div
                              className="tag-breakdown-fill"
                              style={{
                                width: `${(item.count / maxTagCount) * 100}%`,
                                backgroundColor: item.tag.color ?? TAG_COLORS[0],
                              }}
                            />
                          </div>
                          <strong>{numberFormatter.format(item.count)}</strong>
                        </div>
                      ))
                    ) : (
                      <div className="empty-state">No tag data for this range.</div>
                    )}
                  </div>
                </section>
                ) : null}

                {activeStatsSections.has("salvation_search") ? (
                  <section className="panel">
                  <h2>Salvation details</h2>
                  <div className="note">
                    Use the filters above and search by name for a more granular view of salvations.
                  </div>
                  <div className="form-grid">
                    <label>
                      Name or notes search
                      <input
                        type="text"
                        value={salvationQuery}
                        onChange={(event) => setSalvationQuery(event.target.value)}
                        placeholder="Search by name or notes"
                      />
                    </label>
                  </div>
                  <div className="salvation-results">
                    {salvationResults.length ? (
                      salvationResults.map(({ person, session, dataset }) => {
                        const sessionLabel = session
                          ? session.name ||
                            (session.type === "standalone" ? "Standalone salvations" : "Soulwinning event")
                          : "Event unknown"
                        const sessionDate = session
                          ? dateFormatter.format(new Date(`${session.date}T00:00:00`))
                          : "Date unknown"
                        return (
                          <div key={person.id} className="salvation-row">
                            <div className="salvation-row__header">
                              <div>
                                <strong>{person.name || "Unknown Name"}</strong>
                                <div className="muted">
                                  {sessionLabel} - {sessionDate}
                                </div>
                              </div>
                              <div className="salvation-row__meta">
                                <span className="muted">{dataset ? dataset.name : "Personal"}</span>
                                <span className="meta-pill meta-pill--compact">
                                  {formatSalvationRoleLabel(person.role)}
                                </span>
                              </div>
                            </div>
                            <div className="tag-grid">
                              {person.tagIds.length ? (
                                person.tagIds.map((tagId) => {
                                  const tag = tagsById.get(tagId)
                                  if (!tag) return null
                                  return (
                                    <span
                                      key={tagId}
                                      className="tag-chip"
                                      style={{
                                        "--tag-color": tag.color ?? TAG_COLORS[0],
                                      } as React.CSSProperties}
                                    >
                                      {tag.name}
                                    </span>
                                  )
                                })
                              ) : (
                                <span className="muted">No tags for this salvation.</span>
                              )}
                            </div>
                            {person.notes ? <div className="note-text">{person.notes}</div> : null}
                          </div>
                        )
                      })
                    ) : (
                      <div className="empty-state">No salvations match those filters.</div>
                    )}
                  </div>
                </section>
                ) : null}

              </div>
              {statsDetail ? (
                <>
                  <button
                    className="modal-backdrop"
                    type="button"
                    aria-label="Close"
                    onClick={() => setStatsDetail(null)}
                  />
                  <div className="modal-panel">
                    <section
                      className="panel modal-panel__content"
                      role="dialog"
                      aria-modal="true"
                      aria-label="Search result details"
                    >
                      <div className="modal-header">
                        <h2>Search result details</h2>
                        <button
                          className="btn btn--ghost btn--compact"
                          type="button"
                          onClick={() => setStatsDetail(null)}
                        >
                          Close
                        </button>
                      </div>
                      {statsDetail.kind === "event"
                        ? (() => {
                            const session = statsDetail.session
                            const dataset = datasetById.get(session.dataset_id)
                            const peopleInSession = peopleBySession.get(session.id) ?? []
                            const unnamedCount = Math.max(
                              0,
                              (session.saved_count ?? 0) - peopleInSession.length,
                            )
                            const sessionLabel =
                              session.name ||
                              (session.type === "standalone"
                                ? "Standalone salvations"
                                : "Soulwinning event")
                            const sessionDate = dateFormatter.format(
                              new Date(`${session.date}T00:00:00`),
                            )
                            const doorsCount =
                              typeof session.doors_knocked === "number" &&
                              Number.isFinite(session.doors_knocked)
                                ? session.doors_knocked
                                : null
                            const showDoors = doorsCount !== null && doorsCount > 0
                            const doorsLabel = showDoors
                              ? numberFormatter.format(doorsCount ?? 0)
                              : ""
                            const startLabel = formatTimeValue(session.start_time)
                            const endLabel = formatTimeValue(session.end_time)
                            const durationMinutes = getSessionDurationMinutes(
                              session.start_time,
                              session.end_time,
                            )
                            const durationLabel =
                              durationMinutes !== null
                                ? formatDurationMinutes(durationMinutes)
                                : ""
                            let timeLabel = ""
                            if (startLabel && endLabel) {
                              timeLabel = `${startLabel} - ${endLabel}`
                            } else if (startLabel) {
                              timeLabel = `Starts ${startLabel}`
                            } else if (endLabel) {
                              timeLabel = `Ends ${endLabel}`
                            }
                            const timeMeta = timeLabel
                              ? durationLabel
                                ? `${timeLabel} (${durationLabel})`
                                : timeLabel
                              : ""
                            return (
                              <div className="stats-detail">
                                <div className="stats-detail__header">
                                  <strong>{sessionLabel}</strong>
                                  <span className="muted">{sessionDate}</span>
                                </div>
                                <div className="session-summary__meta">
                                  <span>{formatSalvationCount(session.saved_count)}</span>
                                  {showDoors ? (
                                    <span className="meta-pill meta-pill--compact">
                                      Doors knocked: {doorsLabel}
                                    </span>
                                  ) : null}
                                  {timeMeta ? (
                                    <span className="meta-pill meta-pill--compact">Time: {timeMeta}</span>
                                  ) : null}
                                  <span>{dataset?.name ?? "Unknown data set"}</span>
                                </div>
                                {session.tagIds.length ? (
                                  <div>
                                    <div className="muted">Event tags</div>
                                    <div className="tag-grid" style={{ marginTop: "6px" }}>
                                      {session.tagIds.map((tagId) => {
                                        const tag = tagsById.get(tagId)
                                        if (!tag) return null
                                        return (
                                          <span
                                            key={tagId}
                                            className="tag-chip"
                                            style={{
                                              "--tag-color": tag.color ?? TAG_COLORS[0],
                                            } as React.CSSProperties}
                                          >
                                            {tag.name}
                                          </span>
                                        )
                                      })}
                                    </div>
                                  </div>
                                ) : null}
                                {session.notes ? (
                                  <div className="session-notes">
                                    <strong>Event notes</strong>
                                    <div className="note-text">{session.notes}</div>
                                  </div>
                                ) : null}
                                <div className="session-people">
                                  {peopleInSession.length ? (
                                    peopleInSession.map((person) => (
                                      <div key={person.id} className="session-person">
                                        <strong>{person.name || "Unknown Name"}</strong>
                                        <span className="meta-pill meta-pill--compact">
                                          {formatSalvationRoleLabel(person.role)}
                                        </span>
                                        {person.tagIds.map((tagId) => {
                                          const tag = tagsById.get(tagId)
                                          if (!tag) return null
                                          return (
                                            <span
                                              key={tagId}
                                              className="tag-chip"
                                              style={{
                                                "--tag-color": tag.color ?? TAG_COLORS[0],
                                              } as React.CSSProperties}
                                            >
                                              {tag.name}
                                            </span>
                                          )
                                        })}
                                        {person.notes ? (
                                          <div className="person-note">{person.notes}</div>
                                        ) : null}
                                      </div>
                                    ))
                                  ) : (
                                    <div className="muted">No salvations listed for this event.</div>
                                  )}
                                </div>
                                {unnamedCount > 0 ? (
                                  <div className="note-text">
                                    {formatCount(
                                      unnamedCount,
                                      "unnamed salvation",
                                      "unnamed salvations",
                                    )}{" "}
                                    remaining.
                                  </div>
                                ) : null}
                              </div>
                            )
                          })()
                        : (() => {
                            const person = statsDetail.person
                            const session = statsDetail.session
                            const dataset = session ? datasetById.get(session.dataset_id) : undefined
                            const sessionLabel = session
                              ? session.name ||
                                (session.type === "standalone"
                                  ? "Standalone salvations"
                                  : "Soulwinning event")
                              : "Event unknown"
                            const sessionDate = session
                              ? dateFormatter.format(new Date(`${session.date}T00:00:00`))
                              : "Date unknown"
                            const showTimeSpent =
                              session?.type === "standalone" &&
                              person.time_spent_minutes !== null &&
                              person.time_spent_minutes !== undefined
                            return (
                              <div className="stats-detail">
                                <div className="stats-detail__header">
                                  <strong>{person.name || "Unknown Name"}</strong>
                                  <span className="muted">
                                    {sessionLabel} - {sessionDate}
                                  </span>
                                </div>
                                <div className="session-summary__meta">
                                  <span>{dataset?.name ?? "Unknown data set"}</span>
                                  <span className="meta-pill meta-pill--compact">
                                    {formatSalvationRoleLabel(person.role)}
                                  </span>
                                  {showTimeSpent ? (
                                    <span className="meta-pill meta-pill--compact">
                                      Time spent: {formatDurationMinutes(person.time_spent_minutes)}
                                    </span>
                                  ) : null}
                                </div>
                                <div className="tag-grid">
                                  {person.tagIds.length ? (
                                    person.tagIds.map((tagId) => {
                                      const tag = tagsById.get(tagId)
                                      if (!tag) return null
                                      return (
                                        <span
                                          key={tagId}
                                          className="tag-chip"
                                          style={{
                                            "--tag-color": tag.color ?? TAG_COLORS[0],
                                          } as React.CSSProperties}
                                        >
                                          {tag.name}
                                        </span>
                                      )
                                    })
                                  ) : (
                                    <span className="muted">No tags for this salvation.</span>
                                  )}
                                </div>
                                {session?.tagIds?.length ? (
                                  <div>
                                    <div className="muted">Event tags</div>
                                    <div className="tag-grid" style={{ marginTop: "6px" }}>
                                      {session.tagIds.map((tagId) => {
                                        const tag = tagsById.get(tagId)
                                        if (!tag) return null
                                        return (
                                          <span
                                            key={tagId}
                                            className="tag-chip"
                                            style={{
                                              "--tag-color": tag.color ?? TAG_COLORS[0],
                                            } as React.CSSProperties}
                                          >
                                            {tag.name}
                                          </span>
                                        )
                                      })}
                                    </div>
                                  </div>
                                ) : null}
                                {person.notes ? <div className="note-text">{person.notes}</div> : null}
                              </div>
                            )
                          })()}
                      <div className="form-actions" style={{ marginTop: "16px" }}>
                        <button className="btn btn--primary" type="button" onClick={handleViewStatsDetailInLog}>
                          View in Log
                        </button>
                        <button
                          className="btn btn--ghost"
                          type="button"
                          onClick={() => setStatsDetail(null)}
                        >
                          Close
                        </button>
                      </div>
                    </section>
                  </div>
                </>
              ) : null}
          </>
        ) : null}
        {settingsOpen ? (
          <>
            <button
              className="settings-backdrop"
              type="button"
              aria-label="Close settings"
              onClick={() => setSettingsOpen(false)}
            />
            <aside className="settings-drawer" role="dialog" aria-modal="true" aria-label="Settings">
              <div className="settings-drawer__header">
                <h2>Settings</h2>
                <button
                  className="btn btn--ghost btn--compact"
                  type="button"
                  onClick={() => setSettingsOpen(false)}
                >
                  Close
                </button>
              </div>
              <div className="settings-drawer__body">
                <div className="settings-quick-toggle">
                  <div className="themeToggle">
                    <span className="themeToggle__label">Dark mode</span>
                    <label className="themeToggle__switch">
                      <input
                        type="checkbox"
                        checked={themeMode === "dark"}
                        onChange={() =>
                          setThemeMode((prev) => (prev === "dark" ? "light" : "dark"))
                        }
                        aria-label="Toggle dark mode"
                      />
                      <span className="themeToggle__slider" />
                    </label>
                  </div>
                </div>
                {settingsSection !== "overview" ? (
                  <button
                    className="btn btn--ghost btn--compact"
                    type="button"
                    onClick={() => setSettingsSection("overview")}
                  >
                    Back to settings
                  </button>
                ) : null}

                {settingsSection === "overview" ? (
                  <div className="settings-menu">
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("appearance")}
                    >
                      <span className="settings-menu-title">Appearance</span>
                      <span className="settings-menu-text">Full-app theme palettes.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("tags")}
                    >
                      <span className="settings-menu-title">Tags</span>
                      <span className="settings-menu-text">Create and manage your tag library.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("datasets")}
                    >
                      <span className="settings-menu-title">Data sets</span>
                      <span className="settings-menu-text">Create and manage data sets.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("exports")}
                    >
                      <span className="settings-menu-title">CSV export</span>
                      <span className="settings-menu-text">Export data in CSV format.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("backup")}
                    >
                      <span className="settings-menu-title">Backup &amp; transfer</span>
                      <span className="settings-menu-text">Move data between devices.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("sync")}
                    >
                      <span className="settings-menu-title">Cloud sync</span>
                      <span className="settings-menu-text">Save and restore data across devices.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("howto")}
                    >
                      <span className="settings-menu-title">How To</span>
                      <span className="settings-menu-text">How to use various app features.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={handleOpenSupport}
                    >
                      <span className="settings-menu-title">Support</span>
                      <span className="settings-menu-text">Send support via Ko-fi.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("feedback")}
                    >
                      <span className="settings-menu-title">Submit Feedback</span>
                      <span className="settings-menu-text">Send feedback or report an issue.</span>
                    </button>
                    <button
                      className="settings-menu-item"
                      type="button"
                      onClick={() => setSettingsSection("update")}
                    >
                      <span className="settings-menu-title">App updates</span>
                      <span className="settings-menu-text">Check for the latest desktop release.</span>
                    </button>
                    <button
                      className="settings-menu-item settings-menu-item--danger"
                      type="button"
                      onClick={() => setSettingsSection("danger")}
                    >
                      <span className="settings-menu-title">Delete All Data</span>
                      <span className="settings-menu-text">Clear all locally stored data.</span>
                    </button>
                  </div>
                ) : null}

                {settingsSection === "summary" ? (
                  <div className="settings-section">
                    <section className="panel panel--soft">
                      <h3>Data summary</h3>
                      <div className="note">
                        Select a tile to view details in the highlighted panel below.
                      </div>
                      <div className="stats-grid">
                        <button
                          className={`metric-card metric-card--button ${
                            settingsDetail === "sessions" ? "is-active" : ""
                          }`}
                          type="button"
                          onClick={() => toggleSettingsDetail("sessions")}
                        >
                          <h4>Events</h4>
                          <div className="metric-value">{numberFormatter.format(sessions.length)}</div>
                        </button>
                        <button
                          className={`metric-card metric-card--button ${
                            settingsDetail === "salvations" ? "is-active" : ""
                          }`}
                          type="button"
                          onClick={() => toggleSettingsDetail("salvations")}
                        >
                          <h4>Salvations</h4>
                          <div className="metric-value">
                            {numberFormatter.format(salvationLogEntries.length)}
                          </div>
                        </button>
                        <button
                          className={`metric-card metric-card--button ${
                            settingsDetail === "tags" ? "is-active" : ""
                          }`}
                          type="button"
                          onClick={() => toggleSettingsDetail("tags")}
                        >
                          <h4>Tags</h4>
                          <div className="metric-value">{numberFormatter.format(tags.length)}</div>
                        </button>
                        <button
                          className={`metric-card metric-card--button ${
                            settingsDetail === "datasets" ? "is-active" : ""
                          }`}
                          type="button"
                          onClick={() => toggleSettingsDetail("datasets")}
                        >
                          <h4>Data sets</h4>
                          <div className="metric-value">{numberFormatter.format(datasets.length)}</div>
                        </button>
                      </div>
                    </section>

                    {settingsDetail ? (
                      <section className="panel panel--accent" ref={settingsDetailRef}>
                        <div className="panel-header">
                          <h3>Data summary: {settingsDetailLabel}</h3>
                          <div className="panel-header__actions">
                            <button
                              className="btn btn--ghost btn--compact"
                              type="button"
                              onClick={() => setSettingsDetail(null)}
                            >
                              Close
                            </button>
                          </div>
                        </div>
                        <div className="note">
                          Showing {settingsDetailLabel.toLowerCase()} referenced by the data summary tiles.
                        </div>
                        {settingsDetail === "sessions" ? (
                          <div className="settings-list">
                            {sessionsSorted.length ? (
                              sessionsSorted.map((session) => {
                                const dataset = datasetById.get(session.dataset_id)
                                const sessionLabel =
                                  session.name ||
                                  (session.type === "standalone"
                                    ? "Standalone salvations"
                                    : "Soulwinning event")
                                return (
                                  <div key={session.id} className="settings-list-item">
                                    <strong>{sessionLabel}</strong>
                                    <div className="settings-list-item__meta">
                                      {dateFormatter.format(new Date(`${session.date}T00:00:00`))} -{" "}
                                      {formatSalvationCount(session.saved_count)}
                                      {dataset ? ` - ${dataset.name}` : ""}
                                    </div>
                                  </div>
                                )
                              })
                            ) : (
                              <div className="empty-state">No events yet.</div>
                            )}
                          </div>
                        ) : null}

                        {settingsDetail === "salvations" ? (
                          <div className="settings-list">
                            {salvationLogEntries.length ? (
                              salvationLogEntries.map(({ person, session, dataset }) => {
                                if (!session) return null
                                const sessionLabel =
                                  session.name ||
                                  (session.type === "standalone"
                                    ? "Standalone salvations"
                                    : "Soulwinning event")
                                const sessionDate = dateFormatter.format(
                                  new Date(`${session.date}T00:00:00`),
                                )
                                return (
                                  <div key={person.id} className="settings-list-item">
                                    <strong>{person.name || "Unknown Name"}</strong>
                                    <div className="settings-list-item__meta">
                                      {sessionLabel} - {sessionDate}
                                      {dataset ? ` - ${dataset.name}` : ""}
                                    </div>
                                  </div>
                                )
                              })
                            ) : (
                              <div className="empty-state">No salvations yet.</div>
                            )}
                          </div>
                        ) : null}

                        {settingsDetail === "tags" ? (
                          <div className="settings-list">
                            {tagsSortedByUsage.length ? (
                              tagsSortedByUsage.map((tag) => {
                                const usage = buildTagUsagePreview(tag.id)
                                return (
                                  <div key={tag.id} className="settings-list-item">
                                    <strong>{tag.name}</strong>
                                    <div className="settings-list-item__meta">
                                      {usage.hasEntries ? (
                                        <>
                                          Recent: {usage.label}
                                          {usage.remaining ? ` +${usage.remaining} more` : ""}
                                        </>
                                      ) : (
                                        usage.label
                                      )}
                                    </div>
                                  </div>
                                )
                              })
                            ) : (
                              <div className="empty-state">No tags yet.</div>
                            )}
                          </div>
                        ) : null}

                        {settingsDetail === "datasets" ? (
                          <div className="settings-list">
                            {datasets.length ? (
                              datasets.map((dataset) => {
                                const counts = datasetCounts.get(dataset.id)
                                return (
                                  <div key={dataset.id} className="settings-list-item">
                                    <strong>{dataset.name}</strong>
                                    <div className="settings-list-item__meta">
                                      {formatEventCount(counts?.sessions ?? 0)} -{" "}
                                      {formatSalvationCount(counts?.salvations ?? 0)}
                                    </div>
                                  </div>
                                )
                              })
                            ) : (
                              <div className="empty-state">No data sets yet.</div>
                            )}
                          </div>
                        ) : null}
                      </section>
                    ) : null}
                  </div>
                ) : null}

                {settingsSection === "appearance" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>Appearance</h3>
                    <div className="note">Choose a theme palette for the app.</div>
                    <div className="themePickerHint">
                      Each theme includes light and dark palettes. Current theme:{" "}
                      {AVAILABLE_THEMES.find((item) => item.id === themeName)?.name ?? "Default"}
                    </div>
                    <div className="themeGrid">
                      {AVAILABLE_THEMES.map((themeDef) => {
                        const isSelected = themeDef.id === themeName
                        const lightBg = themeDef.light["--panel-bg"] ?? themeDef.light["--bg"]
                        const lightBorder =
                          themeDef.light["--border-strong"] ?? themeDef.light["--border"]
                        const lightAccent =
                          themeDef.light["--accent-2"] ??
                          themeDef.light["--accent"] ??
                          themeDef.light["--chip-bg"] ??
                          lightBorder
                        const lightChip =
                          themeDef.light["--chip-bg"] ??
                          themeDef.light["--accent-2"] ??
                          themeDef.light["--accent"] ??
                          lightBorder
                        const darkBg = themeDef.dark["--panel-bg"] ?? themeDef.dark["--bg"]
                        const darkBorder =
                          themeDef.dark["--border-strong"] ?? themeDef.dark["--border"]
                        const darkAccent =
                          themeDef.dark["--accent-2"] ??
                          themeDef.dark["--accent"] ??
                          themeDef.dark["--chip-bg"] ??
                          darkBorder
                        const darkChip =
                          themeDef.dark["--chip-bg"] ??
                          themeDef.dark["--accent-2"] ??
                          themeDef.dark["--accent"] ??
                          darkBorder

                        return (
                          <button
                            key={themeDef.id}
                            type="button"
                            className={`themeCard ${isSelected ? "themeCard--active" : ""}`}
                            onClick={() => setThemeName(themeDef.id)}
                          >
                            <div className="themeCard__row">
                              <div className="themeCard__name">{themeDef.name}</div>
                              {isSelected ? <div className="themeCard__badge">Active</div> : null}
                            </div>
                            <div className="themePreview">
                              <div
                                className="themePreview__tile"
                                style={{
                                  background: lightBg,
                                  borderColor: lightBorder,
                                  color: themeDef.light["--text"],
                                }}
                              >
                                <div className="themePreview__bar" style={{ background: lightAccent }} />
                                <div className="themePreview__chip" style={{ background: lightChip }} />
                              </div>
                              <div
                                className="themePreview__tile"
                                style={{
                                  background: darkBg,
                                  borderColor: darkBorder,
                                  color: themeDef.dark["--text"],
                                }}
                              >
                                <div className="themePreview__bar" style={{ background: darkAccent }} />
                                <div className="themePreview__chip" style={{ background: darkChip }} />
                              </div>
                            </div>
                            <div className="themeCard__desc">{themeDef.description}</div>
                          </button>
                        )
                      })}
                    </div>
                  </section>
                ) : null}

                {settingsSection === "tags" ? (
                  <section className="panel panel--soft settings-section">
                    <div className="panel-header">
                      <h3>Tags</h3>
                      <button
                        className="info-button"
                        type="button"
                        aria-label="Tag info"
                        data-tooltip={`You can define tags here for extra granularity on your statistics. Example: you can create tags like "Apartment" or "Residential", and assign them to individual salvations tracked so that you can get a better breakdown of statistics in the More Stats tab.`}
                      >
                        ?
                      </button>
                    </div>
                    <form className="form-grid" onSubmit={handleCreateTag}>
                      <label>
                        Tag name
                        <input
                          type="text"
                          value={tagName}
                          onChange={(event) => setTagName(event.target.value)}
                          placeholder="Example: Apartment, Residential, Park, or other descriptors"
                        />
                      </label>
                      <div>
                        <div className="muted">Color</div>
                        <div className="tag-picker">
                          {TAG_COLORS.map((color) => (
                            <button
                              key={color}
                              type="button"
                              className={`color-dot ${tagColor === color ? "is-active" : ""}`}
                              style={{ backgroundColor: color }}
                              onClick={() => setTagColor(color)}
                              aria-label={`Tag color ${color}`}
                            />
                          ))}
                        </div>
                      </div>
                      <div className="form-actions">
                        <button className="btn btn--primary" type="submit">
                          Add tag
                        </button>
                      </div>
                    </form>
                    <div className="tag-card-list" style={{ marginTop: "16px" }}>
                      {tags.length ? (
                        tags.map((tag) => {
                          const usage = buildTagUsagePreview(tag.id)
                          return (
                            <div key={tag.id} className="tag-card">
                              {editingTagId === tag.id ? (
                                <form
                                  className="tag-card__edit"
                                  onSubmit={(event) => handleSaveTagEdit(event, tag.id)}
                                >
                                  <label>
                                    Tag name
                                    <input
                                      type="text"
                                      value={editTagName}
                                      onChange={(event) => setEditTagName(event.target.value)}
                                    />
                                  </label>
                                  <div>
                                    <div className="muted">Color</div>
                                    <div className="tag-picker">
                                      {TAG_COLORS.map((color) => (
                                        <button
                                          key={color}
                                          type="button"
                                          className={`color-dot ${editTagColor === color ? "is-active" : ""}`}
                                          style={{ backgroundColor: color }}
                                          onClick={() => setEditTagColor(color)}
                                          aria-label={`Tag color ${color}`}
                                        />
                                      ))}
                                    </div>
                                  </div>
                                  <div className="form-actions">
                                    <button className="btn btn--primary btn--compact" type="submit">
                                      Save
                                    </button>
                                    <button
                                      className="btn btn--ghost btn--compact"
                                      type="button"
                                      onClick={handleCancelTagEdit}
                                    >
                                      Cancel
                                    </button>
                                  </div>
                                </form>
                              ) : (
                                <div className="tag-card__header">
                                  <span
                                    className="tag-chip"
                                    style={{ "--tag-color": tag.color ?? TAG_COLORS[0] } as React.CSSProperties}
                                  >
                                    {tag.name}
                                  </span>
                                  <div className="tag-card__actions">
                                    <button
                                      className="btn btn--ghost btn--compact"
                                      type="button"
                                      onClick={() => handleStartTagEdit(tag)}
                                    >
                                      Edit
                                    </button>
                                    <button
                                      className="btn btn--danger btn--compact"
                                      type="button"
                                      onClick={() => handleDeleteTag(tag.id)}
                                    >
                                      Delete
                                    </button>
                                  </div>
                                </div>
                              )}
                              <div className="tag-card__details">
                                {usage.hasEntries ? (
                                  <span className="tag-card__meta">
                                    Recent: {usage.label}
                                    {usage.remaining ? ` +${usage.remaining} more` : ""}
                                  </span>
                                ) : (
                                  <span className="muted">{usage.label}</span>
                                )}
                              </div>
                            </div>
                          )
                        })
                      ) : (
                        <div className="empty-state">No tags yet.</div>
                      )}
                    </div>
                  </section>
                ) : null}

                {settingsSection === "datasets" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>Data sets</h3>
                    <div className="note">
                      Create data sets for personal, family, or church totals and filter them in More Stats.
                    </div>
                    <form className="form-grid" onSubmit={handleCreateDataset}>
                      <label>
                        Data set name
                        <input
                          type="text"
                          value={datasetName}
                          onChange={(event) => setDatasetName(event.target.value)}
                          placeholder="Example: Church Stats, Personal Stats, etc."
                        />
                      </label>
                      <div className="form-actions">
                        <button className="btn btn--primary" type="submit">
                          Add data set
                        </button>
                      </div>
                    </form>
                    <div className="tracker-grid">
                      {datasets.map((dataset) => {
                        const counts = datasetCounts.get(dataset.id)
                        return (
                          <div key={dataset.id} className="tracker-card">
                            {editingDatasetId === dataset.id ? (
                              <form
                                className="tracker-card__edit"
                                onSubmit={(event) => handleRenameDataset(event, dataset.id)}
                              >
                                <label>
                                  Data set name
                                  <input
                                    type="text"
                                    value={datasetRename}
                                    onChange={(event) => setDatasetRename(event.target.value)}
                                  />
                                </label>
                                <div className="form-actions">
                                  <button className="btn btn--primary btn--compact" type="submit">
                                    Save
                                  </button>
                                  <button
                                    className="btn btn--ghost btn--compact"
                                    type="button"
                                    onClick={handleCancelDatasetRename}
                                  >
                                    Cancel
                                  </button>
                                </div>
                              </form>
                            ) : (
                              <div className="tracker-card__header">
                                <strong>{dataset.name}</strong>
                                <div className="tracker-card__actions">
                                  <button
                                    className="btn btn--ghost btn--compact"
                                    type="button"
                                    onClick={() => handleStartDatasetRename(dataset)}
                                  >
                                    Rename
                                  </button>
                                  {datasets.length > 1 ? (
                                    <button
                                      className="btn btn--danger btn--compact"
                                      type="button"
                                      onClick={() => handleDeleteDataset(dataset.id)}
                                    >
                                      Delete
                                    </button>
                                  ) : null}
                                </div>
                              </div>
                            )}
                            <div className="tracker-card__meta">
                              <span>{formatEventCount(counts?.sessions ?? 0)}</span>
                              <span>{formatSalvationCount(counts?.salvations ?? 0)}</span>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </section>
                ) : null}

                {settingsSection === "howto" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>How To</h3>
                    <div className="howto-list">
                      <div className="howto-item">
                        <strong>Log a soulwinning time</strong>
                        <div className="howto-item__body">
                          Go to Home. Fill in the date, data set, salvations, and doors (optional). Use "Add
                          Salvations to This Event" to add names and tags, then click "Save event" at the
                          bottom.
                        </div>
                      </div>
                      <div className="howto-item">
                        <strong>Add a standalone salvation</strong>
                        <div className="howto-item__body">
                          In Home, use "Add Other Salvations" for one-off salvations like a grocery store or gas
                          station. Choose date and data set, add name/tags, then click "Save salvation".
                        </div>
                      </div>
                      <div className="howto-item">
                        <strong>Use tags</strong>
                      <div className="howto-item__body">
                          Create tags from the tag search while logging, or manage them in Settings (top right). Assign tags
                          while adding salvations, then use tag filters in More Stats to see breakdowns.
                      </div>
                    </div>
                    <div className="howto-item">
                      <strong>Use data sets</strong>
                      <div className="howto-item__body">
                          Create data sets in Settings (top right) for personal, family, or church tracking. Choose a data
                          set when you log an event or standalone salvation. Filter by data set in More Stats.
                      </div>
                    </div>
                      <div className="howto-item">
                        <strong>Review your history</strong>
                        <div className="howto-item__body">
                          Use the Log tab to view all events, all salvations, or both. Expand an event to see its
                          listed salvations.
                        </div>
                      </div>
                      <div className="howto-item">
                        <strong>Explore More Stats</strong>
                      <div className="howto-item__body">
                          Set date ranges and group-by options. Use tag and data set filters to focus on specific
                          areas. Review charts, breakdowns, and details for insights.
                      </div>
                      </div>
                      <div className="howto-item">
                        <strong>Set goals</strong>
                        <div className="howto-item__body">
                          Use the Goals tab to create custom goals with metrics, timeframes, and filters. Track
                          progress on Home and in the Goals dashboard.
                        </div>
                      </div>
                      <div className="howto-item">
                        <strong>Backups and exports</strong>
                      <div className="howto-item__body">
                          In Settings (top right), export CSV for spreadsheets, or export JSON for full backups. Import JSON
                          to restore or transfer to another device.
                      </div>
                    </div>
                    </div>
                  </section>
                ) : null}

                {settingsSection === "exports" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>CSV export</h3>
                    <div className="note">Exports all saved data (not just filtered rows).</div>
                    <div className="form-actions">
                      <button className="btn btn--soft" type="button" onClick={exportSessionsCsv}>
                        Export events CSV
                      </button>
                      <button className="btn btn--soft" type="button" onClick={exportPeopleCsv}>
                        Export salvations CSV
                      </button>
                    </div>
                  </section>
                ) : null}

                {settingsSection === "backup" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>Backup &amp; transfer</h3>
                    <div className="note">
                      Your data is stored locally on this device. Export JSON to move it to another device, and
                      import JSON to restore or sync.
                    </div>
                    <div className="note">
                      Tip: test a backup after exporting by importing it on another device.
                    </div>
                    <div className="form-actions">
                      <button className="btn btn--soft" type="button" onClick={exportJson}>
                        Export JSON backup
                      </button>
                      <button className="btn btn--soft" type="button" onClick={handleImportClick}>
                        Import JSON backup
                      </button>
                      <button className="btn btn--ghost" type="button" onClick={handleBackupCheckClick}>
                        Verify JSON backup
                      </button>
                    </div>
                    {backupCheckResult ? (
                      <>
                        <div
                          className={`status status--inline ${
                            backupCheckResult.ok ? "status--success" : "status--error"
                          }`}
                        >
                          {backupCheckResult.message}
                        </div>
                        {backupCheckResult.summary ? (
                          <div className="note-text">{backupCheckResult.summary}</div>
                        ) : null}
                      </>
                    ) : null}
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept="application/json"
                      onChange={handleImportJson}
                      style={{ display: "none" }}
                    />
                    <input
                      ref={backupCheckInputRef}
                      type="file"
                      accept="application/json"
                      onChange={handleBackupCheckFile}
                      style={{ display: "none" }}
                    />
                  </section>
                ) : null}

                {settingsSection === "sync" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>Cloud sync</h3>
                    <div className="note">
                      Sign in with email/password and your data will auto-save to cloud while you use the app.
                    </div>
                    {!isSupabaseConfigured ? (
                      <div className="empty-state">
                        Supabase is not configured. Add `VITE_SUPABASE_URL` and
                        `VITE_SUPABASE_ANON_KEY` to this deployment&apos;s environment variables.
                      </div>
                    ) : !authUserId ? (
                      <form className="form-grid auth-card" onSubmit={handleAuthSubmit}>
                        <div className="auth-toggle">
                          <button
                            className={`btn btn--compact ${authMode === "signin" ? "btn--primary" : "btn--ghost"}`}
                            type="button"
                            onClick={() => setAuthMode("signin")}
                            disabled={authLoading}
                          >
                            Sign in
                          </button>
                          <button
                            className={`btn btn--compact ${authMode === "signup" ? "btn--primary" : "btn--ghost"}`}
                            type="button"
                            onClick={() => setAuthMode("signup")}
                            disabled={authLoading}
                          >
                            Create account
                          </button>
                        </div>
                        <label>
                          Email
                          <input
                            type="email"
                            value={authEmail}
                            onChange={(event) => setAuthEmail(event.target.value)}
                            placeholder="you@example.com"
                            autoComplete="email"
                            required
                          />
                        </label>
                        <label>
                          Password
                          <input
                            type="password"
                            value={authPassword}
                            onChange={(event) => setAuthPassword(event.target.value)}
                            placeholder="Enter password"
                            autoComplete={authMode === "signin" ? "current-password" : "new-password"}
                            required
                          />
                        </label>
                        <div className="note-text">
                          {authMode === "signup"
                            ? "Account sign-up may require email confirmation depending on your Supabase auth settings."
                            : "Sign in to upload or restore your cloud backup."}
                        </div>
                        <div className="form-actions">
                          <button
                            className="btn btn--primary"
                            type="submit"
                            disabled={authLoading || syncLoading || loading}
                          >
                            {authLoading
                              ? "Working..."
                              : authMode === "signin"
                                ? "Sign in"
                                : "Create account"}
                          </button>
                        </div>
                      </form>
                    ) : (
                      <div className="form-grid">
                        <div className="note-text">
                          Signed in as <strong>{authUserEmail || "Unknown user"}</strong>.
                        </div>
                        <div className="note-text">
                          Auto-save: <strong>{autoSyncing ? "Saving..." : "On"}</strong>
                        </div>
                        <div className="note-text">
                          Last upload: {syncLastUploadedLabel || "Not uploaded yet."}
                        </div>
                        <div className="note-text">
                          Last download: {syncLastDownloadedLabel || "Not downloaded yet."}
                        </div>
                        <div className="note-text">
                          Rollback snapshot: {rollbackSavedLabel || "Not available yet."}
                        </div>
                        {autoSyncError ? (
                          <div className="status status--error status--inline">
                            Auto-save error: {autoSyncError}
                          </div>
                        ) : null}
                        <div className="note">
                          Auto-save writes local changes to cloud. Use Download to pull your latest cloud
                          backup onto this device. Conflict guard prevents auto-overwrite of unsynced local
                          data.
                        </div>
                        <div className="form-actions">
                          <button
                            className="btn btn--primary"
                            type="button"
                            onClick={() => {
                              void handleCloudSyncUpload()
                            }}
                            disabled={syncLoading || authLoading || loading || autoSyncing}
                          >
                            {syncLoading ? "Working..." : "Sync now"}
                          </button>
                          <button
                            className="btn btn--soft"
                            type="button"
                            onClick={() => {
                              void handleCloudSyncDownload()
                            }}
                            disabled={syncLoading || authLoading || loading || autoSyncing}
                          >
                            Download from cloud
                          </button>
                          <button
                            className="btn btn--ghost"
                            type="button"
                            onClick={handleRestoreRollbackSnapshot}
                            disabled={syncLoading || authLoading || loading || autoSyncing || !rollbackSavedAt}
                          >
                            Restore rollback
                          </button>
                          <button
                            className="btn btn--ghost"
                            type="button"
                            onClick={() => {
                              void handleSignOut()
                            }}
                            disabled={syncLoading || authLoading || loading || autoSyncing}
                          >
                            Sign out
                          </button>
                        </div>
                      </div>
                    )}
                  </section>
                ) : null}

                {settingsSection === "feedback" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>Submit Feedback</h3>
                    <div className="note">
                      Send feedback right here. No sign-in or contact details required.
                    </div>
                    {!isSupabaseConfigured ? (
                      <div className="empty-state">
                        Feedback is not configured. Add `VITE_SUPABASE_URL` and
                        `VITE_SUPABASE_ANON_KEY` to this deployment&apos;s environment variables.
                      </div>
                    ) : (
                      <form className="form-grid" onSubmit={handleSubmitFeedback}>
                        <label>
                          Type
                          <select
                            value={feedbackCategory}
                            onChange={(event) => setFeedbackCategory(event.target.value as FeedbackCategory)}
                            disabled={feedbackSubmitting}
                          >
                            <option value="general">General feedback</option>
                            <option value="bug">Bug report</option>
                            <option value="idea">Feature idea</option>
                          </select>
                        </label>
                        <label>
                          Feedback
                          <textarea
                            value={feedbackMessage}
                            onChange={(event) => setFeedbackMessage(event.target.value)}
                            placeholder="What would you like changed or improved?"
                            rows={5}
                            maxLength={5000}
                            disabled={feedbackSubmitting}
                            required
                          />
                        </label>
                        <div className="note-text">
                          Sent anonymously. Please avoid sharing sensitive personal information.
                        </div>
                        <div className="form-actions">
                          <button
                            className="btn btn--primary"
                            type="submit"
                            disabled={feedbackSubmitting || loading}
                          >
                            {feedbackSubmitting ? "Sending..." : "Send feedback"}
                          </button>
                        </div>
                      </form>
                    )}
                    <div className="form-grid">
                      <div className="field-block">
                        <span>Data location</span>
                        <div className="note-text">
                          {userDataPath || "Local app data folder (path unavailable)."}
                        </div>
                      </div>
                    </div>
                  </section>
                ) : null}

                {settingsSection === "update" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>App updates</h3>
                    <div className="note">Check GitHub for the latest desktop release.</div>
                    <div className="note">Current version: {appVersion}</div>
                    {updateStatus ? (
                      <div className={`status status--inline ${updateStatusTone}`}>{updateStatus}</div>
                    ) : null}
                    {typeof updateProgress === "number" ? (
                      <div className="update-progress">
                        <div className="update-progress__bar">
                          <div
                            className="update-progress__fill"
                            style={{ width: `${updateProgress}%` }}
                          />
                        </div>
                        <div className="update-progress__meta">
                          {Number.isFinite(updateProgress) ? updateProgress.toFixed(0) : "0"}%
                        </div>
                      </div>
                    ) : null}
                    <div className="form-actions">
                      <button
                        className="btn btn--primary"
                        type="button"
                        onClick={handleCheckForUpdates}
                      >
                        Check for updates
                      </button>
                      {updateReady ? (
                        <button className="btn btn--ghost" type="button" onClick={handleInstallUpdate}>
                          Restart to update
                        </button>
                      ) : null}
                    </div>
                    {updateReady ? (
                      <div className="note">Update downloaded. Click "Restart to update".</div>
                    ) : null}
                  </section>
                ) : null}

                {settingsSection === "danger" ? (
                  <section className="panel panel--soft settings-section">
                    <h3>Delete All Data</h3>
                    <div className="note">
                      This clears all events, salvations, tags, and data sets stored on this device.
                    </div>
                    <div className="form-actions">
                      <button className="btn btn--danger" type="button" onClick={handleClearAllData}>
                        Clear all data
                      </button>
                    </div>
                  </section>
                ) : null}
              </div>
            </aside>
          </>
        ) : null}
      </main>
    </div>
  )
}

export default App
