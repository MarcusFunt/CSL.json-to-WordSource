export type CslName = {
  family?: string
  given?: string
  suffix?: string
  literal?: string
  "non-dropping-particle"?: string
  "dropping-particle"?: string
}

export type CslDate = {
  "date-parts"?: Array<Array<string | number>>
  raw?: string
}

export type CslItem = {
  id?: string | number
  type?: string
  title?: string
  "title-short"?: string
  "container-title"?: string
  "container-title-short"?: string
  "event-title"?: string
  "collection-title"?: string
  author?: CslName[]
  editor?: CslName[]
  translator?: CslName[]
  compiler?: CslName[]
  composer?: CslName[]
  director?: CslName[]
  interviewer?: CslName[]
  interviewee?: CslName[]
  recipient?: CslName[]
  issued?: CslDate
  accessed?: CslDate
  volume?: string | number
  issue?: string | number
  page?: string | number
  "chapter-number"?: string | number
  edition?: string | number
  "number-of-volumes"?: string | number
  publisher?: string
  "publisher-place"?: string
  department?: string
  institution?: string
  authority?: string
  version?: string
  medium?: string
  genre?: string
  number?: string | number
  ISBN?: string
  ISSN?: string
  DOI?: string
  URL?: string
  language?: string
  note?: string
  annote?: string
  [key: string]: unknown
}

export type NamespaceMode = "office2006" | "legacy2004"

export type ConversionOptions = {
  namespace?: NamespaceMode
  defaultLcid?: number
  deduplicate?: boolean
}

export type ConversionResult = {
  xml: string
  itemCount: number
  skippedDuplicates: number
  warnings: string[]
}

const NAMESPACES: Record<NamespaceMode, string> = {
  office2006: "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
  legacy2004: "http://schemas.microsoft.com/office/word/2004/10/bibliography",
}

const TYPE_MAP: Record<string, string> = {
  article: "ArticleInAPeriodical",
  "article-journal": "JournalArticle",
  "article-magazine": "ArticleInAPeriodical",
  "article-newspaper": "ArticleInAPeriodical",
  bill: "Misc",
  book: "Book",
  broadcast: "ElectronicSource",
  chapter: "BookSection",
  classic: "Book",
  collection: "Book",
  dataset: "ElectronicSource",
  document: "Misc",
  entry: "BookSection",
  "entry-dictionary": "BookSection",
  "entry-encyclopedia": "BookSection",
  event: "ConferenceProceedings",
  figure: "Art",
  graphic: "Art",
  hearing: "Case",
  interview: "Interview",
  legal_case: "Case",
  legislation: "Misc",
  manuscript: "Report",
  map: "Art",
  motion_picture: "Film",
  musical_score: "Art",
  pamphlet: "Report",
  "paper-conference": "ConferenceProceedings",
  patent: "Patent",
  performance: "Performance",
  periodical: "ArticleInAPeriodical",
  personal_communication: "Interview",
  post: "DocumentFromInternetSite",
  "post-weblog": "DocumentFromInternetSite",
  regulation: "Misc",
  report: "Report",
  review: "ArticleInAPeriodical",
  "review-book": "ArticleInAPeriodical",
  software: "ElectronicSource",
  song: "SoundRecording",
  speech: "Performance",
  standard: "Report",
  thesis: "Report",
  treaty: "Misc",
  webpage: "DocumentFromInternetSite",
}

const LANGUAGE_TO_LCID: Record<string, number> = {
  da: 1030,
  "da-dk": 1030,
  de: 1031,
  "de-de": 1031,
  en: 1033,
  "en-us": 1033,
  "en-gb": 2057,
  fr: 1036,
  "fr-fr": 1036,
  es: 3082,
  "es-es": 3082,
  it: 1040,
  "it-it": 1040,
  nl: 1043,
  "nl-nl": 1043,
  nb: 1044,
  "nb-no": 1044,
  no: 1044,
  sv: 1053,
  "sv-se": 1053,
}

const ROLE_MAP: Record<string, string> = {
  author: "Author",
  editor: "Editor",
  translator: "Translator",
  compiler: "Compiler",
  composer: "Composer",
  director: "Director",
  interviewer: "Interviewer",
  interviewee: "Interviewee",
  recipient: "Interviewee",
}

function scalar(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined
  if (typeof value === "string") return value.trim() || undefined
  if (typeof value === "number" || typeof value === "boolean") return String(value)
  return undefined
}

function first(item: CslItem, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = scalar(item[key])
    if (value) return value
  }
  return undefined
}

function esc(value: unknown): string {
  const text = scalar(value) ?? ""
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;")
}

function element(name: string, value: unknown, indent = 4): string {
  const text = scalar(value)
  if (!text) return ""
  const pad = " ".repeat(indent)
  return `${pad}<b:${name}>${esc(text)}</b:${name}>\n`
}

function dateParts(value: unknown): [string?, string?, string?] {
  if (!value || typeof value !== "object") return []
  const date = value as CslDate
  const rows = date["date-parts"]
  const row = Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : undefined
  if (!row) return []
  const year = row[0] !== undefined ? String(row[0]) : undefined
  let month: string | undefined
  if (row[1] !== undefined) {
    const numeric = Number(row[1])
    const months = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December",
    ]
    month = Number.isInteger(numeric) && numeric >= 1 && numeric <= 12 ? months[numeric - 1] : String(row[1])
  }
  const day = row[2] !== undefined ? String(row[2]) : undefined
  return [year, month, day]
}

function lcidFor(item: CslItem, fallback: number): number {
  const language = scalar(item.language)
  if (!language) return fallback
  const normalized = language.toLowerCase().replaceAll("_", "-")
  return LANGUAGE_TO_LCID[normalized] ?? LANGUAGE_TO_LCID[normalized.split("-")[0]] ?? fallback
}

function normalizeTag(raw: string): string {
  let tag = raw.trim().replace(/\s+/g, "_").replace(/[^A-Za-z0-9_.-]/g, "").replace(/^[._-]+|[._-]+$/g, "")
  if (!tag) tag = "src"
  if (/^\d/.test(tag)) tag = `src_${tag}`
  return tag.slice(0, 120)
}

function makeTag(item: CslItem, used: Set<string>): string {
  let candidate = scalar(item.id)
  if (!candidate) {
    const firstAuthor = Array.isArray(item.author) ? item.author[0] : undefined
    const author = firstAuthor?.family ?? firstAuthor?.literal ?? ""
    const [year] = dateParts(item.issued)
    const firstWord = (scalar(item.title)?.split(/\s+/)[0] ?? "source").replace(/\W+/g, "")
    candidate = `${author}${year ?? ""}${firstWord}`
  }

  const base = normalizeTag(candidate)
  let tag = base
  let index = 2
  while (used.has(tag)) {
    const suffix = `_${index}`
    tag = `${base.slice(0, Math.max(1, 120 - suffix.length))}${suffix}`
    index += 1
  }
  used.add(tag)
  return tag
}

function makeGuid(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return `{${crypto.randomUUID().toUpperCase()}}`
  }
  const values = new Uint8Array(16)
  crypto.getRandomValues(values)
  values[6] = (values[6] & 0x0f) | 0x40
  values[8] = (values[8] & 0x3f) | 0x80
  const hex = Array.from(values, (value) => value.toString(16).padStart(2, "0")).join("")
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`.toUpperCase()
}

function renderPerson(person: CslName): string {
  let family = scalar(person.family)
  let given = scalar(person.given)
  const nonDropping = scalar(person["non-dropping-particle"])
  const dropping = scalar(person["dropping-particle"])
  const literal = scalar(person.literal)

  if (nonDropping) family = `${nonDropping} ${family ?? ""}`.trim()
  if (dropping) given = `${given ?? ""} ${dropping}`.trim()
  if (literal && !family && !given) family = literal

  return [
    "          <b:Person>",
    family ? `            <b:Last>${esc(family)}</b:Last>` : "",
    given ? `            <b:First>${esc(given)}</b:First>` : "",
    scalar(person.suffix) ? `            <b:Suffix>${esc(person.suffix)}</b:Suffix>` : "",
    "          </b:Person>",
  ].filter(Boolean).join("\n")
}

function renderRole(roleName: string, people: unknown, warnings: string[], label: string): string {
  if (!Array.isArray(people) || people.length === 0) return ""
  const valid = people.filter((person): person is CslName => !!person && typeof person === "object")
  if (!valid.length) return ""

  if (valid.length === 1 && scalar(valid[0].literal)) {
    return [
      `      <b:${roleName}>`,
      `        <b:Corporate>${esc(valid[0].literal)}</b:Corporate>`,
      `      </b:${roleName}>`,
    ].join("\n")
  }

  if (valid.some((person) => scalar(person.literal))) {
    warnings.push(`${label}: mixed corporate/personal names in ${roleName}; corporate literals are preserved as Last names.`)
  }

  return [
    `      <b:${roleName}>`,
    "        <b:NameList>",
    ...valid.map(renderPerson),
    "        </b:NameList>",
    `      </b:${roleName}>`,
  ].join("\n")
}

function renderAuthors(item: CslItem, warnings: string[], label: string): string {
  const roles: Array<[string, unknown]> = []
  if (item.type === "patent" && Array.isArray(item.author)) {
    roles.push(["Inventor", item.author])
  } else {
    for (const [cslRole, wordRole] of Object.entries(ROLE_MAP)) {
      if (Array.isArray(item[cslRole])) roles.push([wordRole, item[cslRole]])
    }
  }

  const rendered = roles
    .map(([role, people]) => renderRole(role, people, warnings, label))
    .filter(Boolean)

  if (!rendered.length) return ""
  return ["    <b:Author>", ...rendered, "    </b:Author>"].join("\n") + "\n"
}

function containerFields(item: CslItem, wordType: string): string {
  const container = first(item, "container-title", "container-title-short")
  const eventTitle = scalar(item["event-title"])
  let output = ""

  if (wordType === "JournalArticle") output += element("JournalName", container)
  else if (wordType === "ArticleInAPeriodical") output += element("PeriodicalTitle", container)
  else if (wordType === "BookSection") output += element("BookTitle", container)
  else if (wordType === "ConferenceProceedings") output += element("ConferenceName", eventTitle ?? container)
  else if (wordType === "DocumentFromInternetSite" || wordType === "InternetSite") {
    output += element("InternetSiteTitle", container)
  }

  output += element("PublicationTitle", item["collection-title"])
  return output
}

function standardNumber(item: CslItem, wordType: string): string | undefined {
  if (wordType === "Book" || wordType === "BookSection") return first(item, "ISBN", "ISSN", "number")
  if (wordType === "JournalArticle" || wordType === "ArticleInAPeriodical") return first(item, "ISSN", "ISBN", "number")
  return first(item, "ISBN", "ISSN", "number")
}

function duplicateKey(item: CslItem): string {
  const doi = scalar(item.DOI)
  if (doi) return `doi:${doi.toLowerCase().replace(/^https?:\/\/(dx\.)?doi\.org\//, "")}`
  const url = scalar(item.URL)
  if (url) return `url:${url.toLowerCase().replace(/\/$/, "")}`
  const title = (scalar(item.title) ?? "").toLowerCase().replace(/\W+/g, " ").trim()
  const [year] = dateParts(item.issued)
  return `title:${title}|year:${year ?? ""}`
}

function dedupe(items: CslItem[]): { items: CslItem[]; skipped: number } {
  const seen = new Set<string>()
  const result: CslItem[] = []
  let skipped = 0
  for (const item of items) {
    const key = duplicateKey(item)
    if (seen.has(key)) {
      skipped += 1
      continue
    }
    seen.add(key)
    result.push(item)
  }
  return { items: result, skipped }
}

function convertItem(
  item: CslItem,
  usedTags: Set<string>,
  defaultLcid: number,
  warnings: string[],
  index: number,
): string {
  const cslType = scalar(item.type) ?? "document"
  const wordType = TYPE_MAP[cslType] ?? "Misc"
  const label = scalar(item.title) ?? scalar(item.id) ?? `Source ${index}`
  if (!(cslType in TYPE_MAP)) warnings.push(`${label}: unknown CSL type "${cslType}", mapped to Word Misc.`)

  const tag = makeTag(item, usedTags)
  const [year, month, day] = dateParts(item.issued)
  const [yearAccessed, monthAccessed, dayAccessed] = dateParts(item.accessed)

  let body = ""
  body += element("Tag", tag)
  body += element("SourceType", wordType)
  body += element("Guid", makeGuid())
  body += element("LCID", lcidFor(item, defaultLcid))
  body += renderAuthors(item, warnings, label)
  body += element("Title", item.title)
  body += element("ShortTitle", item["title-short"])
  body += element("Year", year)
  body += element("Month", month)
  body += element("Day", day)
  body += containerFields(item, wordType)
  body += element("Volume", item.volume)
  body += element("Issue", item.issue)
  body += element("Pages", item.page)
  body += element("ChapterNumber", item["chapter-number"])
  body += element("Edition", item.edition)
  body += element("NumberVolumes", item["number-of-volumes"])
  body += element("Publisher", item.publisher)
  body += element("City", item["publisher-place"])
  body += element("Department", item.department)
  body += element("Institution", first(item, "institution", "authority"))
  body += element("Version", item.version)
  body += element("Medium", item.medium)

  if (cslType === "thesis") body += element("ThesisType", first(item, "genre"))
  if (cslType === "patent") {
    body += element("PatentNumber", item.number)
    body += element("Type", item.genre)
  }
  if (cslType === "legal_case") {
    body += element("Court", item.authority)
    body += element("CaseNumber", item.number)
  }

  body += element("StandardNumber", standardNumber(item, wordType))
  body += element("DOI", item.DOI)
  body += element("URL", item.URL)
  body += element("YearAccessed", yearAccessed)
  body += element("MonthAccessed", monthAccessed)
  body += element("DayAccessed", dayAccessed)
  body += element("Comments", first(item, "note", "annote"))

  return `  <b:Source>\n${body}  </b:Source>`
}

export function parseCslJson(input: string): CslItem[] {
  const parsed: unknown = JSON.parse(input)

  let items: unknown
  if (Array.isArray(parsed)) items = parsed
  else if (parsed && typeof parsed === "object" && Array.isArray((parsed as { items?: unknown }).items)) {
    items = (parsed as { items: unknown[] }).items
  } else if (parsed && typeof parsed === "object") {
    items = [parsed]
  } else {
    throw new Error("CSL-JSON must be an object, an array of objects, or an object with an items array.")
  }

  if (!Array.isArray(items) || items.length === 0) throw new Error("No CSL items found.")
  if (items.some((item) => !item || typeof item !== "object" || Array.isArray(item))) {
    throw new Error("Every CSL item must be a JSON object.")
  }

  return items as CslItem[]
}

export function convertCslJson(input: string, options: ConversionOptions = {}): ConversionResult {
  const namespaceMode = options.namespace ?? "office2006"
  const namespace = NAMESPACES[namespaceMode]
  const defaultLcid = options.defaultLcid ?? 1033
  const warnings: string[] = []

  let items = parseCslJson(input)
  let skippedDuplicates = 0
  if (options.deduplicate ?? true) {
    const result = dedupe(items)
    items = result.items
    skippedDuplicates = result.skipped
  }

  if (skippedDuplicates) warnings.push(`Skipped ${skippedDuplicates} duplicate source${skippedDuplicates === 1 ? "" : "s"}.`)

  const usedTags = new Set<string>()
  const sources = items.map((item, index) => convertItem(item, usedTags, defaultLcid, warnings, index + 1))

  const xml = [
    '<?xml version="1.0" encoding="utf-8"?>',
    `<b:Sources xmlns:b="${namespace}" SelectedStyle="">`,
    ...sources,
    "</b:Sources>",
    "",
  ].join("\n")

  return { xml, itemCount: items.length, skippedDuplicates, warnings }
}

export const exampleCslJson = JSON.stringify(
  [
    {
      id: "JIS-K-7161-1-2024",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Determination of tensile properties — Part 1: General principles",
      number: "JIS K 7161-1:2024",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2024, 3, 21]] },
      URL: "https://webdesk.jsa.or.jp/",
      language: "en",
    },
    {
      id: "Smith2025",
      type: "article-journal",
      author: [
        { family: "Smith", given: "Alex J." },
        { family: "Jensen", given: "Peter" },
      ],
      title: "Mechanical behaviour of FDM printed PLA",
      "container-title": "Additive Manufacturing",
      issued: { "date-parts": [[2025, 4, 12]] },
      volume: "42",
      issue: "3",
      page: "100-112",
      DOI: "10.1234/example.2025.001",
      language: "en",
    },
  ],
  null,
  2,
)
