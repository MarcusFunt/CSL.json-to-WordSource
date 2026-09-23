import { parseCslJson, type CslItem, type NamespaceMode } from "./cslToWord"

export const WORKSPACE_SCHEMA = "csl-word-source-workspace" as const
export const WORKSPACE_VERSION = 1 as const

export type WorkspaceSource = {
  key: string
  csl: CslItem
  selected: boolean
  addedAt?: string
  category?: string
  purpose?: string
  tags?: string[]
}

export type WorkspaceSettings = {
  namespace: NamespaceMode
  defaultLcid: number
  deduplicate: boolean
  trackAddedDates: boolean
}

export type WorkspaceFile = {
  schema: typeof WORKSPACE_SCHEMA
  version: typeof WORKSPACE_VERSION
  name: string
  exportedAt?: string
  settings: WorkspaceSettings
  sources: WorkspaceSource[]
}

export type MergeResult = {
  sources: WorkspaceSource[]
  added: number
  skipped: number
}

export const defaultSettings: WorkspaceSettings = {
  namespace: "office2006",
  defaultLcid: 1033,
  deduplicate: true,
  trackAddedDates: false,
}

function scalar(value: unknown): string | undefined {
  if (typeof value === "string") return value.trim() || undefined
  if (typeof value === "number") return String(value)
  return undefined
}

function sourceYear(item: CslItem): string {
  const rows = item.issued?.["date-parts"]
  return Array.isArray(rows) && Array.isArray(rows[0]) && rows[0][0] !== undefined
    ? String(rows[0][0])
    : ""
}

export function sourceFingerprint(item: CslItem): string {
  const doi = scalar(item.DOI)
  if (doi) return `doi:${doi.toLowerCase().replace(/^https?:\/\/(dx\.)?doi\.org\//, "")}`

  const number = scalar(item.number)
  if (item.type === "standard" && number) return `standard:${number.toLowerCase()}`

  const url = scalar(item.URL)
  if (url) return `url:${url.toLowerCase().replace(/\/$/, "")}`

  const id = scalar(item.id)
  if (id) return `id:${id.toLowerCase()}`

  const title = (scalar(item.title) ?? "").toLowerCase().replace(/\W+/g, " ").trim()
  return `title:${title}|year:${sourceYear(item)}`
}

export function sourceTitle(source: WorkspaceSource): string {
  return scalar(source.csl.title) ?? scalar(source.csl.id) ?? "Untitled source"
}

export function sourceAuthor(source: WorkspaceSource): string {
  const authors = source.csl.author
  if (!Array.isArray(authors) || authors.length === 0) return "Unknown author"
  return authors
    .map((author) => author.literal ?? [author.family, author.given].filter(Boolean).join(", "))
    .filter(Boolean)
    .join("; ")
}

export function sourceIdentifier(source: WorkspaceSource): string | undefined {
  return (
    scalar(source.csl.number) ??
    scalar(source.csl.DOI) ??
    scalar(source.csl.ISBN) ??
    scalar(source.csl.ISSN) ??
    scalar(source.csl.URL)
  )
}

function makeKey(item: CslItem): string {
  const base = (scalar(item.id) ?? scalar(item.number) ?? scalar(item.DOI) ?? scalar(item.title) ?? "source")
    .replace(/[^A-Za-z0-9_.-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 72) || "source"
  return `${base}-${globalThis.crypto.randomUUID().slice(0, 8)}`
}

export function makeWorkspaceSource(
  item: CslItem,
  trackAddedDates: boolean,
  metadata: Partial<Pick<WorkspaceSource, "category" | "purpose" | "tags">> = {},
  now = new Date().toISOString(),
): WorkspaceSource {
  return {
    key: makeKey(item),
    csl: item,
    selected: true,
    addedAt: trackAddedDates ? now : undefined,
    ...metadata,
  }
}

export function mergeCslItems(
  existing: WorkspaceSource[],
  incoming: CslItem[],
  trackAddedDates: boolean,
  now = new Date().toISOString(),
): MergeResult {
  const fingerprints = new Set(existing.map((source) => sourceFingerprint(source.csl)))
  const sources = [...existing]
  let added = 0
  let skipped = 0

  for (const item of incoming) {
    const fingerprint = sourceFingerprint(item)
    if (fingerprints.has(fingerprint)) {
      skipped += 1
      continue
    }
    fingerprints.add(fingerprint)
    sources.push(makeWorkspaceSource(item, trackAddedDates, {}, now))
    added += 1
  }

  return { sources, added, skipped }
}

export function serializeWorkspace(workspace: WorkspaceFile): string {
  return JSON.stringify(
    {
      ...workspace,
      exportedAt: new Date().toISOString(),
    },
    null,
    2,
  )
}

export function parseWorkspace(input: string): WorkspaceFile {
  const parsed: unknown = JSON.parse(input)
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("Workspace file must be a JSON object.")
  }

  const value = parsed as Partial<WorkspaceFile>
  if (value.schema !== WORKSPACE_SCHEMA || value.version !== WORKSPACE_VERSION) {
    throw new Error("This is not a supported CSL → Word workspace file.")
  }
  if (!value.settings || !Array.isArray(value.sources)) {
    throw new Error("Workspace file is missing settings or sources.")
  }

  const sources = value.sources.map((source, index) => {
    if (!source || typeof source !== "object" || !source.csl || typeof source.csl !== "object") {
      throw new Error(`Workspace source ${index + 1} is invalid.`)
    }
    return {
      ...source,
      key: typeof source.key === "string" && source.key ? source.key : makeKey(source.csl),
      selected: source.selected !== false,
    } satisfies WorkspaceSource
  })

  return {
    schema: WORKSPACE_SCHEMA,
    version: WORKSPACE_VERSION,
    name: typeof value.name === "string" && value.name.trim() ? value.name : "Imported workspace",
    exportedAt: typeof value.exportedAt === "string" ? value.exportedAt : undefined,
    settings: {
      namespace: value.settings.namespace === "legacy2004" ? "legacy2004" : "office2006",
      defaultLcid: Number(value.settings.defaultLcid) || 1033,
      deduplicate: value.settings.deduplicate !== false,
      trackAddedDates: value.settings.trackAddedDates === true,
    },
    sources,
  }
}

export function parseCslOrWorkspace(input: string):
  | { kind: "workspace"; workspace: WorkspaceFile }
  | { kind: "csl"; items: CslItem[] } {
  try {
    const parsed = JSON.parse(input) as { schema?: unknown }
    if (parsed && typeof parsed === "object" && parsed.schema === WORKSPACE_SCHEMA) {
      return { kind: "workspace", workspace: parseWorkspace(input) }
    }
  } catch {
    // parseCslJson below will provide the useful JSON error.
  }
  return { kind: "csl", items: parseCslJson(input) }
}

const projectItems: Array<{
  csl: CslItem
  category: string
  purpose: string
  tags: string[]
}> = [
  {
    category: "Core",
    purpose: "Coupon and test-specimen geometry",
    tags: ["coupons", "geometry", "tensile"],
    csl: {
      id: "JIS-K-7139-2026",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Test specimens",
      number: "JIS K 7139:2026",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2026, 2, 20]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+K+7139%3A2026",
      language: "en",
      note: "Test-specimen geometries including A1/A2 dogbones, reduced Axy specimens, B bars, C small tensile specimens, and plates. Identical to ISO 20753:2023.",
    },
  },
  {
    category: "Core",
    purpose: "Tensile testing — general principles",
    tags: ["tensile", "stress", "strain", "modulus"],
    csl: {
      id: "JIS-K-7161-1-2024",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Determination of tensile properties — Part 1: General principles",
      number: "JIS K 7161-1:2024",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2024, 3, 21]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+K+7161-1%3A2024",
      language: "en",
      note: "General tensile-testing principles: stress, strain, modulus, speed, extensometry, and calculations. Identical to ISO 527-1:2019.",
    },
  },
  {
    category: "Core",
    purpose: "Tensile testing — moulded/extruded plastics",
    tags: ["tensile", "plastics"],
    csl: {
      id: "JIS-K-7161-2-2014",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Determination of tensile properties — Part 2: Test conditions for moulding and extrusion plastics",
      number: "JIS K 7161-2:2014",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2014, 9, 22]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+K+7161-2%3A2014",
      language: "en",
      note: "Tensile-test conditions for moulded and extruded plastics; companion to JIS K 7161-1. Identical to ISO 527-2:2012.",
    },
  },
  {
    category: "Core",
    purpose: "Three-point bending",
    tags: ["bending", "flexural", "modulus"],
    csl: {
      id: "JIS-K-7171-2022",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Determination of flexural properties",
      number: "JIS K 7171:2022",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2022, 1, 20]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+K+7171%3A2022",
      language: "en",
      note: "Three-point bending: flexural strength, modulus, and stress–strain behaviour. Identical to ISO 178:2019.",
    },
  },
  {
    category: "Core",
    purpose: "Charpy impact testing",
    tags: ["charpy", "impact"],
    csl: {
      id: "JIS-K-7111-1-2012",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Determination of Charpy impact properties — Part 1: Non-instrumented impact test",
      number: "JIS K 7111-1:2012",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2012, 12, 20]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/index/?bunsyo_id=JIS+K+7111-1%3A2012",
      language: "en",
      note: "Conventional non-instrumented Charpy impact testing. Modified adoption of ISO 179-1:2010.",
    },
  },
  {
    category: "Supporting",
    purpose: "Specimen conditioning and test atmosphere",
    tags: ["conditioning", "temperature", "humidity"],
    csl: {
      id: "JIS-K-7100-1999",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Plastics — Standard atmospheres for conditioning and testing",
      number: "JIS K 7100:1999",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[1999]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+K+7100%3A1999",
      language: "en",
      note: "Specimen conditioning and test atmosphere, including temperature and humidity before and during testing.",
    },
  },
  {
    category: "Machine verification",
    purpose: "Force-measurement calibration",
    tags: ["calibration", "force", "load-cell"],
    csl: {
      id: "JIS-B-7721-2018",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Tension/compression testing machines — Calibration and verification of the force-measuring system",
      number: "JIS B 7721:2018",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2018, 1, 22]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+B+7721%3A2018",
      language: "en",
      note: "Calibration and verification of the tensile/compression machine force-measuring system. Identical to ISO 7500-1:2015.",
    },
  },
  {
    category: "Machine verification",
    purpose: "Extensometer calibration",
    tags: ["calibration", "extensometer", "camera"],
    csl: {
      id: "JIS-B-7741-2019",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Calibration of extensometer systems used in uniaxial testing",
      number: "JIS B 7741:2019",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2019, 11, 20]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+B+7741%3A2019",
      language: "en",
      note: "Calibration of contact and non-contact extensometer systems; relevant to camera-based extensometry. Corresponds to ISO 9513:2012.",
    },
  },
  {
    category: "Machine verification",
    purpose: "Impact-machine verification",
    tags: ["calibration", "impact", "pendulum"],
    csl: {
      id: "JIS-B-7739-2020",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "Pendulum-type impact-testing machines for non-metallic materials — Verification of testing machines",
      number: "JIS B 7739:2020",
      publisher: "Japanese Standards Association",
      issued: { "date-parts": [[2020, 8, 20]] },
      URL: "https://webdesk.jsa.or.jp/books/W11M0090/?bunsyo_id=JIS+B+7739%3A2020",
      language: "en",
      note: "Verification of pendulum impact-testing machines used for Charpy, Izod, and tensile-impact testing.",
    },
  },
]

export function createProjectWorkspace(): WorkspaceFile {
  return {
    schema: WORKSPACE_SCHEMA,
    version: WORKSPACE_VERSION,
    name: "FDM strength project",
    settings: { ...defaultSettings },
    sources: projectItems.map((entry, index) => ({
      key: `project-${index + 1}-${String(entry.csl.id)}`,
      csl: structuredClone(entry.csl),
      selected: true,
      category: entry.category,
      purpose: entry.purpose,
      tags: [...entry.tags],
    })),
  }
}

export function createEmptyWorkspace(): WorkspaceFile {
  return {
    schema: WORKSPACE_SCHEMA,
    version: WORKSPACE_VERSION,
    name: "Untitled workspace",
    settings: { ...defaultSettings },
    sources: [],
  }
}
