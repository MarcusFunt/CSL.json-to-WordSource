import { describe, expect, it } from "vitest"
import {
  createProjectWorkspace,
  mergeCslItems,
  parseWorkspace,
  serializeWorkspace,
  sourceFingerprint,
} from "./workspace"

describe("workspace files", () => {
  it("round-trips a workspace", () => {
    const workspace = createProjectWorkspace()
    workspace.settings.trackAddedDates = true
    const parsed = parseWorkspace(serializeWorkspace(workspace))
    expect(parsed.name).toBe("FDM strength project")
    expect(parsed.sources).toHaveLength(9)
    expect(parsed.settings.trackAddedDates).toBe(true)
  })

  it("project preset contains the nine selected JIS standards", () => {
    const workspace = createProjectWorkspace()
    expect(workspace.sources).toHaveLength(9)
    expect(workspace.sources.map((source) => source.csl.number)).toContain("JIS K 7139:2026")
    expect(workspace.sources.map((source) => source.csl.number)).toContain("JIS B 7739:2020")
  })

  it("records addedAt only when tracking is enabled", () => {
    const item = { id: "x", type: "book", title: "Test" }
    const tracked = mergeCslItems([], [item], true, "2026-09-23T10:00:00.000Z")
    const untracked = mergeCslItems([], [item], false, "2026-09-23T10:00:00.000Z")
    expect(tracked.sources[0].addedAt).toBe("2026-09-23T10:00:00.000Z")
    expect(untracked.sources[0].addedAt).toBeUndefined()
  })

  it("deduplicates standards by standard number", () => {
    const a = { id: "a", type: "standard", number: "JIS X 1:2026", title: "One" }
    const b = { id: "b", type: "standard", number: "JIS X 1:2026", title: "Duplicate" }
    expect(sourceFingerprint(a)).toBe(sourceFingerprint(b))
    const result = mergeCslItems([], [a, b], false)
    expect(result.added).toBe(1)
    expect(result.skipped).toBe(1)
  })
})
