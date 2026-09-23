import { describe, expect, it } from "vitest"
import { convertCslJson, parseCslJson } from "./cslToWord"

describe("parseCslJson", () => {
  it("accepts a single object", () => {
    expect(parseCslJson('{"type":"book","title":"Test"}')).toHaveLength(1)
  })

  it("accepts an items wrapper", () => {
    expect(parseCslJson('{"items":[{"type":"book","title":"Test"}]}')).toHaveLength(1)
  })
})

describe("convertCslJson", () => {
  it("maps journal articles and structured authors", () => {
    const result = convertCslJson(JSON.stringify([{
      id: "smith2025",
      type: "article-journal",
      author: [{ family: "Smith", given: "Alex" }],
      title: "A paper",
      "container-title": "A Journal",
      issued: { "date-parts": [[2025, 4, 12]] },
      DOI: "10.1234/test"
    }]))

    expect(result.itemCount).toBe(1)
    expect(result.xml).toContain("<b:SourceType>JournalArticle</b:SourceType>")
    expect(result.xml).toContain("<b:Last>Smith</b:Last>")
    expect(result.xml).toContain("<b:First>Alex</b:First>")
    expect(result.xml).toContain("<b:JournalName>A Journal</b:JournalName>")
    expect(result.xml).toContain("<b:Month>April</b:Month>")
    expect(result.xml).toContain("<b:DOI>10.1234/test</b:DOI>")
  })

  it("preserves a single corporate author", () => {
    const result = convertCslJson(JSON.stringify({
      id: "std",
      type: "standard",
      author: [{ literal: "Japanese Standards Association" }],
      title: "A standard",
      number: "JIS X 0000:2026"
    }))

    expect(result.xml).toContain("<b:Corporate>Japanese Standards Association</b:Corporate>")
    expect(result.xml).toContain("<b:StandardNumber>JIS X 0000:2026</b:StandardNumber>")
  })

  it("deduplicates by DOI", () => {
    const input = JSON.stringify([
      { type: "article-journal", title: "A", DOI: "10.1/x" },
      { type: "article-journal", title: "A copy", DOI: "https://doi.org/10.1/x" }
    ])
    const result = convertCslJson(input, { deduplicate: true })
    expect(result.itemCount).toBe(1)
    expect(result.skippedDuplicates).toBe(1)
  })

  it("supports the legacy namespace", () => {
    const result = convertCslJson('{"type":"book","title":"Test"}', { namespace: "legacy2004" })
    expect(result.xml).toContain("http://schemas.microsoft.com/office/word/2004/10/bibliography")
  })
})
