import { useMemo, useRef, useState } from "react"
import {
  CalendarClock,
  CheckCircle2,
  Clipboard,
  Database,
  Download,
  FileDown,
  FileJson,
  FolderOpen,
  Github,
  Plus,
  RotateCcw,
  Search,
  Trash2,
  Upload,
  XCircle,
} from "lucide-react"
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { convertCslJson, type ConversionResult, type NamespaceMode } from "@/lib/cslToWord"
import {
  createEmptyWorkspace,
  createProjectWorkspace,
  mergeCslItems,
  parseCslOrWorkspace,
  serializeWorkspace,
  sourceAuthor,
  sourceIdentifier,
  sourceTitle,
  type WorkspaceFile,
  type WorkspaceSource,
} from "@/lib/workspace"

function downloadText(filename: string, content: string, mimeType: string) {
  const blob = new Blob([content], { type: mimeType })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

function safeFilename(value: string) {
  return value.trim().replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "workspace"
}

function formatAddedAt(value?: string) {
  if (!value) return "Not tracked"
  const date = new Date(value)
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString()
}

export default function App() {
  const [workspace, setWorkspace] = useState<WorkspaceFile>(() => createProjectWorkspace())
  const [importText, setImportText] = useState("")
  const [search, setSearch] = useState("")
  const [result, setResult] = useState<ConversionResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [status, setStatus] = useState<string | null>("Loaded the FDM project library with 9 JIS standards.")
  const [copied, setCopied] = useState(false)
  const cslFileRef = useRef<HTMLInputElement>(null)
  const workspaceFileRef = useRef<HTMLInputElement>(null)

  const selectedSources = useMemo(
    () => workspace.sources.filter((source) => source.selected),
    [workspace.sources],
  )

  const visibleSources = useMemo(() => {
    const needle = search.trim().toLowerCase()
    if (!needle) return workspace.sources
    return workspace.sources.filter((source) => {
      const haystack = [
        sourceTitle(source),
        sourceAuthor(source),
        sourceIdentifier(source),
        source.category,
        source.purpose,
        ...(source.tags ?? []),
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
      return haystack.includes(needle)
    })
  }, [search, workspace.sources])

  const updateSettings = (patch: Partial<WorkspaceFile["settings"]>) => {
    setWorkspace((current) => ({
      ...current,
      settings: { ...current.settings, ...patch },
    }))
    setResult(null)
  }

  const updateSource = (key: string, patch: Partial<WorkspaceSource>) => {
    setWorkspace((current) => ({
      ...current,
      sources: current.sources.map((source) => (source.key === key ? { ...source, ...patch } : source)),
    }))
    setResult(null)
  }

  const importPayload = (text: string, replaceWorkspace = false) => {
    const parsed = parseCslOrWorkspace(text)
    if (parsed.kind === "workspace") {
      setWorkspace(parsed.workspace)
      setImportText("")
      setResult(null)
      setError(null)
      setStatus(`Opened workspace "${parsed.workspace.name}" with ${parsed.workspace.sources.length} sources.`)
      return
    }

    if (replaceWorkspace) {
      const empty = createEmptyWorkspace()
      const merged = mergeCslItems([], parsed.items, empty.settings.trackAddedDates)
      setWorkspace({ ...empty, sources: merged.sources })
      setStatus(`Created a new workspace with ${merged.added} imported sources.`)
    } else {
      const merged = mergeCslItems(
        workspace.sources,
        parsed.items,
        workspace.settings.trackAddedDates,
      )
      setWorkspace((current) => ({ ...current, sources: merged.sources }))
      setStatus(
        `Added ${merged.added} source${merged.added === 1 ? "" : "s"}${merged.skipped ? `; skipped ${merged.skipped} duplicate${merged.skipped === 1 ? "" : "s"}` : ""}.`,
      )
    }
    setImportText("")
    setResult(null)
    setError(null)
  }

  const readFile = async (file?: File) => {
    if (!file) return
    try {
      importPayload(await file.text())
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Could not import that file.")
    }
  }

  const buildWordXml = (download = false) => {
    try {
      if (selectedSources.length === 0) throw new Error("Select at least one source to export.")
      const next = convertCslJson(JSON.stringify(selectedSources.map((source) => source.csl)), {
        namespace: workspace.settings.namespace,
        defaultLcid: workspace.settings.defaultLcid,
        deduplicate: workspace.settings.deduplicate,
      })
      setResult(next)
      setError(null)
      setStatus(`Built Word XML for ${next.itemCount} selected source${next.itemCount === 1 ? "" : "s"}.`)
      if (download) downloadText("Sources.xml", next.xml, "application/xml;charset=utf-8")
    } catch (reason) {
      setResult(null)
      setError(reason instanceof Error ? reason.message : "Could not create Word XML.")
    }
  }

  const exportWorkspace = () => {
    downloadText(
      `${safeFilename(workspace.name)}.cslword-workspace.json`,
      serializeWorkspace(workspace),
      "application/json;charset=utf-8",
    )
    setStatus("Workspace configuration downloaded.")
  }

  const exportCsl = () => {
    downloadText(
      `${safeFilename(workspace.name)}.csl.json`,
      JSON.stringify(workspace.sources.map((source) => source.csl), null, 2),
      "application/json;charset=utf-8",
    )
  }

  const copyXml = async () => {
    if (!result) return
    await navigator.clipboard.writeText(result.xml)
    setCopied(true)
    window.setTimeout(() => setCopied(false), 1500)
  }

  const setAllSelected = (selected: boolean) => {
    setWorkspace((current) => ({
      ...current,
      sources: current.sources.map((source) => ({ ...source, selected })),
    }))
    setResult(null)
  }

  return (
    <main className="min-h-screen bg-neutral-50">
      <div className="mx-auto max-w-7xl px-4 py-8 sm:px-6 lg:px-8">
        <header className="mb-6 flex flex-col gap-4 border-b border-neutral-200 pb-6 md:flex-row md:items-end md:justify-between">
          <div className="max-w-3xl">
            <div className="mb-3 flex flex-wrap items-center gap-2">
              <Badge>CSL-JSON</Badge>
              <Badge>Word Sources.xml</Badge>
              <Badge>Workspace files</Badge>
            </div>
            <h1 className="text-3xl font-bold tracking-tight text-neutral-950 sm:text-4xl">
              CSL.json → Word Sources.xml
            </h1>
            <p className="mt-2 text-sm leading-6 text-neutral-600">
              Manage a portable source library, import CSL-JSON, and export selected sources to Word.
            </p>
          </div>
          <Button variant="outline" asChild>
            <a href="https://github.com/MarcusFunt/CSL.json-to-WordSource" target="_blank" rel="noreferrer">
              <Github className="size-4" />
              GitHub
            </a>
          </Button>
        </header>

        <Card className="mb-6">
          <CardContent className="grid gap-4 pt-6 lg:grid-cols-[1fr_auto] lg:items-end">
            <div className="space-y-2">
              <Label htmlFor="workspace-name">Workspace name</Label>
              <input
                id="workspace-name"
                value={workspace.name}
                onChange={(event) => setWorkspace((current) => ({ ...current, name: event.target.value }))}
                className="h-10 w-full rounded-md border border-neutral-200 bg-white px-3 text-sm shadow-sm outline-none focus:ring-2 focus:ring-neutral-400"
              />
              <p className="text-xs text-neutral-500">
                {workspace.sources.length} sources · {selectedSources.length} selected for Word export
              </p>
            </div>
            <div className="flex flex-wrap gap-2">
              <Button variant="outline" onClick={() => workspaceFileRef.current?.click()}>
                <FolderOpen className="size-4" />
                Open workspace
              </Button>
              <input
                ref={workspaceFileRef}
                className="hidden"
                type="file"
                accept=".json,application/json"
                onChange={(event) => void readFile(event.target.files?.[0])}
              />
              <Button variant="outline" onClick={exportWorkspace}>
                <FileDown className="size-4" />
                Save workspace
              </Button>
              <Button onClick={() => buildWordXml(true)}>
                <Download className="size-4" />
                Export Word XML
              </Button>
            </div>
          </CardContent>
        </Card>

        {status && (
          <div className="mb-6 flex items-center gap-2 rounded-lg border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">
            <CheckCircle2 className="size-4 shrink-0" />
            {status}
          </div>
        )}

        {error && (
          <div className="mb-6 flex items-center gap-2 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">
            <XCircle className="size-4 shrink-0" />
            {error}
          </div>
        )}

        <div className="grid gap-6 lg:grid-cols-[minmax(0,1fr)_340px]">
          <Card>
            <CardHeader>
              <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
                <div>
                  <CardTitle className="flex items-center gap-2">
                    <Database className="size-5" />
                    Source library
                  </CardTitle>
                  <CardDescription className="mt-2">
                    Keep sources in the workspace even when they are not selected for Word export.
                  </CardDescription>
                </div>
                <div className="relative min-w-64">
                  <Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-neutral-400" />
                  <input
                    value={search}
                    onChange={(event) => setSearch(event.target.value)}
                    placeholder="Search sources…"
                    className="h-10 w-full rounded-md border border-neutral-200 bg-white pl-9 pr-3 text-sm shadow-sm outline-none focus:ring-2 focus:ring-neutral-400"
                  />
                </div>
              </div>
            </CardHeader>
            <CardContent>
              <div className="mb-4 flex flex-wrap gap-2">
                <Button size="sm" variant="outline" onClick={() => setAllSelected(true)}>Select all</Button>
                <Button size="sm" variant="outline" onClick={() => setAllSelected(false)}>Select none</Button>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => {
                    setWorkspace(createProjectWorkspace())
                    setResult(null)
                    setError(null)
                    setStatus("Restored the 9-source FDM project preset.")
                  }}
                >
                  <RotateCcw className="size-4" />
                  Load FDM preset
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  onClick={() => {
                    setWorkspace(createEmptyWorkspace())
                    setResult(null)
                    setStatus("Started an empty workspace.")
                  }}
                >
                  New empty workspace
                </Button>
              </div>

              <div className="space-y-3">
                {visibleSources.length === 0 ? (
                  <div className="rounded-lg border border-dashed border-neutral-300 p-8 text-center text-sm text-neutral-500">
                    No matching sources.
                  </div>
                ) : (
                  visibleSources.map((source) => (
                    <div key={source.key} className="rounded-lg border border-neutral-200 bg-white p-4">
                      <div className="flex gap-3">
                        <input
                          className="mt-1 size-4 shrink-0"
                          type="checkbox"
                          checked={source.selected}
                          onChange={(event) => updateSource(source.key, { selected: event.target.checked })}
                          aria-label={`Select ${sourceTitle(source)}`}
                        />
                        <div className="min-w-0 flex-1">
                          <div className="flex flex-wrap items-start justify-between gap-2">
                            <div>
                              <h3 className="font-medium leading-6 text-neutral-950">{sourceTitle(source)}</h3>
                              <p className="mt-1 text-xs text-neutral-500">{sourceAuthor(source)}</p>
                            </div>
                            <div className="flex flex-wrap gap-1">
                              {source.category && <Badge>{source.category}</Badge>}
                              {source.csl.type && <Badge>{String(source.csl.type)}</Badge>}
                            </div>
                          </div>

                          {sourceIdentifier(source) && (
                            <p className="mt-2 break-all font-mono text-xs text-neutral-600">{sourceIdentifier(source)}</p>
                          )}
                          {source.purpose && <p className="mt-2 text-sm text-neutral-700">{source.purpose}</p>}

                          <div className="mt-3 flex flex-wrap gap-1">
                            {(source.tags ?? []).map((tag) => (
                              <span key={tag} className="rounded bg-neutral-100 px-2 py-1 text-[11px] text-neutral-600">
                                {tag}
                              </span>
                            ))}
                          </div>

                          {source.addedAt && (
                            <div className="mt-3 flex items-center gap-1.5 text-xs text-neutral-500">
                              <CalendarClock className="size-3.5" />
                              Added {formatAddedAt(source.addedAt)}
                            </div>
                          )}

                          <details className="mt-3">
                            <summary className="cursor-pointer text-xs font-medium text-neutral-500">Edit workspace metadata</summary>
                            <div className="mt-3 grid gap-3 sm:grid-cols-2">
                              <div>
                                <Label className="text-xs">Category</Label>
                                <input
                                  value={source.category ?? ""}
                                  onChange={(event) => updateSource(source.key, { category: event.target.value })}
                                  className="mt-1 h-9 w-full rounded-md border border-neutral-200 px-3 text-sm"
                                />
                              </div>
                              <div>
                                <Label className="text-xs">Tags, comma-separated</Label>
                                <input
                                  value={(source.tags ?? []).join(", ")}
                                  onChange={(event) =>
                                    updateSource(source.key, {
                                      tags: event.target.value.split(",").map((tag) => tag.trim()).filter(Boolean),
                                    })
                                  }
                                  className="mt-1 h-9 w-full rounded-md border border-neutral-200 px-3 text-sm"
                                />
                              </div>
                              <div className="sm:col-span-2">
                                <Label className="text-xs">Purpose / project note</Label>
                                <input
                                  value={source.purpose ?? ""}
                                  onChange={(event) => updateSource(source.key, { purpose: event.target.value })}
                                  className="mt-1 h-9 w-full rounded-md border border-neutral-200 px-3 text-sm"
                                />
                              </div>
                            </div>
                          </details>
                        </div>
                        <button
                          type="button"
                          className="self-start rounded-md p-2 text-neutral-400 transition hover:bg-red-50 hover:text-red-600"
                          onClick={() =>
                            setWorkspace((current) => ({
                              ...current,
                              sources: current.sources.filter((item) => item.key !== source.key),
                            }))
                          }
                          aria-label={`Delete ${sourceTitle(source)}`}
                        >
                          <Trash2 className="size-4" />
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </CardContent>
          </Card>

          <div className="space-y-6">
            <Card>
              <CardHeader>
                <CardTitle>Workspace settings</CardTitle>
                <CardDescription>These settings are stored in the workspace file.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-5">
                <div className="space-y-2">
                  <Label htmlFor="namespace">Word bibliography namespace</Label>
                  <select
                    id="namespace"
                    className="h-10 w-full rounded-md border border-neutral-200 bg-white px-3 text-sm shadow-sm outline-none focus:ring-2 focus:ring-neutral-400"
                    value={workspace.settings.namespace}
                    onChange={(event) => updateSettings({ namespace: event.target.value as NamespaceMode })}
                  >
                    <option value="office2006">Office 2006+ (recommended)</option>
                    <option value="legacy2004">Legacy Word 2004</option>
                  </select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="lcid">Fallback bibliography language</Label>
                  <select
                    id="lcid"
                    className="h-10 w-full rounded-md border border-neutral-200 bg-white px-3 text-sm shadow-sm outline-none focus:ring-2 focus:ring-neutral-400"
                    value={workspace.settings.defaultLcid}
                    onChange={(event) => updateSettings({ defaultLcid: Number(event.target.value) })}
                  >
                    <option value={1033}>English (US) — 1033</option>
                    <option value={2057}>English (UK) — 2057</option>
                    <option value={1030}>Danish — 1030</option>
                    <option value={1031}>German — 1031</option>
                    <option value={1053}>Swedish — 1053</option>
                    <option value={1044}>Norwegian — 1044</option>
                  </select>
                </div>

                <label className="flex cursor-pointer items-start gap-3 rounded-lg border border-neutral-200 p-3">
                  <input
                    className="mt-1 size-4"
                    type="checkbox"
                    checked={workspace.settings.trackAddedDates}
                    onChange={(event) => updateSettings({ trackAddedDates: event.target.checked })}
                  />
                  <span>
                    <span className="block text-sm font-medium">Track date added</span>
                    <span className="mt-1 block text-xs leading-5 text-neutral-500">
                      New sources receive an <code>addedAt</code> timestamp saved in the workspace file. Existing sources are not back-dated.
                    </span>
                  </span>
                </label>

                <label className="flex cursor-pointer items-start gap-3 rounded-lg border border-neutral-200 p-3">
                  <input
                    className="mt-1 size-4"
                    type="checkbox"
                    checked={workspace.settings.deduplicate}
                    onChange={(event) => updateSettings({ deduplicate: event.target.checked })}
                  />
                  <span>
                    <span className="block text-sm font-medium">Deduplicate Word export</span>
                    <span className="mt-1 block text-xs leading-5 text-neutral-500">
                      Uses DOI, URL, or title/year during conversion.
                    </span>
                  </span>
                </label>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle>Portable database</CardTitle>
                <CardDescription>No server database is required. Save the entire state to a file.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-2">
                <Button className="w-full justify-start" variant="outline" onClick={exportWorkspace}>
                  <FileDown className="size-4" />
                  Save workspace configuration
                </Button>
                <Button className="w-full justify-start" variant="outline" onClick={() => workspaceFileRef.current?.click()}>
                  <FolderOpen className="size-4" />
                  Open workspace configuration
                </Button>
                <Button className="w-full justify-start" variant="outline" onClick={exportCsl}>
                  <FileJson className="size-4" />
                  Export raw CSL-JSON
                </Button>
              </CardContent>
            </Card>

            {result && (
              <Card className="border-emerald-200">
                <CardHeader>
                  <CardTitle className="text-emerald-800">Word XML ready</CardTitle>
                  <CardDescription>{result.itemCount} sources converted.</CardDescription>
                </CardHeader>
                <CardContent>
                  <div className="flex flex-wrap gap-2">
                    <Button onClick={() => downloadText("Sources.xml", result.xml, "application/xml;charset=utf-8")}>
                      <Download className="size-4" />
                      Download
                    </Button>
                    <Button variant="outline" onClick={() => void copyXml()}>
                      <Clipboard className="size-4" />
                      {copied ? "Copied" : "Copy XML"}
                    </Button>
                  </div>
                  <details className="mt-4">
                    <summary className="cursor-pointer text-sm font-medium text-neutral-600">Preview XML</summary>
                    <pre className="mt-3 max-h-72 overflow-auto rounded-lg bg-neutral-950 p-4 text-[10px] leading-5 text-neutral-100">
                      {result.xml}
                    </pre>
                  </details>
                </CardContent>
              </Card>
            )}
          </div>
        </div>

        <Card className="mt-6">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <Plus className="size-5" />
              Add sources
            </CardTitle>
            <CardDescription>
              Paste CSL-JSON or upload a CSL/workspace JSON file. Raw CSL sources are appended to the current library.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div
              className="mb-4 flex cursor-pointer items-center justify-center gap-3 rounded-lg border border-dashed border-neutral-300 bg-neutral-50 px-4 py-4 text-sm text-neutral-600 transition hover:bg-neutral-100"
              onClick={() => cslFileRef.current?.click()}
              onDragOver={(event) => event.preventDefault()}
              onDrop={(event) => {
                event.preventDefault()
                void readFile(event.dataTransfer.files[0])
              }}
              role="button"
              tabIndex={0}
            >
              <Upload className="size-4" />
              Drop CSL-JSON or a workspace file here, or click to browse
              <input
                ref={cslFileRef}
                className="hidden"
                type="file"
                accept=".json,application/json"
                onChange={(event) => void readFile(event.target.files?.[0])}
              />
            </div>
            <Textarea
              value={importText}
              onChange={(event) => setImportText(event.target.value)}
              className="min-h-48 resize-y font-mono text-xs leading-5"
              placeholder='[{"type":"article-journal","title":"…"}]'
              spellCheck={false}
            />
            <div className="mt-4 flex flex-wrap gap-2">
              <Button
                onClick={() => {
                  try {
                    if (!importText.trim()) throw new Error("Paste CSL-JSON first.")
                    importPayload(importText)
                  } catch (reason) {
                    setError(reason instanceof Error ? reason.message : "Could not import CSL-JSON.")
                  }
                }}
              >
                <Plus className="size-4" />
                Add to library
              </Button>
              <Button variant="outline" onClick={() => setImportText("")}>Clear</Button>
              <Button variant="outline" onClick={() => buildWordXml(false)}>Preview selected Word XML</Button>
            </div>
          </CardContent>
        </Card>
      </div>
    </main>
  )
}
