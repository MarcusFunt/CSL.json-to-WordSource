import { useMemo, useRef, useState } from "react"
import {
  BookOpen,
  CheckCircle2,
  Clipboard,
  Download,
  FileJson,
  FileOutput,
  Github,
  Info,
  ShieldCheck,
  Upload,
  WandSparkles,
  XCircle,
} from "lucide-react"
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import {
  convertCslJson,
  exampleCslJson,
  parseCslJson,
  type ConversionResult,
  type NamespaceMode,
} from "@/lib/cslToWord"

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

export default function App() {
  const [input, setInput] = useState(exampleCslJson)
  const [namespace, setNamespace] = useState<NamespaceMode>("office2006")
  const [defaultLcid, setDefaultLcid] = useState(1033)
  const [deduplicate, setDeduplicate] = useState(true)
  const [result, setResult] = useState<ConversionResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [copied, setCopied] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const parsedCount = useMemo(() => {
    try {
      return parseCslJson(input).length
    } catch {
      return null
    }
  }, [input])

  const convert = () => {
    try {
      const next = convertCslJson(input, { namespace, defaultLcid, deduplicate })
      setResult(next)
      setError(null)
    } catch (reason) {
      setResult(null)
      setError(reason instanceof Error ? reason.message : "Could not convert the input.")
    }
  }

  const readFile = async (file?: File) => {
    if (!file) return
    try {
      const text = await file.text()
      setInput(text)
      setResult(null)
      setError(null)
    } catch {
      setError("Could not read that file.")
    }
  }

  const copyXml = async () => {
    if (!result) return
    await navigator.clipboard.writeText(result.xml)
    setCopied(true)
    window.setTimeout(() => setCopied(false), 1500)
  }

  return (
    <main className="min-h-screen bg-neutral-50">
      <div className="mx-auto max-w-7xl px-4 py-10 sm:px-6 lg:px-8">
        <header className="mb-8 flex flex-col gap-5 border-b border-neutral-200 pb-8 md:flex-row md:items-end md:justify-between">
          <div className="max-w-3xl">
            <div className="mb-3 flex flex-wrap items-center gap-2">
              <Badge>Client-side</Badge>
              <Badge>CSL-JSON</Badge>
              <Badge>Word Sources.xml</Badge>
            </div>
            <h1 className="text-3xl font-bold tracking-tight text-neutral-950 sm:text-4xl">
              CSL.json → Word Sources.xml
            </h1>
            <p className="mt-3 max-w-2xl text-base leading-7 text-neutral-600">
              Convert structured CSL references into Microsoft Word bibliography sources.
              Everything runs locally in your browser.
            </p>
          </div>
          <Button variant="outline" asChild>
            <a
              href="https://github.com/MarcusFunt/CSL.json-to-WordSource"
              target="_blank"
              rel="noreferrer"
            >
              <Github className="size-4" />
              GitHub
            </a>
          </Button>
        </header>

        <div className="grid gap-6 lg:grid-cols-[minmax(0,1.2fr)_minmax(340px,0.8fr)]">
          <Card>
            <CardHeader>
              <div className="flex items-center justify-between gap-4">
                <div>
                  <CardTitle className="flex items-center gap-2">
                    <FileJson className="size-5" />
                    CSL-JSON input
                  </CardTitle>
                  <CardDescription className="mt-2">
                    Paste one item, an array, or an object containing an <code>items</code> array.
                  </CardDescription>
                </div>
                {parsedCount !== null && <Badge>{parsedCount} source{parsedCount === 1 ? "" : "s"}</Badge>}
              </div>
            </CardHeader>
            <CardContent>
              <div
                className="mb-4 flex cursor-pointer items-center justify-center gap-3 rounded-lg border border-dashed border-neutral-300 bg-neutral-50 px-4 py-4 text-sm text-neutral-600 transition hover:bg-neutral-100"
                onClick={() => fileInputRef.current?.click()}
                onDragOver={(event) => event.preventDefault()}
                onDrop={(event) => {
                  event.preventDefault()
                  void readFile(event.dataTransfer.files[0])
                }}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") fileInputRef.current?.click()
                }}
              >
                <Upload className="size-4" />
                Drop a .json file here or click to browse
                <input
                  ref={fileInputRef}
                  className="hidden"
                  type="file"
                  accept=".json,application/json"
                  onChange={(event) => void readFile(event.target.files?.[0])}
                />
              </div>

              <Textarea
                value={input}
                onChange={(event) => {
                  setInput(event.target.value)
                  setResult(null)
                  setError(null)
                }}
                className="min-h-[520px] resize-y font-mono text-xs leading-5"
                spellCheck={false}
                aria-label="CSL-JSON input"
              />

              <div className="mt-4 flex flex-wrap gap-2">
                <Button onClick={convert}>
                  <WandSparkles className="size-4" />
                  Convert
                </Button>
                <Button
                  variant="outline"
                  onClick={() => {
                    setInput(exampleCslJson)
                    setResult(null)
                    setError(null)
                  }}
                >
                  Load example
                </Button>
                <Button
                  variant="ghost"
                  onClick={() => {
                    setInput("")
                    setResult(null)
                    setError(null)
                  }}
                >
                  Clear
                </Button>
              </div>
            </CardContent>
          </Card>

          <div className="flex flex-col gap-6">
            <Card>
              <CardHeader>
                <CardTitle>Conversion settings</CardTitle>
                <CardDescription>
                  Defaults are appropriate for current desktop versions of Word.
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-5">
                <div className="space-y-2">
                  <Label htmlFor="namespace">Word bibliography namespace</Label>
                  <select
                    id="namespace"
                    className="h-10 w-full rounded-md border border-neutral-200 bg-white px-3 text-sm shadow-sm outline-none focus:ring-2 focus:ring-neutral-400"
                    value={namespace}
                    onChange={(event) => setNamespace(event.target.value as NamespaceMode)}
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
                    value={defaultLcid}
                    onChange={(event) => setDefaultLcid(Number(event.target.value))}
                  >
                    <option value={1033}>English (United States) — 1033</option>
                    <option value={2057}>English (United Kingdom) — 2057</option>
                    <option value={1030}>Danish — 1030</option>
                    <option value={1031}>German — 1031</option>
                    <option value={1053}>Swedish — 1053</option>
                    <option value={1044}>Norwegian — 1044</option>
                  </select>
                  <p className="text-xs leading-5 text-neutral-500">
                    A CSL <code>language</code> field overrides this fallback when recognized.
                  </p>
                </div>

                <label className="flex cursor-pointer items-start gap-3 rounded-lg border border-neutral-200 p-3">
                  <input
                    className="mt-1 size-4"
                    type="checkbox"
                    checked={deduplicate}
                    onChange={(event) => setDeduplicate(event.target.checked)}
                  />
                  <span>
                    <span className="block text-sm font-medium">Remove obvious duplicates</span>
                    <span className="mt-1 block text-xs leading-5 text-neutral-500">
                      Matches DOI first, then URL, then normalized title + year.
                    </span>
                  </span>
                </label>
              </CardContent>
            </Card>

            {error && (
              <Card className="border-red-200 bg-red-50">
                <CardContent className="flex gap-3 pt-6 text-sm text-red-800">
                  <XCircle className="mt-0.5 size-5 shrink-0" />
                  <div>
                    <p className="font-semibold">Conversion failed</p>
                    <p className="mt-1">{error}</p>
                  </div>
                </CardContent>
              </Card>
            )}

            {result ? (
              <Card className="border-emerald-200">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-emerald-800">
                    <CheckCircle2 className="size-5" />
                    Sources.xml ready
                  </CardTitle>
                  <CardDescription>
                    Converted {result.itemCount} source{result.itemCount === 1 ? "" : "s"}.
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <div className="flex flex-wrap gap-2">
                    <Button
                      onClick={() => downloadText("Sources.xml", result.xml, "application/xml;charset=utf-8")}
                    >
                      <Download className="size-4" />
                      Download Sources.xml
                    </Button>
                    <Button variant="outline" onClick={() => void copyXml()}>
                      <Clipboard className="size-4" />
                      {copied ? "Copied" : "Copy XML"}
                    </Button>
                  </div>

                  {result.warnings.length > 0 && (
                    <div className="mt-5 rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">
                      <p className="font-semibold">Notes</p>
                      <ul className="mt-2 list-disc space-y-1 pl-5">
                        {result.warnings.map((warning) => <li key={warning}>{warning}</li>)}
                      </ul>
                    </div>
                  )}

                  <details className="mt-5">
                    <summary className="cursor-pointer text-sm font-medium text-neutral-700">
                      Preview generated XML
                    </summary>
                    <pre className="mt-3 max-h-80 overflow-auto rounded-lg bg-neutral-950 p-4 text-[11px] leading-5 text-neutral-100">
                      {result.xml}
                    </pre>
                  </details>
                </CardContent>
              </Card>
            ) : (
              <Card>
                <CardContent className="flex items-start gap-3 pt-6 text-sm text-neutral-600">
                  <FileOutput className="mt-0.5 size-5 shrink-0" />
                  <p>Convert valid CSL-JSON to preview and download the resulting Word source library.</p>
                </CardContent>
              </Card>
            )}
          </div>
        </div>

        <section className="mt-8 grid gap-4 md:grid-cols-3">
          <div className="rounded-xl border border-neutral-200 bg-white p-5">
            <ShieldCheck className="size-5" />
            <h2 className="mt-3 font-semibold">Private by design</h2>
            <p className="mt-2 text-sm leading-6 text-neutral-600">
              No backend and no upload API. Conversion happens entirely inside the browser.
            </p>
          </div>
          <div className="rounded-xl border border-neutral-200 bg-white p-5">
            <BookOpen className="size-5" />
            <h2 className="mt-3 font-semibold">Structured names</h2>
            <p className="mt-2 text-sm leading-6 text-neutral-600">
              Personal names and single corporate authors are kept structured instead of reparsed from citation text.
            </p>
          </div>
          <div className="rounded-xl border border-neutral-200 bg-white p-5">
            <Info className="size-5" />
            <h2 className="mt-3 font-semibold">Word import</h2>
            <p className="mt-2 text-sm leading-6 text-neutral-600">
              In Word, open References → Manage Sources → Browse and select the downloaded Sources.xml.
            </p>
          </div>
        </section>
      </div>
    </main>
  )
}
