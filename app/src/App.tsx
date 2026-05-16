import { useMemo, useRef, useState } from 'react'
import { Document, HeadingLevel, Packer, PageBreak, Paragraph, ShadingType, TextRun, AlignmentType } from 'docx'

type Uploadable = File & { webkitRelativePath?: string }

interface SourceFile {
  id: string
  filePath: string
  fileName: string
  extension: string
  projectName: string
  content: string
  included: boolean
  order: number
}

const DEFAULT_EXTENSIONS = [
  '.cs', '.aspx', '.master', '.cshtml', '.razor',
  '.css', '.js', '.ts', '.json', '.xml', '.config',
  '.html', '.xaml', '.resx', '.sln', '.csproj',
]

// Any segment of the path matching these words is a build artifact — always hidden
const EXCLUDED_SEGMENT_RE =
  /^(bin|obj|debug|release|\.vs|\.git|packages|node_modules|dist|build|out|publish|testresults|coverage|artifacts|_build|_output)$/i

const EXCLUDED_FILE_RE = [
  /\.(dll|exe|pdb|cache|nupkg|zip|7z|rar|png|jpg|jpeg|gif|webp|bmp|ico|pdf|mp4|mov|avi)$/i,
  /\.designer\.cs$/i,
  /\.g\.cs$/i,
  /\.g\.i\.cs$/i,
  /\.min\.(js|css)$/i,
]

// Famous 3rd-party library filenames — never include regardless of extension settings
const EXCLUDED_THIRD_PARTY_RE = [
  /^jquery[.\-]/i,
  /^bootstrap[.\-]/i,
  /^normalize\./i,
  /^modernizr[.\-]/i,
  /^angular[.\-]/i,
  /^vue[.\-]/i,
  /^react[.\-0-9]/i,
  /^react-dom[.\-]/i,
  /^lodash[.\-]/i,
  /^underscore[.\-]/i,
  /^moment[.\-]/i,
  /^popper[.\-]/i,
  /^axios[.\-]/i,
  /^signalr[.\-]/i,
  /^knockout[.\-]/i,
  /^d3[.\-v]/i,
  /^chart[.\-]/i,
  /^select2[.\-]/i,
  /^datatables[.\-]/i,
  /^fontawesome[.\-]/i,
  /^font-awesome[.\-]/i,
  /^respond[.\-]/i,
  /^html5shiv[.\-]/i,
  /^toastr[.\-]/i,
  /^sweetalert[.\-0-9]/i,
  /^animate\.(css|min\.css)$/i,
  /^tippy[.\-]/i,
  /^flatpickr[.\-]/i,
  /^cldr[.\-]/i,
  /^globalize[.\-]/i,
]

const LS_KEY = 'code-extractor-ext'

interface RawFile {
  filePath: string
  fileName: string
  extension: string
  content: string
}

function loadExtensions(): Set<string> {
  try {
    const saved = localStorage.getItem(LS_KEY)
    if (saved) {
      const arr = JSON.parse(saved) as unknown
      if (Array.isArray(arr) && arr.length > 0) {
        const valid = (arr as unknown[]).filter(
          (e): e is string => typeof e === 'string' && /^\.[a-z0-9]+$/.test(e),
        )
        if (valid.length > 0) return new Set(valid)
      }
    }
  } catch { /* ignore */ }
  return new Set(DEFAULT_EXTENSIONS)
}

function shouldExcludeByName(fileName: string): boolean {
  return EXCLUDED_THIRD_PARTY_RE.some((re) => re.test(fileName))
}

// ─── VS Code Dark+ token colours ───────────────────────────────────────────
// Palette: keyword, string, comment, number, type/class, operator, default
const CLR_BG      = 'F8F8F8'  // near-white like VS Code light theme
const CLR_DEFAULT = '1E1E1E'  // almost-black
const CLR_COMMENT = '008000'  // green
const CLR_STRING  = 'A31515'  // dark red
const CLR_NUMBER  = '098658'  // teal-green
const CLR_KEYWORD: Record<string, string> = {
  // C# / generic
  default: '0000FF',  // blue keywords
}
const CLR_TYPE    = '267F99'  // teal type names
const CLR_PREPROC = '808080'  // grey preprocessor
const CLR_XML_TAG = '800000'  // dark red tags
const CLR_XML_ATTR = 'FF0000' // red attr name
const CLR_XML_VAL  = '0000FF' // blue attr values

// C# keywords
const CS_KEYWORDS = new Set([
  'abstract','as','base','bool','break','byte','case','catch','char','checked',
  'class','const','continue','decimal','default','delegate','do','double','else',
  'enum','event','explicit','extern','false','finally','fixed','float','for',
  'foreach','goto','if','implicit','in','int','interface','internal','is','lock',
  'long','namespace','new','null','object','operator','out','override','params',
  'private','protected','public','readonly','ref','return','sbyte','sealed',
  'short','sizeof','stackalloc','static','string','struct','switch','this',
  'throw','true','try','typeof','uint','ulong','unchecked','unsafe','ushort',
  'using','virtual','void','volatile','while','async','await','var','dynamic',
  'get','set','value','yield','partial','record','init','required','with','and',
  'or','not','when','global','file','scoped',
])
// JS/TS keywords
const JS_KEYWORDS = new Set([
  'break','case','catch','class','const','continue','debugger','default','delete',
  'do','else','export','extends','false','finally','for','function','if','import',
  'in','instanceof','let','new','null','return','static','super','switch','this',
  'throw','true','try','typeof','undefined','var','void','while','with','yield',
  'async','await','of','from','as','type','interface','enum','implements','declare',
  'abstract','readonly','override','namespace','module','any','never','unknown',
  'boolean','number','string','object','symbol','bigint','public','private','protected',
])

type Token = { text: string; color: string; bold?: boolean }

function tokenizeCS(line: string): Token[] {
  const tokens: Token[] = []
  // Single-line comment
  const commentIdx = line.indexOf('//')
  let code = line
  let comment = ''
  if (commentIdx >= 0) {
    code = line.slice(0, commentIdx)
    comment = line.slice(commentIdx)
  }
  // Preprocessor
  if (code.trimStart().startsWith('#')) {
    tokens.push({ text: code, color: CLR_PREPROC })
    if (comment) tokens.push({ text: comment, color: CLR_COMMENT })
    return tokens
  }
  // Tokenize code part
  const re = /(@?"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|@?\$"(?:[^"\\]|\\.)*"|[a-zA-Z_][a-zA-Z0-9_]*|\d+\.?\d*[fFdDmMlLuU]*|[{}()\[\];:,.<>+\-*/%=!&|^~?]|\s+)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(code)) !== null) {
    const t = m[0]
    if (t.startsWith('"') || t.startsWith("'") || t.startsWith('@"') || t.startsWith('$"')) {
      tokens.push({ text: t, color: CLR_STRING })
    } else if (/^\d/.test(t)) {
      tokens.push({ text: t, color: CLR_NUMBER })
    } else if (CS_KEYWORDS.has(t)) {
      tokens.push({ text: t, color: CLR_KEYWORD.default, bold: false })
    } else if (/^[A-Z][a-zA-Z0-9_]*$/.test(t)) {
      tokens.push({ text: t, color: CLR_TYPE })
    } else {
      tokens.push({ text: t, color: CLR_DEFAULT })
    }
  }
  if (comment) tokens.push({ text: comment, color: CLR_COMMENT })
  return tokens
}

function tokenizeJS(line: string): Token[] {
  const tokens: Token[] = []
  const commentIdx = line.indexOf('//')
  let code = line
  let comment = ''
  if (commentIdx >= 0) {
    code = line.slice(0, commentIdx)
    comment = line.slice(commentIdx)
  }
  const re = /(`[^`]*`|"|'|[a-zA-Z_$][a-zA-Z0-9_$]*|\d+\.?\d*|[{}()\[\];:,.<>+\-*/%=!&|^~?]|\s+)/g
  let inStr = ''
  let strBuf = ''
  let m: RegExpExecArray | null
  // Simplified: handle template literals and quoted strings as atomic tokens
  const atomic = /(`[^`]*`|"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|[a-zA-Z_$][a-zA-Z0-9_$]*|\d+\.?\d*|.)/g
  void re; void inStr; void strBuf
  while ((m = atomic.exec(code)) !== null) {
    const t = m[0]
    if (t.startsWith('"') || t.startsWith("'") || t.startsWith('`')) {
      tokens.push({ text: t, color: CLR_STRING })
    } else if (/^\d/.test(t)) {
      tokens.push({ text: t, color: CLR_NUMBER })
    } else if (JS_KEYWORDS.has(t)) {
      tokens.push({ text: t, color: CLR_KEYWORD.default })
    } else if (/^[A-Z][a-zA-Z0-9_]*$/.test(t)) {
      tokens.push({ text: t, color: CLR_TYPE })
    } else {
      tokens.push({ text: t, color: CLR_DEFAULT })
    }
  }
  if (comment) tokens.push({ text: comment, color: CLR_COMMENT })
  return tokens
}

function tokenizeXML(line: string): Token[] {
  // Very simple XML/HTML coloring
  const tokens: Token[] = []
  if (/^\s*<!--/.test(line)) {
    tokens.push({ text: line, color: CLR_COMMENT })
    return tokens
  }
  // Regex built from parts to avoid quote-escaping issues in TS
  const reStr = [
    String.raw`<\/?[A-Za-z][A-Za-z0-9.:_\-]*`,
    String.raw`>`,
    String.raw`\/>`,
    String.raw`"[^"]*"`,
    String.raw`'[^']*'`,
    String.raw`[A-Za-z][A-Za-z0-9.:_\-]*(?=\s*=)`,
    String.raw`[^<>"'=\s]+`,
    String.raw`\s+`,
    String.raw`.`,
  ].join('|')
  const re = new RegExp(`(${reStr})`, 'g')
  let m: RegExpExecArray | null
  let inTag = false
  while ((m = re.exec(line)) !== null) {
    const t = m[0]
    if (t.startsWith('</') || t.startsWith('<')) {
      inTag = true
      tokens.push({ text: t, color: CLR_XML_TAG })
    } else if (t === '>' || t === '/>') {
      inTag = false
      tokens.push({ text: t, color: CLR_XML_TAG })
    } else if ((t.startsWith('"') || t.startsWith("'")) && inTag) {
      tokens.push({ text: t, color: CLR_XML_VAL })
    } else if (inTag && /^[A-Za-z]/.test(t)) {
      tokens.push({ text: t, color: CLR_XML_ATTR })
    } else {
      tokens.push({ text: t, color: CLR_DEFAULT })
    }
  }
  return tokens
}

function tokenizePlain(line: string, _ext: string): Token[] {
  return [{ text: line, color: CLR_DEFAULT }]
}

// ─── CSS tokenizer ────────────────────────────────────────────────────────
const CLR_CSS_AT     = 'AF00DB'  // purple  — @media, @keyframes
const CLR_CSS_SEL    = '800000'  // dark-red — selectors
const CLR_CSS_PROP   = 'FF0000'  // red      — property names
const CLR_CSS_VAL    = '0070C1'  // blue     — keyword values

function tokenizeCSSValue(value: string): Token[] {
  const tokens: Token[] = []
  const re = /('(?:[^'\\]|\\.)*'|"(?:[^"\\]|\\.)*"|#[0-9a-fA-F]{3,8}|\d+\.?\d*(?:%|px|em|rem|vh|vw|vmin|vmax|pt|pc|cm|mm|in|ex|ch|fr|deg|rad|turn|s|ms)?|[a-zA-Z][a-zA-Z0-9-]*|.)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(value)) !== null) {
    const t = m[0]
    if (t.startsWith('"') || t.startsWith("'")) {
      tokens.push({ text: t, color: CLR_STRING })
    } else if (/^#[0-9a-fA-F]/.test(t) || /^\d/.test(t)) {
      tokens.push({ text: t, color: CLR_NUMBER })
    } else if (/^[a-zA-Z]/.test(t)) {
      tokens.push({ text: t, color: CLR_CSS_VAL })
    } else {
      tokens.push({ text: t, color: CLR_DEFAULT })
    }
  }
  return tokens
}

function tokenizeCSS(line: string): Token[] {
  // Comments
  if (/^\/\*|^\*/.test(line)) return [{ text: line, color: CLR_COMMENT }]
  // @-rules
  const atMatch = line.match(/^(@[a-zA-Z-]+)(.*)/)
  if (atMatch) {
    return [
      { text: atMatch[1], color: CLR_CSS_AT },
      { text: atMatch[2], color: CLR_DEFAULT },
    ]
  }
  // Property: value;  (line like "color: red;")
  const propMatch = line.match(/^([a-zA-Z-]+)(\s*:\s*)(.+)$/)
  if (propMatch) {
    return [
      { text: propMatch[1], color: CLR_CSS_PROP },
      { text: propMatch[2], color: CLR_DEFAULT },
      ...tokenizeCSSValue(propMatch[3]),
    ]
  }
  // Selector or brace lines
  if (line.includes('{') || line === '}' || line.includes(',')) {
    return [{ text: line, color: CLR_CSS_SEL }]
  }
  return [{ text: line, color: CLR_DEFAULT }]
}

function tokenizeLine(line: string, ext: string): Token[] {
  switch (ext) {
    case '.cs': case '.razor': case '.cshtml': return tokenizeCS(line)
    case '.js': case '.ts': case '.tsx': case '.jsx': return tokenizeJS(line)
    case '.css': return tokenizeCSS(line)
    case '.xml': case '.config': case '.xaml': case '.resx':
    case '.html': case '.aspx': case '.master': case '.cshtml': return tokenizeXML(line)
    default: return tokenizePlain(line, ext)
  }
}

function normalizePath(path: string): string {
  return path.replace(/\\/g, '/')
}

function extensionOf(path: string): string {
  const dot = path.lastIndexOf('.')
  return dot >= 0 ? path.slice(dot).toLowerCase() : ''
}

function shouldExclude(path: string): boolean {
  const segments = path.split('/')
  // Exclude if any path segment is a known build/artifact folder
  if (segments.some((seg) => EXCLUDED_SEGMENT_RE.test(seg))) return true
  // Exclude by file extension / pattern
  if (EXCLUDED_FILE_RE.some((pattern) => pattern.test(path))) return true
  // Exclude well-known 3rd-party library filenames
  const fileName = segments[segments.length - 1] ?? ''
  return shouldExcludeByName(fileName)
}

function parentDir(path: string): string {
  const slash = path.lastIndexOf('/')
  return slash >= 0 ? path.slice(0, slash) : ''
}

function detectProjectRoots(paths: string[]): string[] {
  const csprojRoots = paths
    .filter((path) => path.toLowerCase().endsWith('.csproj'))
    .map((path) => parentDir(path))
    .filter(Boolean)
    .sort((a, b) => b.length - a.length)

  if (csprojRoots.length > 0) return csprojRoots

  return Array.from(new Set(paths.map((path) => path.split('/')[0] ?? '').filter(Boolean)))
    .sort((a, b) => b.length - a.length)
}

function projectFromPath(path: string, roots: string[]): string {
  const matched = roots.find((root) => path.startsWith(root + '/') || path === root)
  if (!matched) return path.split('/')[0] ?? 'Project'
  const segments = matched.split('/').filter(Boolean)
  return segments[segments.length - 1] ?? 'Project'
}

function expandTabs(line: string): string {
  return line.replace(/\t/g, '    ')
}

function nbspLeading(expanded: string): string {
  const leading = expanded.match(/^\s+/)?.[0] ?? ''
  return leading.replace(/ /g, '\u00A0')
}

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const dragItem = useRef<{ projectName: string; id: string } | null>(null)
  const dragOverItem = useRef<{ projectName: string; id: string } | null>(null)
  const dragProject = useRef<string | null>(null)
  const dragOverProject = useRef<string | null>(null)
  const [rawUploads, setRawUploads] = useState<RawFile[]>([])
  const [allFiles, setAllFiles] = useState<SourceFile[]>([])
  const [projectTitle, setProjectTitle] = useState('Final Project Book - Code Appendix')
  const [isBusy, setIsBusy] = useState(false)
  const [status, setStatus] = useState('Upload one or more project folders to begin.')
  const [supportedExtensions, setSupportedExtensions] = useState<Set<string>>(loadExtensions)
  const [projectOrder, setProjectOrder] = useState<string[]>([])
  // Settings modal state
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [pendingExts, setPendingExts] = useState<Set<string>>(new Set())
  const [pendingInput, setPendingInput] = useState('')
  const [pendingError, setPendingError] = useState('')

  const grouped = useMemo(() => {
    const sorted = [...allFiles].sort((a, b) => {
      const pa = projectOrder.indexOf(a.projectName)
      const pb = projectOrder.indexOf(b.projectName)
      if (pa !== pb) return pa - pb
      return a.order - b.order
    })
    const byProject = new Map<string, SourceFile[]>()
    for (const file of sorted) {
      const bucket = byProject.get(file.projectName) ?? []
      bucket.push(file)
      byProject.set(file.projectName, bucket)
    }
    return byProject
  }, [allFiles, projectOrder])

  const includedCount = useMemo(() => allFiles.filter((f) => f.included).length, [allFiles])

  // Rebuilds the visible file list from raw uploads + current extensions.
  // Preserves user checkbox state from existingFiles.
  function rebuildFromRaw(raw: RawFile[], exts: Set<string>, existingFiles: SourceFile[]) {
    const filtered = raw.filter((f) => exts.has(f.extension))
    const existingIncluded = new Map(existingFiles.map((f) => [f.filePath, f.included]))

    if (filtered.length === 0) {
      setAllFiles([])
      if (raw.length > 0) {
        setStatus('No files match the current extensions. Open ⚙ Settings to add more.')
      }
      return
    }

    const roots = detectProjectRoots(filtered.map((f) => f.filePath))
    const projectMap = new Map<string, RawFile[]>()
    for (const file of filtered) {
      const pname = projectFromPath(file.filePath, roots)
      const bucket = projectMap.get(pname) ?? []
      bucket.push(file)
      projectMap.set(pname, bucket)
    }

    const normalized: SourceFile[] = []
    const projectNames = Array.from(projectMap.keys()).sort((a, b) => a.localeCompare(b))
    for (const pname of projectNames) {
      const inProject = (projectMap.get(pname) ?? []).sort((a, b) => a.filePath.localeCompare(b.filePath))
      inProject.forEach((file, index) => {
        normalized.push({
          id: `${pname}:${file.filePath}`,
          filePath: file.filePath,
          fileName: file.fileName,
          extension: file.extension,
          projectName: pname,
          content: file.content,
          // Preserve prior checkbox state; new files default included
          included: existingIncluded.get(file.filePath) ?? true,
          order: index,
        })
      })
    }

    setAllFiles(normalized)
    // Preserve existing project order; append newly detected projects at the end
    setProjectOrder((prev) => {
      const keep = prev.filter((p) => projectNames.includes(p))
      const added = projectNames.filter((p) => !prev.includes(p))
      return [...keep, ...added]
    })
    const inc = normalized.filter((f) => f.included).length
    setStatus(`${inc} / ${normalized.length} file(s) across ${projectNames.length} project(s).`)
  }

  async function handleFiles(files: File[]) {
    const uploadables = files as Uploadable[]

    setIsBusy(true)
    try {
      // Load ALL files that pass path/binary exclusions — no extension filter here.
      // Extension filtering happens in rebuildFromRaw so Apply works without re-uploading.
      const candidates = uploadables
        .map((file) => ({
          file,
          relative: normalizePath(file.webkitRelativePath || file.name),
        }))
        .filter(({ relative }) => !shouldExclude(relative))

      if (candidates.length === 0) {
        setStatus('No source files found — only build outputs or binaries detected.')
        return
      }

      const loaded = await Promise.all(
        candidates.map(async ({ file, relative }) => ({
          filePath: relative,
          fileName: file.name,
          extension: extensionOf(relative),
          content: await file.text(),
        })),
      )

      const seen = new Set(rawUploads.map((f) => f.filePath))
      const merged = [...rawUploads, ...loaded.filter((f) => !seen.has(f.filePath))]
      setRawUploads(merged)
      rebuildFromRaw(merged, supportedExtensions, allFiles)
    } finally {
      setIsBusy(false)
    }
  }

  function toggleIncluded(id: string) {
    setAllFiles((prev) => prev.map((file) => (file.id === id ? { ...file, included: !file.included } : file)))
  }

  // ─── Settings modal ──────────────────────────────────────────────────────
  function openSettings() {
    setPendingExts(new Set(supportedExtensions))
    setPendingInput('')
    setPendingError('')
    setSettingsOpen(true)
  }

  function closeSettings() {
    setSettingsOpen(false)
  }

  function applySettings() {
    setSupportedExtensions(pendingExts)
    try { localStorage.setItem(LS_KEY, JSON.stringify([...pendingExts])) } catch { /* ignore */ }
    rebuildFromRaw(rawUploads, pendingExts, allFiles)
    setSettingsOpen(false)
  }

  function addPendingExt() {
    const raw = pendingInput.trim().toLowerCase()
    const ext = raw.startsWith('.') ? raw : `.${raw}`
    if (!/^\.[a-z0-9]+$/.test(ext)) {
      setPendingError('Must be a simple extension like .txt or .sql')
      return
    }
    if (pendingExts.has(ext)) {
      setPendingError(`${ext} is already in the list`)
      return
    }
    if (EXCLUDED_FILE_RE.some((re) => re.test(`x${ext}`))) {
      setPendingError(`${ext} is a blocked file type (binary/media/generated)`)
      return
    }
    setPendingExts((prev) => new Set([...prev, ext]))
    setPendingInput('')
    setPendingError('')
  }

  function removePendingExt(ext: string) {
    if (pendingExts.size <= 1) return
    setPendingExts((prev) => {
      const next = new Set(prev)
      next.delete(ext)
      return next
    })
  }

  function resetPendingToDefaults() {
    setPendingExts(new Set(DEFAULT_EXTENSIONS))
    setPendingInput('')
    setPendingError('')
  }

  function onDragStart(projectName: string, id: string) {
    dragItem.current = { projectName, id }
  }

  function onDragEnter(projectName: string, id: string) {
    dragOverItem.current = { projectName, id }
  }

  function onDragEnd() {
    const from = dragItem.current
    const to = dragOverItem.current
    dragItem.current = null
    dragOverItem.current = null
    if (!from || !to || from.id === to.id || from.projectName !== to.projectName) return
    setAllFiles((prev) => {
      const inProject = prev
        .filter((f) => f.projectName === from.projectName)
        .sort((a, b) => a.order - b.order)
      const fromIdx = inProject.findIndex((f) => f.id === from.id)
      const toIdx   = inProject.findIndex((f) => f.id === to.id)
      if (fromIdx < 0 || toIdx < 0) return prev
      const reordered = [...inProject]
      const [moved] = reordered.splice(fromIdx, 1)
      reordered.splice(toIdx, 0, moved)
      const updates = new Map(reordered.map((f, order) => [f.id, { ...f, order }]))
      return prev.map((f) => updates.get(f.id) ?? f)
    })
  }

  // ─── Project-level controls ──────────────────────────────────────────────
  function toggleProject(projectName: string) {
    const projectFiles = allFiles.filter((f) => f.projectName === projectName)
    const allIncluded = projectFiles.every((f) => f.included)
    setAllFiles((prev) =>
      prev.map((f) => (f.projectName === projectName ? { ...f, included: !allIncluded } : f)),
    )
  }

  function removeProject(projectName: string) {
    setAllFiles((prev) => prev.filter((f) => f.projectName !== projectName))
    setRawUploads((prev) => {
      const paths = prev.map((f) => f.filePath)
      const roots = detectProjectRoots(paths)
      return prev.filter((f) => projectFromPath(f.filePath, roots) !== projectName)
    })
    setProjectOrder((prev) => prev.filter((p) => p !== projectName))
  }

  function onProjectDragStart(name: string) {
    dragProject.current = name
  }

  function onProjectDragEnter(name: string) {
    dragOverProject.current = name
  }

  function onProjectDragEnd() {
    const from = dragProject.current
    const to = dragOverProject.current
    dragProject.current = null
    dragOverProject.current = null
    if (!from || !to || from === to) return
    setProjectOrder((prev) => {
      const arr = [...prev]
      const fromIdx = arr.indexOf(from)
      const toIdx   = arr.indexOf(to)
      if (fromIdx < 0 || toIdx < 0) return prev
      arr.splice(fromIdx, 1)
      arr.splice(toIdx, 0, from)
      return arr
    })
  }

  async function exportDocx() {
    const included = [...allFiles]
      .filter((file) => file.included)
      .sort((a, b) => {
        if (a.projectName !== b.projectName) return a.projectName.localeCompare(b.projectName)
        return a.order - b.order
      })

    if (included.length === 0) {
      setStatus('Select at least one file before exporting.')
      return
    }

    setIsBusy(true)
    try {
      const paragraphs: Paragraph[] = [
        new Paragraph({
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
          spacing: { before: 800, after: 300 },
          children: [new TextRun({ text: projectTitle, bold: true, size: 44 })],
        }),
      ]

      let currentProject = ''
      for (const file of included) {
        if (file.projectName !== currentProject) {
          currentProject = file.projectName
          paragraphs.push(new Paragraph({ children: [new PageBreak()] }))
          paragraphs.push(
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun({ text: file.projectName, bold: true })],
            }),
          )
        }

        paragraphs.push(new Paragraph({ children: [new PageBreak()] }))
        paragraphs.push(
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun({ text: file.fileName, bold: true })],
          }),
        )
        paragraphs.push(
          new Paragraph({
            children: [new TextRun({ text: file.filePath, color: '64748B', italics: true })],
          }),
        )

        const lines = file.content.replace(/\r\n/g, '\n').split('\n')

        for (const rawLine of lines) {
          const expanded = expandTabs(rawLine)
          const leading  = nbspLeading(expanded)
          const rest     = expanded.slice((expanded.match(/^\s+/)?.[0] ?? '').length)
          const tokens   = tokenizeLine(rest, file.extension)

          paragraphs.push(
            new Paragraph({
              spacing: { before: 0, after: 0 },
              shading: { type: ShadingType.SOLID, fill: CLR_BG, color: CLR_BG },
              children: [
                // Preserve indentation as non-breaking spaces in Consolas
                ...(leading ? [new TextRun({ text: leading, font: 'Consolas', size: 18 })] : []),
                ...tokens.map((tok) =>
                  new TextRun({
                    text: tok.text,
                    font: 'Consolas',
                    size: 18,
                    color: tok.color,
                    bold: tok.bold ?? false,
                  }),
                ),
              ],
            }),
          )
        }
      }

      const doc = new Document({
        sections: [{ children: paragraphs }],
      })

      const blob = await Packer.toBlob(doc)
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = `${projectTitle.replace(/[^a-zA-Z0-9-_ ]/g, '_')}-code.docx`
      link.click()
      URL.revokeObjectURL(url)

      setStatus(`Exported ${included.length} file(s) to Word successfully.`)
    } finally {
      setIsBusy(false)
    }
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <h1>Code Extractor for Final Project Book</h1>
        <p>
          Upload one or more C# / ASP.NET project folders, select and reorder files, then export a Word document with one
          file per page.
        </p>
      </header>

      <section className="panel">
        <label className="field">
          <span>Project Book Title</span>
          <input
            value={projectTitle}
            onChange={(event) => setProjectTitle(event.target.value)}
            placeholder="Final Project Book - Code Appendix"
          />
        </label>

        <div className="actions-row">
          <button type="button" onClick={() => fileInputRef.current?.click()} disabled={isBusy}>
            Upload Folder(s)
          </button>
          <button type="button" className="primary" onClick={() => void exportDocx()} disabled={isBusy || includedCount === 0}>
            Export .docx
          </button>
          <button type="button" className="btn-icon" onClick={openSettings} title="Extension settings" aria-label="Extension settings">
            ⚙
          </button>
          <input
            ref={fileInputRef}
            type="file"
            multiple
            // @ts-expect-error webkitdirectory is supported by Chromium-based browsers
            webkitdirectory=""
            className="hidden-input"
            onChange={(event) => {
              const files = Array.from(event.target.files ?? [])
              if (files.length > 0) {
                void handleFiles(files)
              }
              event.target.value = ''
            }}
          />
        </div>

        <p className="status">{status}</p>
        <p className="meta">Included: {includedCount} / {allFiles.length} files</p>
      </section>

      {allFiles.length > 0 && (
        <section className="panel">
          <h2>Files by Project</h2>
          {Array.from(grouped.entries()).map(([projectName, files]) => {
            const allIncluded = files.every((f) => f.included)
            const someIncluded = files.some((f) => f.included)
            return (
            <div key={projectName} className="project-block">
              <div
                className="project-header"
                draggable
                onDragStart={() => onProjectDragStart(projectName)}
                onDragEnter={() => onProjectDragEnter(projectName)}
                onDragEnd={onProjectDragEnd}
                onDragOver={(e) => e.preventDefault()}
              >
                <span className="drag-handle" title="Drag to reorder project">⠿</span>
                <input
                  type="checkbox"
                  checked={allIncluded}
                  ref={(el) => { if (el) el.indeterminate = someIncluded && !allIncluded }}
                  onChange={() => toggleProject(projectName)}
                  title={allIncluded ? 'Exclude all files' : 'Include all files'}
                />
                <span className="project-name">{projectName}</span>
                <span className="project-count">{files.filter((f) => f.included).length}/{files.length}</span>
                <button
                  type="button"
                  className="btn-remove-project"
                  onClick={() => removeProject(projectName)}
                  title="Remove project"
                  aria-label={`Remove ${projectName}`}
                >
                  ✕
                </button>
              </div>
              <ul>
                {files.map((file) => (
                  <li
                    key={file.id}
                    draggable
                    onDragStart={() => onDragStart(projectName, file.id)}
                    onDragEnter={() => onDragEnter(projectName, file.id)}
                    onDragEnd={onDragEnd}
                    onDragOver={(e) => e.preventDefault()}
                    className={file.included ? '' : 'excluded'}
                  >
                    <span className="drag-handle" title="Drag to reorder">⠿</span>
                    <input type="checkbox" checked={file.included} onChange={() => toggleIncluded(file.id)} />
                    <span className="ext-badge">{file.extension}</span>
                    <span className="file-name" title={file.filePath}>{file.fileName}</span>
                  </li>
                ))}
              </ul>
            </div>
            )
          })}
        </section>
      )}

      {settingsOpen && (
        <div className="modal-overlay" role="dialog" aria-modal="true" aria-label="Extension Settings" onClick={closeSettings}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Extension Settings</h2>
              <button type="button" className="modal-close" onClick={closeSettings} aria-label="Close">×</button>
            </div>

            <div className="modal-body">
              <p className="modal-desc">
                Files with these extensions will be loaded when you upload a folder.
                Click <strong>Apply</strong> to re-filter already-uploaded files instantly — no re-upload needed.
              </p>
              <div className="ext-tags">
                {Array.from(pendingExts).sort().map((ext) => (
                  <span key={ext} className="ext-tag">
                    {ext}
                    <button
                      type="button"
                      className="ext-remove"
                      onClick={() => removePendingExt(ext)}
                      title={`Remove ${ext}`}
                      aria-label={`Remove ${ext}`}
                      disabled={pendingExts.size <= 1}
                    >
                      ×
                    </button>
                  </span>
                ))}
              </div>
              <div className="ext-add-row">
                <input
                  className="ext-input"
                  value={pendingInput}
                  onChange={(e) => { setPendingInput(e.target.value); setPendingError('') }}
                  onKeyDown={(e) => { if (e.key === 'Enter') addPendingExt() }}
                  placeholder=".txt"
                  aria-label="New extension"
                  autoFocus
                />
                <button type="button" onClick={addPendingExt} disabled={pendingInput.trim() === ''}>
                  Add
                </button>
              </div>
              {pendingError && <p className="ext-error">{pendingError}</p>}
            </div>

            <div className="modal-footer">
              <button type="button" className="btn-ghost" onClick={resetPendingToDefaults}>
                Reset Defaults
              </button>
              <span className="modal-footer-spacer" />
              <button type="button" onClick={closeSettings}>Cancel</button>
              <button type="button" className="primary" onClick={applySettings}>Apply</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default App
