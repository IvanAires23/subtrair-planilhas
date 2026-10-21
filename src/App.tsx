import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { type Columns, type Result, type Row, type Sheet, detectColumns, subtract, toSheet } from './subtract'
import { OUTPUT_HEADER, toOutput } from './output'
import { sumSheets } from './sum'
import './App.css'

interface LoadedFile {
  name: string
  workbook: XLSX.WorkBook
  sheetName: string
  sheet: Sheet
  columns: Columns
}

type Tab = 'somar' | 'subtrair'

const ACCEPT = '.xlsx,.xls,.xlsm,.ods,.csv,.txt'

async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  const buffer = await file.arrayBuffer()
  if (/\.(csv|txt)$/i.test(file.name)) {
    // CSV exportado no Brasil costuma vir em Latin-1; tenta UTF-8 primeiro.
    let text: string
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(buffer)
    } catch {
      text = new TextDecoder('windows-1252').decode(buffer)
    }
    return XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true })
  }
  return XLSX.read(buffer, { type: 'array', cellDates: true })
}

function loadSheet(workbook: XLSX.WorkBook, sheetName: string): { sheet: Sheet; columns: Columns } {
  const matrix = XLSX.utils.sheet_to_json<Row>(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true })
  const sheet = toSheet(matrix)
  return { sheet, columns: detectColumns(sheet.header) }
}

async function loadFile(file: File): Promise<LoadedFile> {
  const workbook = await readWorkbook(file)
  const sheetName = workbook.SheetNames[0]
  return { name: file.name, workbook, sheetName, ...loadSheet(workbook, sheetName) }
}

function fromSheet(name: string, sheet: Sheet): LoadedFile {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([sheet.header, ...sheet.rows]), 'Soma')
  return { name, workbook, sheetName: 'Soma', sheet, columns: detectColumns(sheet.header) }
}

function download(data: Sheet, baseName: string, sheetName: string, format: 'xlsx' | 'csv') {
  const sheet = XLSX.utils.aoa_to_sheet([data.header, ...data.rows])
  if (format === 'csv') {
    // ";" e BOM UTF-8 para o Excel em português abrir com colunas e acentos certos.
    const csv = XLSX.utils.sheet_to_csv(sheet, { FS: ';' })
    const url = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }))
    const link = document.createElement('a')
    link.href = url
    link.download = `${baseName}.csv`
    link.click()
    URL.revokeObjectURL(url)
    return
  }
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName)
  XLSX.writeFile(workbook, `${baseName}.xlsx`)
}

const count = (n: number) => n.toLocaleString('pt-BR')
const baseName = (name: string) => name.replace(/\.[^.]+$/, '')

function columnLabel(header: Row, index: number): string {
  const name = String(header[index] ?? '').trim()
  return name || `Coluna ${XLSX.utils.encode_col(index)}`
}

function DropZone({
  multiple,
  onFiles,
  children,
}: {
  multiple?: boolean
  onFiles: (files: File[]) => void
  children: React.ReactNode
}) {
  const [dragging, setDragging] = useState(false)
  return (
    <label
      className={`drop${dragging ? ' dragging' : ''}`}
      onDragOver={e => {
        e.preventDefault()
        setDragging(true)
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => {
        e.preventDefault()
        setDragging(false)
        const files = Array.from(e.dataTransfer.files)
        if (files.length) onFiles(multiple ? files : files.slice(0, 1))
      }}
    >
      <input
        type="file"
        accept={ACCEPT}
        multiple={multiple}
        onChange={e => {
          const files = Array.from(e.target.files ?? [])
          if (files.length) onFiles(files)
          e.target.value = ''
        }}
      />
      {children}
    </label>
  )
}

function FilePicker({
  title,
  hint,
  file,
  onLoad,
  onChange,
}: {
  title: string
  hint: string
  file: LoadedFile | null
  onLoad: (file: File) => void
  onChange: (file: LoadedFile) => void
}) {
  const setColumn = (key: keyof Columns, value: string) => {
    if (!file) return
    onChange({ ...file, columns: { ...file.columns, [key]: value === '' ? null : Number(value) } })
  }

  const selectSheet = (sheetName: string) => {
    if (!file) return
    onChange({ ...file, sheetName, ...loadSheet(file.workbook, sheetName) })
  }

  const columnSelect = (key: keyof Columns, label: string) => (
    <label className="field">
      <span>{label}</span>
      <select value={file?.columns[key] ?? ''} onChange={e => setColumn(key, e.target.value)}>
        <option value="">— nenhuma —</option>
        {file?.sheet.header.map((_, i) => (
          <option key={i} value={i}>
            {columnLabel(file.sheet.header, i)}
          </option>
        ))}
      </select>
    </label>
  )

  return (
    <section className="card">
      <h2>{title}</h2>
      <p className="hint">{hint}</p>
      <DropZone onFiles={([picked]) => onLoad(picked)}>
        {file ? (
          <>
            <strong>{file.name}</strong>
            <span>{count(file.sheet.rows.length)} linhas · clique para trocar</span>
          </>
        ) : (
          <>
            <strong>Arraste o arquivo aqui</strong>
            <span>ou clique para escolher (.xlsx, .xls, .csv)</span>
          </>
        )}
      </DropZone>

      {file && (
        <div className="fields">
          {file.workbook.SheetNames.length > 1 && (
            <label className="field">
              <span>Aba</span>
              <select value={file.sheetName} onChange={e => selectSheet(e.target.value)}>
                {file.workbook.SheetNames.map(name => (
                  <option key={name}>{name}</option>
                ))}
              </select>
            </label>
          )}
          {columnSelect('name', 'Nome')}
          {columnSelect('doc', 'CPF / documento')}
          {columnSelect('ddd', 'DDD')}
          {columnSelect('phone', 'Telefone')}
          {columnSelect('email', 'E-mail')}
        </div>
      )}
    </section>
  )
}

function OutputNote() {
  return <p className="hint output-note">O arquivo sai só com as colunas: {OUTPUT_HEADER.join(', ')}.</p>
}

function SumPanel({
  files,
  setFiles,
  onUse,
  setError,
}: {
  files: LoadedFile[]
  setFiles: (update: (files: LoadedFile[]) => LoadedFile[]) => void
  onUse: (file: LoadedFile, target: 'A' | 'B') => void
  setError: (message: string) => void
}) {
  const total = useMemo(() => (files.length ? sumSheets(files) : null), [files])
  const name = `Soma de ${files.length} planilha${files.length === 1 ? '' : 's'}`

  const add = async (picked: File[]) => {
    setError('')
    const loaded: LoadedFile[] = []
    const failed: string[] = []
    for (const file of picked) {
      try {
        loaded.push(await loadFile(file))
      } catch {
        failed.push(file.name)
      }
    }
    if (failed.length) setError(`Não consegui ler: ${failed.join(', ')}.`)
    setFiles(current => [...current, ...loaded])
  }

  const move = (index: number, delta: number) =>
    setFiles(current => {
      const next = [...current]
      const [item] = next.splice(index, 1)
      next.splice(index + delta, 0, item)
      return next
    })

  return (
    <>
      <section className="card">
        <h2>Planilhas para somar</h2>
        <p className="hint">
          Junta todas as linhas numa planilha só. Colunas com o mesmo nome ficam juntas, e uma coluna “Planilha de
          origem” mostra de onde veio cada linha. Repetidos são mantidos.
        </p>
        <DropZone multiple onFiles={add}>
          <strong>Arraste as planilhas aqui</strong>
          <span>ou clique para escolher (pode selecionar várias de uma vez)</span>
        </DropZone>

        {files.length > 0 && (
          <ol className="file-list">
            {files.map((file, i) => (
              <li key={`${file.name}-${i}`}>
                <span className="file-name">{file.name}</span>
                <span className="file-rows">{count(file.sheet.rows.length)} linhas</span>
                <span className="file-actions">
                  <button className="icon" disabled={i === 0} onClick={() => move(i, -1)} aria-label="Subir">
                    ↑
                  </button>
                  <button
                    className="icon"
                    disabled={i === files.length - 1}
                    onClick={() => move(i, 1)}
                    aria-label="Descer"
                  >
                    ↓
                  </button>
                  <button
                    className="icon"
                    onClick={() => setFiles(current => current.filter((_, j) => j !== i))}
                    aria-label="Remover"
                  >
                    ×
                  </button>
                </span>
              </li>
            ))}
          </ol>
        )}
      </section>

      {total && (
        <section className="card result">
          <h2>Resultado da soma</h2>
          <dl>
            <div>
              <dt>Planilhas</dt>
              <dd>{count(files.length)}</dd>
            </div>
            <div className="total">
              <dt>Linhas no total</dt>
              <dd>{count(total.rows.length)}</dd>
            </div>
          </dl>
          <div className="downloads">
            {(['xlsx', 'csv'] as const).map(format => (
              <button
                key={format}
                onClick={() => download(toOutput(total, detectColumns(total.header)), name, 'Soma', format)}
              >
                Baixar .{format}
              </button>
            ))}
            <button className="secondary" onClick={() => onUse(fromSheet(name, total), 'A')}>
              Usar como planilha A →
            </button>
            <button className="secondary" onClick={() => onUse(fromSheet(name, total), 'B')}>
              Usar como planilha B →
            </button>
          </div>
          <OutputNote />
        </section>
      )}
    </>
  )
}

export default function App() {
  const [tab, setTab] = useState<Tab>('subtrair')
  const [sumFiles, setSumFiles] = useState<LoadedFile[]>([])
  const [fileA, setFileA] = useState<LoadedFile | null>(null)
  const [fileB, setFileB] = useState<LoadedFile | null>(null)
  const [result, setResult] = useState<Result | null>(null)
  const [error, setError] = useState('')

  const load = (setter: (file: LoadedFile) => void) => async (file: File) => {
    setError('')
    setResult(null)
    try {
      setter(await loadFile(file))
    } catch {
      setError(`Não consegui ler "${file.name}". Confira se é uma planilha válida.`)
    }
  }

  const change = (setter: (file: LoadedFile) => void) => (file: LoadedFile) => {
    setResult(null)
    setter(file)
  }

  const sendSum = (file: LoadedFile, target: 'A' | 'B') => {
    setResult(null)
    if (target === 'A') setFileA(file)
    else setFileB(file)
    setTab('subtrair')
  }

  const hasKey = (file: LoadedFile | null) =>
    file !== null && (file.columns.doc !== null || file.columns.phone !== null || file.columns.email !== null)
  const ready = hasKey(fileA) && hasKey(fileB)

  const run = () => {
    if (!fileA || !fileB) return
    setResult(subtract(fileA.sheet, fileA.columns, fileB.sheet, fileB.columns))
  }

  const switchTab = (next: Tab) => {
    setError('')
    setTab(next)
  }

  return (
    <main>
      <header>
        <h1>Planilhas</h1>
        <nav className="tabs" role="tablist">
          <button
            role="tab"
            aria-selected={tab === 'somar'}
            className={tab === 'somar' ? 'active' : ''}
            onClick={() => switchTab('somar')}
          >
            Somar
          </button>
          <button
            role="tab"
            aria-selected={tab === 'subtrair'}
            className={tab === 'subtrair' ? 'active' : ''}
            onClick={() => switchTab('subtrair')}
          >
            Subtrair
          </button>
        </nav>
        {tab === 'somar' ? (
          <p>
            <b>A</b> + <b>B</b> + <b>C</b> + … Junta várias planilhas numa só.
          </p>
        ) : (
          <p>
            <b>A</b> − <b>B</b> = <b>C</b>. Tira da planilha A todo mundo cujo <b>CPF</b>, <b>telefone</b> ou{' '}
            <b>e-mail</b> aparece na planilha B.
          </p>
        )}
      </header>

      {tab === 'somar' ? (
        <SumPanel files={sumFiles} setFiles={setSumFiles} onUse={sendSum} setError={setError} />
      ) : (
        <>
          <div className="grid">
            <FilePicker
              title="Planilha A"
              hint="A base completa (a maior)."
              file={fileA}
              onLoad={load(setFileA)}
              onChange={change(setFileA)}
            />
            <FilePicker
              title="Planilha B"
              hint="Quem deve ser removido de A."
              file={fileB}
              onLoad={load(setFileB)}
              onChange={change(setFileB)}
            />
          </div>

          {(fileA || fileB) && !ready && !error && (
            <p className="hint center">Escolha pelo menos uma coluna (CPF, telefone ou e-mail) em cada planilha.</p>
          )}

          <div className="actions">
            <button onClick={run} disabled={!ready}>
              Subtrair
            </button>
          </div>

          {result && fileA && (
            <section className="card result">
              <h2>Planilha C</h2>
              <dl>
                <div>
                  <dt>Linhas em A</dt>
                  <dd>{count(fileA.sheet.rows.length)}</dd>
                </div>
                <div>
                  <dt>Removidas por CPF</dt>
                  <dd>{count(result.removedByDoc)}</dd>
                </div>
                <div>
                  <dt>Removidas por telefone</dt>
                  <dd>{count(result.removedByPhone)}</dd>
                </div>
                <div>
                  <dt>Removidas por e-mail</dt>
                  <dd>{count(result.removedByEmail)}</dd>
                </div>
                <div className="total">
                  <dt>Ficaram em C</dt>
                  <dd>{count(result.rows.length)}</dd>
                </div>
              </dl>
              <div className="downloads">
                {(['xlsx', 'csv'] as const).map(format => (
                  <button
                    key={format}
                    onClick={() =>
                      download(
                        toOutput({ header: fileA.sheet.header, rows: result.rows }, fileA.columns),
                        `${baseName(fileA.name)} - filtrada`,
                        'Planilha C',
                        format,
                      )
                    }
                  >
                    Baixar .{format}
                  </button>
                ))}
              </div>
              <OutputNote />
            </section>
          )}
        </>
      )}

      {error && <p className="error">{error}</p>}

      <footer>Os arquivos são processados só no seu navegador. Nada é enviado para servidor.</footer>
    </main>
  )
}
