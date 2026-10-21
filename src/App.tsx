import { useState } from 'react'
import * as XLSX from 'xlsx'
import { type Columns, type Result, type Row, type Sheet, detectColumns, subtract, toSheet } from './subtract'
import './App.css'

interface LoadedFile {
  name: string
  workbook: XLSX.WorkBook
  sheetName: string
  sheet: Sheet
  columns: Columns
}

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

function columnLabel(header: Row, index: number): string {
  const name = String(header[index] ?? '').trim()
  return name || `Coluna ${XLSX.utils.encode_col(index)}`
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
  const [dragging, setDragging] = useState(false)

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
          const dropped = e.dataTransfer.files[0]
          if (dropped) onLoad(dropped)
        }}
      >
        <input
          type="file"
          accept=".xlsx,.xls,.xlsm,.ods,.csv,.txt"
          onChange={e => {
            const picked = e.target.files?.[0]
            if (picked) onLoad(picked)
            e.target.value = ''
          }}
        />
        {file ? (
          <>
            <strong>{file.name}</strong>
            <span>{file.sheet.rows.length.toLocaleString('pt-BR')} linhas · clique para trocar</span>
          </>
        ) : (
          <>
            <strong>Arraste o arquivo aqui</strong>
            <span>ou clique para escolher (.xlsx, .xls, .csv)</span>
          </>
        )}
      </label>

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
          {columnSelect('doc', 'CPF / documento')}
          {columnSelect('ddd', 'DDD')}
          {columnSelect('phone', 'Telefone')}
        </div>
      )}
    </section>
  )
}

export default function App() {
  const [fileA, setFileA] = useState<LoadedFile | null>(null)
  const [fileB, setFileB] = useState<LoadedFile | null>(null)
  const [result, setResult] = useState<Result | null>(null)
  const [error, setError] = useState('')

  const load = (setter: (file: LoadedFile) => void) => async (file: File) => {
    setError('')
    setResult(null)
    try {
      const workbook = await readWorkbook(file)
      const sheetName = workbook.SheetNames[0]
      setter({ name: file.name, workbook, sheetName, ...loadSheet(workbook, sheetName) })
    } catch {
      setError(`Não consegui ler "${file.name}". Confira se é uma planilha válida.`)
    }
  }

  const change = (setter: (file: LoadedFile) => void) => (file: LoadedFile) => {
    setResult(null)
    setter(file)
  }

  const hasKey = (file: LoadedFile | null) => file !== null && (file.columns.doc !== null || file.columns.phone !== null)
  const ready = hasKey(fileA) && hasKey(fileB)

  const run = () => {
    if (!fileA || !fileB) return
    setResult(subtract(fileA.sheet, fileA.columns, fileB.sheet, fileB.columns))
  }

  const download = (format: 'xlsx' | 'csv') => {
    if (!fileA || !result) return
    const sheet = XLSX.utils.aoa_to_sheet([fileA.sheet.header, ...result.rows])
    const base = `${fileA.name.replace(/\.[^.]+$/, '')} - filtrada`
    if (format === 'csv') {
      // ";" e BOM UTF-8 para o Excel em português abrir com colunas e acentos certos.
      const csv = XLSX.utils.sheet_to_csv(sheet, { FS: ';' })
      const url = URL.createObjectURL(new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' }))
      const link = document.createElement('a')
      link.href = url
      link.download = `${base}.csv`
      link.click()
      URL.revokeObjectURL(url)
      return
    }
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'Planilha C')
    XLSX.writeFile(workbook, `${base}.xlsx`)
  }

  return (
    <main>
      <header>
        <h1>Subtrair planilhas</h1>
        <p>
          <b>A</b> − <b>B</b> = <b>C</b>. Tira da planilha A todo mundo cujo <b>CPF</b> ou <b>telefone</b> aparece na
          planilha B.
        </p>
      </header>

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

      {error && <p className="error">{error}</p>}
      {(fileA || fileB) && !ready && !error && (
        <p className="hint center">Escolha pelo menos a coluna de CPF ou de telefone em cada planilha.</p>
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
              <dd>{fileA.sheet.rows.length.toLocaleString('pt-BR')}</dd>
            </div>
            <div>
              <dt>Removidas por CPF</dt>
              <dd>{result.removedByDoc.toLocaleString('pt-BR')}</dd>
            </div>
            <div>
              <dt>Removidas por telefone</dt>
              <dd>{result.removedByPhone.toLocaleString('pt-BR')}</dd>
            </div>
            <div className="total">
              <dt>Ficaram em C</dt>
              <dd>{result.rows.length.toLocaleString('pt-BR')}</dd>
            </div>
          </dl>
          <div className="downloads">
            <button onClick={() => download('xlsx')}>Baixar .xlsx</button>
            <button onClick={() => download('csv')}>Baixar .csv</button>
          </div>
        </section>
      )}

      <footer>Os arquivos são processados só no seu navegador. Nada é enviado para servidor.</footer>
    </main>
  )
}
