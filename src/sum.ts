import type { Row, Sheet } from './subtract'

export const ORIGIN_COLUMN = 'Planilha de origem'

function headerKey(value: unknown): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .trim()
}

/**
 * Empilha as linhas de várias planilhas numa só. Colunas com o mesmo nome se juntam,
 * colunas que só existem em algumas ficam vazias nas outras. Repetidos são mantidos.
 * Nomes repetidos dentro da mesma planilha (ex.: 3× "Moeda de recebimento") continuam separados.
 */
export function sumSheets(files: { name: string; sheet: Sheet }[]): Sheet {
  const header: Row = []
  const keys: string[] = []
  const plans = files.map(({ sheet }) => {
    const seen = new Map<string, number>()
    return sheet.header.map((cell, i) => {
      const base = headerKey(cell) || `#coluna${i}`
      const occurrence = seen.get(base) ?? 0
      seen.set(base, occurrence + 1)
      const key = `${base}#${occurrence}`
      let target = keys.indexOf(key)
      if (target === -1) {
        target = keys.push(key) - 1
        header.push(cell)
      }
      return target
    })
  })

  header.push(ORIGIN_COLUMN)
  const width = header.length
  const rows: Row[] = []
  files.forEach(({ name, sheet }, f) => {
    for (const row of sheet.rows) {
      const out: Row = new Array(width).fill('')
      plans[f].forEach((target, i) => {
        if (i < row.length) out[target] = row[i]
      })
      out[width - 1] = name
      rows.push(out)
    }
  })
  return { header, rows }
}
