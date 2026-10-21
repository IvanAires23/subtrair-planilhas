import { type Cell, type Columns, type Row, type Sheet, normalizeDoc } from './subtract'

export const OUTPUT_HEADER = ['Nome', 'E-mail', 'Documento', 'DDD', 'Telefone', 'Telefone com DDD']

const text = (value: Cell) => (value === null || value === undefined ? '' : String(value).trim())
const digits = (value: Cell) => text(value).replace(/\D/g, '')

/** DDD e telefone juntos, só dígitos: "11" + "98060-2608" → "11980602608". */
export function joinPhone(ddd: Cell, phone: Cell): string {
  let number = digits(phone).replace(/^0+/, '')
  if (number.length >= 12 && number.startsWith('55')) number = number.slice(2)
  // O telefone já veio com DDD: não repete.
  if (number.length >= 10) return number
  if (!number) return ''
  return digits(ddd).replace(/^0+/, '') + number
}

/** Reduz a planilha às colunas que interessam: nome, e-mail, documento, DDD, telefone e telefone com DDD. */
export function toOutput(sheet: Sheet, columns: Columns): Sheet {
  const get = (row: Row, index: number | null): Cell => (index === null ? '' : row[index])
  const rows = sheet.rows.map(row => {
    const doc = get(row, columns.doc)
    const ddd = get(row, columns.ddd)
    const phone = get(row, columns.phone)
    return [
      text(get(row, columns.name)),
      text(get(row, columns.email)),
      // Documento que virou número no Excel perdeu os zeros à esquerda; texto fica como veio.
      typeof doc === 'number' ? normalizeDoc(doc) : text(doc),
      digits(ddd),
      text(phone),
      joinPhone(ddd, phone),
    ]
  })
  return { header: [...OUTPUT_HEADER], rows }
}
