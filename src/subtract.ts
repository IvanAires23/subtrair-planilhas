export type Cell = string | number | boolean | Date | null | undefined
export type Row = Cell[]

export interface Columns {
  name: number | null
  doc: number | null
  ddd: number | null
  phone: number | null
  email: number | null
}

export interface Sheet {
  header: Row
  rows: Row[]
}

export interface Result {
  rows: Row[]
  removedByDoc: number
  removedByPhone: number
  removedByEmail: number
}

function onlyDigits(value: Cell): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'number') return Number.isInteger(value) ? String(value) : String(Math.round(value))
  return String(value).replace(/\D/g, '')
}

/** CPF/CNPJ só com dígitos; zeros à esquerda que o Excel removeu são recompletados. */
export function normalizeDoc(value: Cell): string {
  const digits = onlyDigits(value)
  if (!digits || /^0+$/.test(digits)) return ''
  if (digits.length <= 11) return digits.padStart(11, '0')
  if (digits.length <= 14) return digits.padStart(14, '0')
  return digits
}

export function normalizeEmail(value: Cell): string {
  const email = String(value ?? '').trim().toLowerCase()
  return email.includes('@') ? email : ''
}

export interface PhoneKey {
  /** DDD de 2 dígitos, ou '' quando não dá para saber. */
  ddd: string
  /** Últimos 8 dígitos do número, assim o 9º dígito de celular não atrapalha. */
  last8: string
}

/**
 * Junta a coluna de DDD (se houver) com a de telefone e reduz para DDD + últimos 8 dígitos.
 * Aceita "+55 (11) 98765-4321", "011987654321", DDD em coluna separada etc.
 */
export function normalizePhone(dddValue: Cell, phoneValue: Cell): PhoneKey | null {
  let ddd = onlyDigits(dddValue).replace(/^0+/, '')
  let number = onlyDigits(phoneValue).replace(/^0+/, '')

  // O número pode já vir com o DDD (e até o 55) mesmo quando existe coluna de DDD.
  if (number.length >= 12 && number.startsWith('55')) number = number.slice(2)
  if (number.length >= 10) {
    ddd = number.slice(0, 2)
    number = number.slice(2)
  }
  if (number.length < 8) return null
  if (ddd.length !== 2) ddd = ''

  return { ddd, last8: number.slice(-8) }
}

/** Remove de A toda linha cujo documento, telefone OU e-mail aparece em B. */
export function subtract(a: Sheet, colsA: Columns, b: Sheet, colsB: Columns): Result {
  const docsB = new Set<string>()
  const phonesB = new Set<string>()
  const last8B = new Set<string>()
  const last8NoDddB = new Set<string>()
  const emailsB = new Set<string>()

  for (const row of b.rows) {
    if (colsB.doc !== null) {
      const doc = normalizeDoc(row[colsB.doc])
      if (doc) docsB.add(doc)
    }
    if (colsB.phone !== null) {
      const phone = normalizePhone(colsB.ddd === null ? '' : row[colsB.ddd], row[colsB.phone])
      if (phone) {
        last8B.add(phone.last8)
        if (phone.ddd) phonesB.add(phone.ddd + phone.last8)
        else last8NoDddB.add(phone.last8)
      }
    }
    if (colsB.email !== null) {
      const email = normalizeEmail(row[colsB.email])
      if (email) emailsB.add(email)
    }
  }

  const rows: Row[] = []
  let removedByDoc = 0
  let removedByPhone = 0
  let removedByEmail = 0

  for (const row of a.rows) {
    if (colsA.doc !== null) {
      const doc = normalizeDoc(row[colsA.doc])
      if (doc && docsB.has(doc)) {
        removedByDoc++
        continue
      }
    }
    if (colsA.phone !== null) {
      const phone = normalizePhone(colsA.ddd === null ? '' : row[colsA.ddd], row[colsA.phone])
      if (phone) {
        // Sem DDD de um dos lados, compara só pelo número.
        const match = phone.ddd
          ? phonesB.has(phone.ddd + phone.last8) || last8NoDddB.has(phone.last8)
          : last8B.has(phone.last8)
        if (match) {
          removedByPhone++
          continue
        }
      }
    }
    if (colsA.email !== null) {
      const email = normalizeEmail(row[colsA.email])
      if (email && emailsB.has(email)) {
        removedByEmail++
        continue
      }
    }
    rows.push(row)
  }

  return { rows, removedByDoc, removedByPhone, removedByEmail }
}

function normalizeHeader(value: Cell): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .trim()
}

const DOC_PATTERN = /cpf|cnpj|documento|^doc\b/
const DDD_PATTERN = /ddd/
const PHONE_PATTERN = /telefone|celular|fone|whats|zap|^tel\b|^tel\d|movel/
const EMAIL_PATTERN = /e-?mail/
const NAME_PATTERN = /nome|name|cliente|comprador/
/** Colunas que não são do cliente, como "Documento do Produtor" nos relatórios da Hotmart. */
const NOT_CUSTOMER = /produtor|afiliad|vendedor|empresa|fornecedor/

function findColumn(
  header: Row,
  exact: string[],
  pattern: RegExp,
  skip: (h: string) => boolean = () => false,
): number | null {
  const names = header.map(normalizeHeader)
  // Nome exato ganha de nome parecido: "Documento" antes de "Documento do Produtor".
  const exactIndex = names.findIndex(h => exact.includes(h))
  if (exactIndex !== -1) return exactIndex
  const index = names.findIndex(h => h !== '' && pattern.test(h) && !NOT_CUSTOMER.test(h) && !skip(h))
  return index === -1 ? null : index
}

export function detectColumns(header: Row): Columns {
  return {
    name: findColumn(header, ['nome', 'nome completo', 'name', 'cliente', 'comprador'], NAME_PATTERN, h =>
      /produto|oferta|cupom|arquivo|planilha/.test(h) ||
      [DOC_PATTERN, DDD_PATTERN, PHONE_PATTERN, EMAIL_PATTERN].some(p => p.test(h)),
    ),
    doc: findColumn(header, ['cpf', 'documento', 'doc', 'cpf/cnpj', 'cpf_cnpj', 'cnpj'], DOC_PATTERN),
    ddd: findColumn(header, ['ddd'], DDD_PATTERN),
    phone: findColumn(
      header,
      ['telefone', 'celular', 'fone', 'whatsapp', 'tel'],
      PHONE_PATTERN,
      h => DDD_PATTERN.test(h) || DOC_PATTERN.test(h),
    ),
    email: findColumn(header, ['email', 'e-mail'], EMAIL_PATTERN),
  }
}

/** Acha a linha de cabeçalho (às vezes a planilha tem título nas primeiras linhas). */
export function toSheet(matrix: Row[]): Sheet {
  const limit = Math.min(matrix.length, 20)
  let headerIndex = 0
  for (let i = 0; i < limit; i++) {
    const cols = detectColumns(matrix[i] ?? [])
    if (cols.doc !== null || cols.phone !== null || cols.email !== null) {
      headerIndex = i
      break
    }
  }
  const rows = matrix
    .slice(headerIndex + 1)
    .filter(row => row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''))
  return { header: matrix[headerIndex] ?? [], rows }
}
