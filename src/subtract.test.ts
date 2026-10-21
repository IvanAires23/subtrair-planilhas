import { describe, expect, it } from 'vitest'
import { detectColumns, normalizeDoc, normalizePhone, subtract, toSheet } from './subtract'

describe('normalizeDoc', () => {
  it('ignora pontuação e recompleta zeros à esquerda', () => {
    expect(normalizeDoc('012.345.678-90')).toBe('01234567890')
    expect(normalizeDoc(1234567890)).toBe('01234567890')
    expect(normalizeDoc('')).toBe('')
    expect(normalizeDoc('000.000.000-00')).toBe('')
  })
})

describe('normalizePhone', () => {
  it('aceita DDD separado ou junto, com ou sem 55 e 9º dígito', () => {
    const key = { ddd: '11', last8: '87654321' }
    expect(normalizePhone('11', '98765-4321')).toEqual(key)
    expect(normalizePhone('(11)', '8765-4321')).toEqual(key)
    expect(normalizePhone('', '+55 (11) 98765-4321')).toEqual(key)
    expect(normalizePhone('', '011987654321')).toEqual(key)
    expect(normalizePhone(11, 11987654321)).toEqual(key)
    expect(normalizePhone('', '98765-4321')).toEqual({ ddd: '', last8: '87654321' })
    expect(normalizePhone('11', '123')).toBeNull()
  })
})

describe('detectColumns', () => {
  it('acha CPF, DDD e telefone pelo cabeçalho', () => {
    expect(detectColumns(['Nome', 'CPF', 'DDD', 'Telefone'])).toEqual({ doc: 1, ddd: 2, phone: 3, email: null })
    expect(detectColumns(['Documento', 'Celular'])).toEqual({ doc: 0, ddd: null, phone: 1, email: null })
  })

  it('ignora colunas do produtor no relatório da Hotmart', () => {
    const header = ['Nome do Produto', 'Documento do Produtor', 'Nome', 'Documento', 'Email', 'DDD', 'Telefone']
    expect(detectColumns(header)).toEqual({ doc: 3, ddd: 5, phone: 6, email: 4 })
    expect(detectColumns(['Documento do Produtor', 'CPF do cliente', 'Tel'])).toEqual({ doc: 1, ddd: null, phone: 2, email: null })
  })
})

describe('subtract', () => {
  const a = toSheet([
    ['Relatório'],
    ['Nome', 'CPF', 'DDD', 'Telefone'],
    ['Ana', '111.111.111-11', '11', '98888-0001'],
    ['Bia', '222.222.222-22', '21', '97777-0002'],
    ['Caio', '333.333.333-33', '31', '96666-0003'],
    ['Davi', '444.444.444-44', '41', '95555-0004'],
    ['', '', '', ''],
  ])
  const b = toSheet([
    ['documento', 'celular'],
    ['11111111111', '(99) 90000-0000'],
    ['', '+55 21 977770002'],
    ['', '88880001'],
  ])

  it('remove quem bate CPF ou telefone', () => {
    const result = subtract(a, detectColumns(a.header), b, detectColumns(b.header))
    expect(result.rows.map(r => r[0])).toEqual(['Caio', 'Davi'])
    expect(result.removedByDoc).toBe(1)
    expect(result.removedByPhone).toBe(1)
  })

  it('não remove por telefone de outro DDD', () => {
    const b2 = toSheet([['CPF', 'DDD', 'Telefone'], ['', '51', '96666-0003']])
    const result = subtract(a, detectColumns(a.header), b2, detectColumns(b2.header))
    expect(result.rows).toHaveLength(4)
  })
})

describe('subtract por e-mail', () => {
  it('remove quem só bate o e-mail', () => {
    const a = toSheet([
      ['Nome', 'Documento', 'Email', 'DDD', 'Telefone'],
      ['Ana', '', ' Ana@Gmail.com ', '', ''],
      ['Bia', '', 'bia@gmail.com', '', ''],
    ])
    const b = toSheet([['E-mail'], ['ana@gmail.com']])
    const result = subtract(a, detectColumns(a.header), b, detectColumns(b.header))
    expect(result.rows.map(r => r[0])).toEqual(['Bia'])
    expect(result.removedByEmail).toBe(1)
  })
})
