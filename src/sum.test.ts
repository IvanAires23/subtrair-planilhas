import { describe, expect, it } from 'vitest'
import { ORIGIN_COLUMN, sumSheets } from './sum'

describe('sumSheets', () => {
  it('junta colunas pelo nome, mantém repetidos e marca a origem', () => {
    const result = sumSheets([
      { name: 'a.xlsx', sheet: { header: ['Nome', 'CPF', 'Moeda', 'Moeda'], rows: [['Ana', '1', 'BRL', 'USD']] } },
      { name: 'b.csv', sheet: { header: ['cpf', 'nome', 'Email'], rows: [['1', 'Ana', 'a@x'], ['2', 'Bia', 'b@x']] } },
    ])
    expect(result.header).toEqual(['Nome', 'CPF', 'Moeda', 'Moeda', 'Email', ORIGIN_COLUMN])
    expect(result.rows).toEqual([
      ['Ana', '1', 'BRL', 'USD', '', 'a.xlsx'],
      ['Ana', '1', '', '', 'a@x', 'b.csv'],
      ['Bia', '2', '', '', 'b@x', 'b.csv'],
    ])
  })
})
