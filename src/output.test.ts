import { describe, expect, it } from 'vitest'
import { joinPhone, OUTPUT_HEADER, toOutput } from './output'
import { detectColumns } from './subtract'

describe('joinPhone', () => {
  it('junta DDD e telefone sem repetir DDD', () => {
    expect(joinPhone('11', '980602608')).toBe('11980602608')
    expect(joinPhone(11, 980602608)).toBe('11980602608')
    expect(joinPhone('(011)', '98060-2608')).toBe('11980602608')
    expect(joinPhone('11', '(11) 98060-2608')).toBe('11980602608')
    expect(joinPhone('', '+55 11 98060-2608')).toBe('11980602608')
    expect(joinPhone('11', '')).toBe('')
  })
})

describe('toOutput', () => {
  it('deixa só nome, e-mail, documento, DDD, telefone e telefone com DDD', () => {
    const header = ['Nome do Produto', 'Documento do Produtor', 'Nome', 'Documento', 'Email', 'DDD', 'Telefone', 'CEP']
    const sheet = {
      header,
      rows: [['Imersão', '348', ' Ana ', 2290023230, 'ana@x.com', '94', '999039374', '68500']],
    }
    expect(toOutput(sheet, detectColumns(header))).toEqual({
      header: OUTPUT_HEADER,
      rows: [['Ana', 'ana@x.com', '02290023230', '94', '999039374', '94999039374']],
    })
  })
})
