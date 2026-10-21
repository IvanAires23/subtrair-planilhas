# Subtrair planilhas

**A − B = C.** Sobe a planilha A (base completa) e a planilha B (quem deve sair) e baixa a planilha C
só com as linhas de A que **não** estão em B.

Uma linha de A é removida quando o **CPF/documento** **ou** o **telefone** dela aparece em B.

Tudo roda no navegador: os arquivos nunca são enviados para servidor.

## Como compara

- **Colunas:** detectadas pelo cabeçalho (`CPF`, `Documento`, `CNPJ`, `DDD`, `Telefone`, `Celular`, `Fone`,
  `WhatsApp`…). Dá para trocar na mão em cada planilha. Se as primeiras linhas tiverem título, o cabeçalho é
  procurado nas 20 primeiras.
- **CPF/CNPJ:** só os dígitos, recompletando zeros à esquerda que o Excel tira (`012.345.678-90` = `1234567890`).
- **Telefone:** junta a coluna de DDD (quando existe) com a de telefone, ignora `+55`, zero de operadora e
  pontuação, e compara DDD + últimos 8 dígitos (assim o 9º dígito do celular não atrapalha).
  `(11) 98765-4321`, `+55 11 987654321` e DDD `11` + `8765-4321` são o mesmo número.
  Se um dos lados não tiver DDD, compara só pelo número.
- **Formatos:** `.xlsx`, `.xls`, `.ods`, `.csv` (UTF-8 ou Latin-1, com `,` ou `;`). A planilha C sai em `.xlsx` ou `.csv` (separado por `;`, abre direto no Excel), com as
  colunas originais de A.

## Desenvolvimento

```bash
npm install
npm run dev      # http://localhost:5173
npm test         # testes da lógica de comparação
npm run build    # gera dist/ (site estático: Vercel, Netlify, GitHub Pages)
```

A lógica fica em `src/subtract.ts`; a interface em `src/App.tsx`.
