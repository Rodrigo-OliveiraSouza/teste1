# Gerador de DOCX em lote

App em Streamlit que abre um DOCX modelo, mapeia textos a colunas de uma planilha (Excel/CSV) e gera um documento por linha.

## Passos rápidos
- Instale deps: `pip install -r requirements.txt`
- Rode: `streamlit run docx_batch_app.py`
- Na interface:
  - Envie o DOCX modelo.
  - Envie a planilha.
  - Na tabela de mapeamento, preencha:
    - **Texto no DOCX**: o texto exato que será substituído.
    - **Coluna do Excel**: a coluna cujo valor entra no lugar.
  - Opcional: ajuste o template do nome do arquivo.
  - Clique em **Gerar documentos**. Os arquivos são salvos em `saida_docx` e também podem ser baixados em ZIP.

## Notas
- Para encontrar o texto, use exatamente o que está no DOCX (evite quebras/formatação dividindo o placeholder).
- A substituição recompõe o parágrafo em um único run; se o trecho tiver estilos mistos (negrito dentro do placeholder), esse formato pode ser perdido. Planeje placeholders simples.

## Versão web (Node/JS)
- Arquivos: `server.js`, `public/index.html`, `package.json`.
- Instale deps Node: `npm install`
- Rode: `npm start` e abra `http://localhost:3000`
- Fluxo na interface web:
  - Envie o template `.docx` com placeholders `{{campo}}`.
  - Envie a planilha (Excel/CSV); o servidor devolve colunas.
  - Associe cada placeholder a uma coluna e clique em “Gerar ZIP” para baixar todos os DOCX gerados.
- Observações:
  - O servidor conserta delimitadores repetidos (`{{{{tag}}}}` -> `{{tag}}`) e preenche placeholders não mapeados com vazio para evitar erros.
