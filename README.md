# Auditoria Patrimonial - Vite + React

Projeto convertido a partir do HTML original para uma estrutura padrão em Vite + React.

## Como rodar

```bash
npm install
npm run dev
```

Depois, acesse o endereço exibido no terminal, normalmente:

```text
http://localhost:5173
```

## Build de produção

```bash
npm run build
npm run preview
```

## Arquivos principais

- `src/App.jsx`: tela principal e lógica da auditoria.
- `src/main.jsx`: entrada do React.
- `src/styles.css`: estilos globais mínimos.
- `package.json`: dependências e scripts.

## Observações

As bibliotecas que antes eram carregadas via CDN no HTML agora estão instaladas como dependências npm:

- `react`
- `react-dom`
- `xlsx`
- `jszip`
