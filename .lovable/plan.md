## Arquivos não usados no projeto

Após varrer todos os `import` em `src/` (exceto `routeTree.gen.ts`), encontrei o seguinte material morto:

### Arquivos a remover

1. **`src/components/ui/aspect-ratio.tsx`** — nenhum import em lugar nenhum. É o único arquivo dentro de `src/components/`, então a pasta `src/components/ui/` (e `src/components/`) também fica vazia e pode ser removida.
2. **`src/lib/utils.ts`** — exporta apenas `cn()`, que não é usado em nenhum lugar (varri `rg "\\bcn\\b" src` e `rg "from.*lib/utils" src` — zero resultados).
3. **`src/hooks/`** — diretório já está vazio, pode ser removido.

### Dependências `package.json` não usadas

Nada em `src/` importa delas, então podem sair com `bun remove`:

- Todos os `@radix-ui/react-*` **exceto** os que estão sendo usados (atualmente nenhum — `@radix-ui/react-aspect-ratio` só era usado pelo arquivo morto acima): `accordion`, `alert-dialog`, `aspect-ratio`, `avatar`, `checkbox`, `collapsible`, `context-menu`, `dialog`, `dropdown-menu`, `hover-card`, `label`, `menubar`, `navigation-menu`, `popover`, `progress`, `radio-group`, `scroll-area`, `select`, `separator`, `slider`, `slot`, `switch`, `tabs`, `toggle`, `toggle-group`, `tooltip`.
- `class-variance-authority`
- `cmdk`
- `input-otp`
- `react-hook-form`
- `sonner`
- `zod`
- `clsx` e `tailwind-merge` (só usados por `utils.ts`, que será removido)

### O que **fica**

`@azure/msal-browser`, `@tanstack/*`, `lucide-react`, `react`/`react-dom`, `tailwindcss` + `@tailwindcss/vite`, `tw-animate-css` (usado em `styles.css`), `vite-tsconfig-paths`, e todas as devDependencies.

### Observação

`src/components/ui/` é o destino padrão do shadcn — se você planeja adicionar componentes shadcn de novo em breve, talvez valha manter `utils.ts` (`cn`), `clsx` e `tailwind-merge`, já que praticamente todo componente shadcn depende deles. Me avise se prefere preservar esses três; nesse caso removo só o resto.
