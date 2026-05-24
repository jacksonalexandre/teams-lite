## Objetivo
Adicionar um atalho de "Chamada de vídeo" visível apenas em chats 1:1, que abre o Teams Web em nova aba para iniciar a chamada.

## Deep-link usado
Para chat 1:1 no Teams Web:
```
https://teams.microsoft.com/l/call/0/0?users=<email_do_outro>&withVideo=true
```
O campo `users` aceita e-mail (UPN) separados por vírgula. Incluímos todos os membros do chat exceto o próprio usuário (em 1:1 só há o outro).

## Mudanças

### `src/lib/graph.ts`
Adicionar helper `buildTeamsVideoLink(members, meId): string | null`:
- Pega o array de membros do chat, filtra quem não é `meId`.
- Verifica se o membro restante tem `email` ou `userPrincipalName`.
- Retorna a URL montada ou `null` se faltarem dados.

### `src/routes/index.tsx`
No header da conversa (`headerTitle` area), condicionalmente renderizar à direita:
- Quando `selection.kind === "chat"` e o chat correspondente tem `chatType === "oneOnOne"` (ou ausente, que significa 1:1):
  - Um pequeno botão circular com ícone `Video` (lucide-react).
  - `disabled` quando `buildTeamsVideoLink` retorna `null`.
  - `onClick` → `window.open(url, "_blank", "noopener")`.

## Escopo
Apenas UI + construção de URL. Nenhuma nova chamada Graph, nenhum estado, nenhum backend.

## Resultado
Botão de vídeo aparece no topo apenas em conversas 1:1. Ao clicar, abre o Teams Web já solicitando chamada com vídeo.