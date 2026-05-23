# Carregar fotos dos usuários

Hoje o `Avatar` em `src/routes/index.tsx` mostra apenas iniciais coloridas — nunca chegamos a buscar a foto real no Microsoft Graph. Para exibir as imagens corretamente faltam quatro coisas:

## 1. Escopo de permissão
Em `src/lib/msal.ts`, adicionar `User.ReadBasic.All` aos `GRAPH_SCOPES`. Sem ele, só conseguimos a foto do próprio usuário (`/me/photo/$value`), não a dos outros membros do chat. Usuários existentes precisarão re-consentir (popup automático na próxima chamada).

## 2. Helper para buscar a foto
Em `src/lib/graph.ts`, criar `getUserPhotoUrl(cfg, userId)`:
- Faz `GET /users/{userId}/photo/$value` com `Authorization: Bearer …` (não usa `graphFetch` porque a resposta é binária, não JSON).
- Em sucesso, converte para `Blob` e retorna `URL.createObjectURL(blob)`.
- Em 404 (usuário sem foto) ou erro, retorna `null`.

## 3. Cache em memória
Criar um `Map<userId, Promise<string | null>>` no módulo para deduplicar requisições — cada userId é buscado uma única vez por sessão. Sem cache, cada render do chat dispararia novas chamadas.

## 4. Wiring no componente
Em `src/routes/index.tsx`:
- Estender `Avatar` para aceitar `userId?: string` e, quando presente, usar um hook (`useUserPhoto`) que consulta o cache e dispara o fetch.
- Enquanto a foto carrega ou se falhar, mantém o fallback atual (iniciais + cor por nome).
- Passar `userId` nos dois pontos de uso:
  - Lista de chats (linha ~436): para `oneOnOne`, usar o `userId` do outro membro; manter ícone de grupo nos demais.
  - Header da conversa selecionada (linha ~490 — header do chat aberto): mesmo critério.
- Nas mensagens (`MessageBubble`), também mostrar avatar do remetente usando `message.from.user.id` — hoje o balão não tem avatar; este é o lugar onde a falta de foto é mais visível no Teams Web.

## Detalhes técnicos

- A resposta de `/photo/$value` é uma imagem JPEG. Object URLs são criadas uma vez e reutilizadas pelo cache; não precisamos revogar durante a sessão.
- Tratar 404 silenciosamente (muitos usuários não têm foto definida) — não logar como erro.
- O cache é por `userId`, então funciona tanto para membros de chats quanto para remetentes de mensagens em canais.
- Não tocar em lógica de chats/mensagens/loader — apenas apresentação.

## Fora de escopo
- Avatar composto (mosaico) para grupos — mantém o ícone de grupo atual.
- Persistência do cache entre reloads (IndexedDB) — só memória por enquanto.
