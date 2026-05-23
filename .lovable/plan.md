## Diagnóstico

### 1. Data incorreta na conversa
Em `src/routes/index.tsx` o sidebar exibe `c.lastUpdatedDateTime`, mas esse campo do Graph reflete metadados do chat (criação/renomeação/membros), **não a última mensagem**. Tanto que o `listChats` em `src/lib/graph.ts` já ordena por `lastMessagePreview/createdDateTime desc` — ou seja, a ordem está certa, mas a data exibida vem de outro campo. Por isso aparece uma data antiga/estranha para o chat com o Shin.

### 2. Nome de usuário externo não aparece
Em `chatTitle` (`src/lib/graph.ts`) o título vem de `members[].displayName`. Para convidados/usuários externos (federados, B2B guest), o Graph frequentemente devolve:
- `displayName: null`
- `userId: null` (ou um id de tenant externo)
- a identidade real vem em `email` ou `tenantId` diferente

Como o filtro é `m.userId !== meId` e o `displayName` é `null`, o `names` fica vazio e cai em `"(sem título)"` — ou, se houver mais membros, some o nome do externo.

Além disso, mesmo `$expand=members` às vezes não traz o `displayName` de externos. O fallback robusto é usar `lastMessagePreview.from.user.displayName` quando os membros não ajudam.

## Plano

### A. Trazer dados melhores do Graph (`src/lib/graph.ts`)
1. Atualizar `listChats` para também pedir `lastMessagePreview`:
   ```
   /me/chats?$expand=members&$select=id,topic,chatType,lastUpdatedDateTime,lastMessagePreview&$orderby=lastMessagePreview/createdDateTime desc&$top=50
   ```
2. Estender o tipo `GraphChat`:
   - adicionar `email?: string` em `members[]`
   - adicionar `lastMessagePreview?: { createdDateTime?: string; from?: { user?: { displayName?: string; id?: string } } | null; body?: { content?: string; contentType?: string } }`

### B. Corrigir título com fallback para externos (`chatTitle` em `src/lib/graph.ts`)
Nova ordem de fallback quando `topic` é vazio:
1. `members` com `userId !== meId` e `displayName` presente → comportamento atual.
2. Se a lista resultante ficar vazia (caso típico de externo), usar:
   - `lastMessagePreview.from.user.displayName` (se diferente do meu nome), ou
   - `members[*].email` do outro lado, ou
   - `"Usuário externo"` como último recurso.

### C. Corrigir a data exibida (`src/routes/index.tsx`)
Trocar `formatDate(c.lastUpdatedDateTime)` por `formatDate(c.lastMessagePreview?.createdDateTime ?? c.lastUpdatedDateTime)` — assim o horário mostrado coincide com a última mensagem real (e com a ordenação).

## Detalhes técnicos
- Sem mudanças de auth/escopos: `Chat.ReadBasic`/`Chat.Read` já cobrem `lastMessagePreview` e `members`.
- Mudanças isoladas em 2 arquivos (`src/lib/graph.ts` e `src/routes/index.tsx`). Sem impacto em mensagens, envio ou polling.
- `MessageBubble` continua comparando por `displayName` — não muda nada.
