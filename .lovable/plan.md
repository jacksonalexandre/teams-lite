## Contexto

Hoje, `MessageBubble` (src/routes/index.tsx:677) faz `stripHtml(m.body.content)` e ignora completamente `attachments[]`. Componentes do Loop chegam exatamente como `<attachment id="…"/>` no corpo HTML + um item em `message.attachments[]` — por isso eles "somem" na UI.

## O que o Microsoft Graph entrega (e o que NÃO entrega)

Um Loop component em uma mensagem do Teams aparece assim:

```json
{
  "body": { "contentType": "html", "content": "<attachment id=\"abc123\"></attachment>" },
  "attachments": [{
    "id": "abc123",
    "contentType": "application/vnd.microsoft.card.fluidEmbedCard",
    "contentUrl": "https://contoso-my.sharepoint.com/.../FluidPreview.aspx?...",
    "name": "Lista de tarefas",
    "content": "{...metadados JSON...}"
  }]
}
```

**A Graph NÃO retorna o conteúdo renderizado** (tabela, lista, tarefas etc.). O Loop é um container vivo do **Fluid Framework** hospedado no OneDrive/SharePoint do usuário. Renderizar de verdade exige:

1. SDK do Fluid Framework + container do Loop (`@fluidframework/azure-client` + `@microsoft/loop-*`), que **não é público** fora dos hosts oficiais (Teams, Outlook, Loop app, Office).
2. Permissões adicionais (`Files.Read.All`, `Sites.Read.All`) para acessar o arquivo `.fluid` no OneDrive do remetente.
3. Acesso ao serviço Azure Fluid Relay autenticado — bloqueado para apps de terceiros.

**Conclusão:** renderizar o Loop *embutido como no Teams* não é viável num app externo. É uma limitação de plataforma, não algo que faltou implementar.

## O que dá para fazer (recomendado)

Renderizar um **"cartão de Loop"** no lugar do `<attachment>`: ícone, nome do componente, tipo (lista/tabela/tarefas/parágrafo) e botão **"Abrir no Teams/Office"** apontando para o `contentUrl`. É exatamente o que o Outlook Web faz quando não consegue carregar o componente ao vivo.

### Mudanças

1. **src/lib/graph.ts** — estender `GraphMessage`:
   ```ts
   attachments?: Array<{
     id: string;
     contentType: string;
     contentUrl?: string;
     name?: string;
     content?: string;
   }>;
   ```
   E adicionar `$expand=...` ou garantir que `listMessages` retorne `attachments` (a Graph já inclui por padrão em `/chats/{id}/messages`).

2. **src/routes/index.tsx — `MessageBubble`**:
   - Parsear o HTML do body procurando `<attachment id="…"/>` e casar com `m.attachments`.
   - Renderizar o texto restante normalmente (`stripHtml`).
   - Para cada attachment Loop (`contentType` começando com `application/vnd.microsoft.card.fluidEmbedCard` ou `loop*`), renderizar um chip:
     ```
     [🔗 ícone Loop]  Nome do componente
                      Lista · Abrir no Microsoft Teams ↗
     ```
     Link abre `contentUrl` em nova aba (o usuário cai no Office/Teams autenticado).
   - Tratar também outros attachments comuns que hoje somem: `reference` (link de arquivo OneDrive), `messageReference` (citações/reply).

3. **Não-mudança**: nenhuma mudança em `msal.ts` / scopes — `Chat.Read` já basta para os metadados.

### Detalhes técnicos

- `contentType`s relevantes para Loop: `application/vnd.microsoft.card.fluidEmbedCard`, `application/vnd.microsoft.card.loopcomponent` (varia conforme a idade da mensagem).
- O campo `attachment.content` (string JSON) traz `componentType` (`fluidlist`, `fluidtable`, `fluidtask`, `fluidparagraph`…) — usar para escolher rótulo/ícone.
- Se o body tem **só** `<attachment>` e nada de texto, hoje a mensagem é totalmente filtrada por `if (!text.trim()) return null;` (linha 681). Precisa relaxar essa condição quando há attachments renderizáveis.

## Fora de escopo (deixar claro ao usuário)

- Edição ao vivo do Loop dentro do app.
- Renderização do conteúdo atual da lista/tabela.
- Sincronização em tempo real do componente.

Ambos exigiriam o Fluid Framework + Loop SDK proprietários da Microsoft.
