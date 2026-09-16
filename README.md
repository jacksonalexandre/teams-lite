# Teams Lite

Cliente web minimalista para o Microsoft Teams. Lê e envia mensagens de chats e canais direto pelo Microsoft Graph, sem Electron e sem o resto do cliente oficial.

O Teams oficial é um aplicativo Electron: cada janela carrega um Chromium inteiro, e o consumo de memória cresce bem além do que uma lista de conversas e um campo de texto justificam. Este projeto nasceu como uma alternativa para quem passa o dia com o Teams aberto só para acompanhar conversas — uma SPA enxuta, rodando na aba do navegador que você já tem aberta, com apenas as funções que realmente são usadas no dia a dia.

> **Projeto pessoal, não oficial.** Não tem qualquer vínculo com a Microsoft. Use por sua conta e risco.

---

## O que ele faz

**Chats**
- Lista as 50 conversas mais recentes, ordenadas pela última mensagem
- Prévia da última mensagem com o nome de quem enviou
- Título automático da conversa: assunto do grupo, ou o nome dos participantes quando não há assunto
- Ocultar e reexibir conversas (`hideForUser` do Graph) — conversas ocultas somem da lista
- Botão de chamada de vídeo em conversas individuais, via deep link para o Teams oficial

**Canais**
- Lista as equipes das quais você participa e os respectivos canais
- Mostra os 12 canais com atividade mais recente, ordenados por data da última mensagem
- Leitura e envio de mensagens nos canais

**Mensagens**
- Envio com Enter (Shift+Enter quebra linha)
- Mensagem otimista: aparece na tela imediatamente com estado de envio, e em caso de falha oferece repetir ou descartar
- Atualização automática por polling, que pausa quando a aba está em segundo plano
- Imagens inline, inclusive as hospedadas no Graph (`hostedContents`), baixadas com o token e exibidas como blob
- Anexos renderizados como cartões clicáveis, com tratamento específico para componentes do Microsoft Loop
- Separadores de data entre os dias
- Avatares reais dos contatos, com fallback para iniciais e cor derivada do nome
- Janela de 7 mensagens com botão "carregar mais", preservando a posição de rolagem
- Aviso flutuante de "novas mensagens" quando chegam mensagens e você não está no fim da lista

## O que ele não faz

Vale deixar explícito, porque a lista é grande:

- Chamadas de voz e vídeo (o botão apenas abre o Teams oficial)
- Envio de arquivos e imagens
- Reações, respostas em thread, edição e exclusão de mensagens
- Menções, formatação rica, emojis ou GIFs
- Notificações push ou notificações do navegador
- Presença e status dos contatos
- Busca
- Reuniões, calendário e aplicativos do Teams
- Tempo real por WebSocket — a atualização é por polling

---

## Stack

| Camada | Escolha |
| --- | --- |
| Framework | TanStack Start (React 19 + SSR) |
| Roteamento | TanStack Router (rotas por arquivo) |
| Build | Vite 8 |
| Estilo | Tailwind CSS 4 |
| Autenticação | MSAL Browser (`@azure/msal-browser`) |
| API | Microsoft Graph v1.0, chamada direto do navegador |
| Ícones | lucide-react |

Não há backend próprio. A única função de servidor existente entrega o `clientId` e o `tenantId` ao cliente; todo o resto são chamadas do navegador direto para o Graph, autenticadas com o token do usuário.

---

## Pré-requisitos

- Node.js 20 ou superior
- [Bun](https://bun.sh) (o repositório usa `bun.lock`; `npm` também funciona)
- Um registro de aplicativo no Microsoft Entra ID (antigo Azure AD)

### Registrando o aplicativo no Entra ID

1. Acesse o [portal do Entra ID](https://entra.microsoft.com) → **Aplicativos** → **Registros de aplicativo** → **Novo registro**.
2. Em **Tipos de conta com suporte**, escolha o que corresponde ao seu caso (normalmente "Somente contas neste diretório organizacional").
3. Em **URI de Redirecionamento**, escolha a plataforma **SPA** (Single-page application) e informe:

   ```
   http://localhost:8080/auth-callback
   ```

   O login acontece em um popup, e essa rota existe só para devolver o resultado à janela principal. Adicione também a URL de produção, se houver.
4. Copie o **ID do aplicativo (cliente)** e o **ID do diretório (locatário)**.
5. Em **Permissões de API**, adicione as seguintes permissões **delegadas** do Microsoft Graph:

   | Permissão | Para quê |
   | --- | --- |
   | `User.Read` | Perfil do usuário autenticado |
   | `User.ReadBasic.All` | Nomes e fotos dos participantes |
   | `Chat.ReadWrite` | Ler chats, ocultar e reexibir conversas |
   | `ChatMessage.Send` | Enviar mensagens em chats |
   | `Team.ReadBasic.All` | Listar equipes |
   | `Channel.ReadBasic.All` | Listar canais |
   | `ChannelMessage.Read.All` | Ler mensagens de canais |
   | `ChannelMessage.Send` | Enviar mensagens em canais |

   `ChannelMessage.Read.All` exige consentimento de um administrador do locatário. Sem ele, a aba de chats funciona normalmente e a de canais falha ao carregar as mensagens.

---

## Rodando localmente

```bash
git clone https://github.com/jacksonalexandre/teams-lite.git
cd teams-lite
bun install
```

Crie um arquivo `.env` na raiz:

```env
TEAMS_CLIENT_ID=00000000-0000-0000-0000-000000000000
TEAMS_TENANT_ID=00000000-0000-0000-0000-000000000000
```

`TEAMS_TENANT_ID` é opcional — quando ausente, o padrão é `common`, que aceita contas de qualquer locatário.

```bash
bun run dev
```

A aplicação sobe em `http://localhost:8080`. Clique em **Entrar com Microsoft**, conclua o login no popup e as conversas aparecem.

### Scripts

| Comando | O que faz |
| --- | --- |
| `bun run dev` | Servidor de desenvolvimento com HMR |
| `bun run build` | Build de produção (cliente + SSR) em `dist/` |
| `bun run build:dev` | Build com as configurações de desenvolvimento |
| `bun run preview` | Serve o build de produção localmente |
| `bun run lint` | ESLint em todo o projeto |

---

## Estrutura

```
src/
├── routes/
│   ├── __root.tsx          Shell HTML, providers, páginas de erro e 404
│   ├── index.tsx           Tela principal: lista lateral, mensagens e composer
│   └── auth-callback.tsx   Popup de retorno do login MSAL
├── hooks/
│   ├── useChats.ts         Lista de chats, polling e ocultar/reexibir
│   ├── useChannels.ts      Equipes, canais e ordenação por atividade
│   └── useMessages.ts      Mensagens, envio otimista e rolagem
├── components/
│   ├── Avatar.tsx          Foto do contato, com fallback para iniciais
│   ├── ChatItem.tsx        Item da lista de chats
│   ├── ChannelItem.tsx     Item da lista de canais
│   ├── MessageBubble.tsx   Balão, imagens inline e cartões de anexo
│   └── DateSeparator.tsx   Divisória de data
├── lib/
│   ├── msal.ts             Configuração do MSAL e aquisição de token
│   ├── graph.ts            Cliente do Microsoft Graph
│   ├── config.functions.ts Função de servidor que entrega clientId/tenantId
│   ├── format.ts           Formatação de datas e limpeza de HTML
│   ├── error-capture.ts    Captura de erros de SSR fora de banda
│   └── error-page.ts       Página de erro estática
├── server.ts               Entrada SSR com tratamento de erro
├── start.ts                Middleware de requisição do TanStack Start
└── styles.css              Design system em Tailwind 4 (tokens oklch)
```

---

## Como funciona

**Autenticação.** O `clientId` e o `tenantId` vêm do servidor por uma server function, para que não fiquem escritos no código do cliente. O MSAL roda no navegador e guarda o cache em `localStorage`, então a sessão sobrevive a recarregamentos. O login usa popup: a janela abre, autentica e a rota `/auth-callback` devolve o resultado à janela principal. Os tokens são renovados silenciosamente e só voltam a pedir interação quando o Graph exige.

**Dados.** Toda chamada ao Graph parte do navegador, com o token do próprio usuário no cabeçalho. Nenhuma mensagem passa por um servidor intermediário.

**Atualização.** Não há WebSocket. A lista de chats é recarregada a cada 15 segundos e as mensagens da conversa aberta a cada 4 segundos. Os dois timers verificam `document.hidden` e pulam a requisição enquanto a aba está em segundo plano, o que evita gastar cota do Graph com uma janela que ninguém está olhando.

**Imagens e avatares.** Fotos de perfil e imagens hospedadas no Graph precisam do cabeçalho de autorização, então não dá para usá-las direto no `src` da tag. Elas são baixadas como blob, convertidas em object URL e guardadas em um `Map` em memória, para que cada imagem só seja buscada uma vez por sessão.

**Erros de SSR.** O h3 engole exceções lançadas durante a renderização no servidor e devolve um JSON genérico de 500. O `src/server.ts` detecta esse formato, recupera o erro original capturado por `error-capture.ts` e responde com uma página de erro legível em vez do JSON.

---

## Limitações conhecidas

- As mensagens de canais dependem de consentimento administrativo, o que pode inviabilizar o uso em locatários corporativos mais restritos.
- O polling tem custo de cota no Graph. Com muitas abas abertas é possível bater em throttling (HTTP 429), que hoje aparece como erro na faixa superior da tela.
- Apenas as 50 mensagens mais recentes de cada conversa são carregadas; não há paginação para além disso.
- O envio é sempre em texto puro — HTML e Markdown digitados aparecem literalmente.
- O design system tem tokens de tema escuro definidos em `styles.css`, mas ainda não existe um seletor para alterná-lo.

## Contribuindo

É um projeto pessoal, tocado no tempo livre. Issues e pull requests são bem-vindos, mas a resposta pode demorar. Se for mexer em algo grande, abra uma issue antes para combinarmos a direção.

Antes de enviar um PR, rode:

```bash
bun run lint
bunx tsc --noEmit
bun run build
```

## Licença

[MIT](LICENSE). Use, modifique e distribua à vontade, mantendo o aviso de copyright.
