## Mudanças em `src/routes/index.tsx`

### 1. Sidebar de chats
- Mover o botão "Carregar mais" do topo da lista para **depois** dos itens renderizados (final da lista).
- Manter `visibleCount` inicial = 7 e incremento de +7.

### 2. Mensagens do chat aberto
- Adicionar estado `messagesVisibleCount` (`useState(7)`).
- Resetar para 7 sempre que o chat selecionado mudar (`useEffect` dependente do chat ativo).
- Renderizar apenas as últimas 7 mensagens: `messages.slice(-messagesVisibleCount)` (preservando ordem cronológica).
- Adicionar botão "Carregar mais (N restantes)" no **início** da lista de mensagens (acima da primeira mensagem visível), exibido somente quando `messages.length > messagesVisibleCount`. Ao clicar, incrementa `messagesVisibleCount` em 7.
- Manter o scroll automático para o fim apenas no carregamento inicial / nova mensagem, não ao clicar em "Carregar mais" (preservar posição de leitura quando o usuário sobe o histórico).
