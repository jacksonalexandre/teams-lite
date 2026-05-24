## Contexto
Atualmente no `MessageBubble`:
- O nome do autor aparece na parte superior interna da bolha (quando não é minha mensagem).
- A data/hora aparece na parte inferior da bolha.

## Alteração
Reorganizar o cabeçalho da mensagem dentro da bolha:
1. **Para mensagens dos outros (alheias):** exibir `Nome do Autor · HH:MM` no topo, em linha única.
2. **Para mensagens do usuário logado:** exibir só o `HH:MM` no topo, sem o nome.
3. Remover o timestamp do rodapé da bolha, evitando duplicação.
4. Ajustar espaçamentos (`mb-*`) para que o novo cabeçalho compacto fique alinhado com o conteúdo.

## Escopo
Apenas o componente `MessageBubble` em `src/routes/index.tsx`. Nenhuma lógica de dados ou API muda.

## Resultado esperado
- Mensagens alheias: cabeçalho no topo com "João Silva · 14:32", conteúdo abaixo.
- Mensagens próprias: cabeçalho no topo apenas com "14:32", conteúdo abaixo.