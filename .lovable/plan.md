## Por que as imagens não aparecem

Imagens coladas/enviadas no Teams vêm no `body.content` (HTML) como:

```html
<img src="https://graph.microsoft.com/v1.0/chats/{chatId}/messages/{msgId}/hostedContents/{cid}/$value"
     itemid="..." width="..." height="..." />
```

Dois problemas hoje em `MessageBubble`:

1. `stripHtml(stripAttachmentTags(...))` apaga toda tag HTML — `<img>` some junto.
2. Mesmo se renderizássemos `<img src="...">` direto, o navegador faria GET **sem** o header `Authorization: Bearer ...` que a Graph exige → 401. URLs de `hostedContents/$value` **só** funcionam com token.

## Solução

Buscar cada `hostedContent` via `fetch` autenticado (com `getAccessToken`), converter para blob e usar `URL.createObjectURL()` no `<img src>`. Cachear por id da imagem.

### Mudanças

1. **src/lib/graph.ts**
   - Adicionar `getHostedContentUrl(cfg, url)`: recebe a URL da Graph (`…/hostedContents/{id}/$value`), faz `fetch` com Bearer, retorna `URL.createObjectURL(blob)`. Cache em `Map<string, Promise<string|null>>` análogo ao `photoCache` já existente.

2. **src/routes/index.tsx — `MessageBubble`**
   - Em vez de `stripHtml`, parsear o HTML do body com `DOMParser` quando `contentType === "html"`:
     - Extrair `<img>` cujo `src` começa com `https://graph.microsoft.com/.../hostedContents/`.
     - Remover essas tags do DOM e usar o `textContent` resultante como texto.
     - Renderizar as imagens abaixo do texto via um novo componente `<HostedImage url=... cfg=... />` que faz `useEffect` chamando `getHostedContentUrl`, guarda em state e exibe `<img src={blobUrl}>` (com placeholder/loader enquanto carrega).
   - Imagens externas (`src` que não é do Graph, ex.: gifs públicos, emojis customizados) podem ir direto como `<img src={…}>` sem auth.
   - Relaxar o guard `if (!text.trim() && attachments.length === 0 && images.length === 0) return null;` para não descartar mensagens só com imagem.

3. **Preview na sidebar** (`listChats` em `index.tsx:451`): quando `preview.body.content` só tem `<img>`, hoje fica vazio. Mostrar fallback `"📷 Imagem"` quando o texto após `stripHtml` é vazio mas o HTML continha `<img>`.

### Detalhes técnicos

- O fetch retorna `image/png`, `image/jpeg`, `image/gif` (animado preserva). Apenas usar `res.blob()`.
- Cache key = URL completa (inclui chatId + msgId + contentId — único e estável).
- Liberar `URL.revokeObjectURL` no unmount não é estritamente necessário porque o cache mantém o blob vivo durante a sessão; aceitável dado o volume.
- Tamanho: respeitar `width`/`height` originais do `<img>` quando presentes, com `max-width: 320px` para não estourar a bolha.
- Clique na imagem → abrir em nova aba (também via blob URL).

## Fora de escopo

- Lightbox/galeria.
- Upload/envio de imagem pelo nosso app.
- Imagens dentro de cartões adaptativos complexos.
