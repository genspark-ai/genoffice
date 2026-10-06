# Ligar um agente de programação

O GenOffice fala o Model Context Protocol, por isso um agente de programação consegue ler, escrever e renderizar os seus documentos através dos mesmos motores que a aplicação usa. O agente não está a adivinhar um formato de ficheiro: recebe os esquemas tipados das operações a partir das mesmas definições contra as quais o executor valida.

## Registá-lo

O caso habitual é um comando:

```sh
genoffice mcp install all
```

Ele encontra os agentes de programação desta máquina — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — e escreve a entrada do servidor stdio na configuração de cada um deles, deixando o resto desse ficheiro como o encontrou.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Um agente instalado num sítio fora do comum precisa de `--dir <path>`; e `--force` reescreve uma entrada que já lá está.

## Executá-lo você mesmo

Para um cliente noutra máquina, sirva-o por HTTP em vez disso:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` muda o endereço onde escuta. Passe o mesmo token ao cliente.

Por HTTP os ficheiros também viajam: `PUT /files/<name>` envia um, todas as ferramentas aceitam um URL `http(s)` no lugar de um caminho, e os resultados voltam como URLs de transferência — e, quando são pequenos o suficiente, como recursos incorporados. As operações, especificações e Markdown passam sempre embutidos.

## O que o agente recebe

Cada comando é uma ferramenta. As interessantes:

- **`docs`, `sheet`, `slides`** — leem e editam um ficheiro pelo caminho de escrita da própria aplicação, uma **operação** (op) de cada vez. Uma apresentação nova segue `deck_start`, `deck_page`, `deck_build`.
- **`render`** — um PNG por página, disposto pelo renderizador da aplicação, para que o agente possa olhar para um slide em vez de adivinhar.
- **`pdf`** — a camada de texto de um PDF, página a página, sem qualquer processo da aplicação.
- **`info`** — metadados e um resumo da estrutura, normalmente a primeira chamada mais barata para um ficheiro desconhecido.
- **`search`, `image`, `media`** — os fornecedores configurados na aplicação, por isso o agente não precisa das suas próprias chaves.
- **`merge`** — preencher um modelo `{{key}}`.

## Esquemas, e um orçamento mais pequeno

`apply` e `create` anunciam os seus parâmetros `ops`, `cells` e `data` com o esquema tipado de cada operação, gerado a partir de `genoffice guide <domain> --json`. É preciso, mas não é pequeno. Um cliente com uma janela de contexto apertada pode pedir matrizes simples em vez disso:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## A competência

Um agente que não conhece o vocabulário das operações vai adivinhar. O `genoffice skill` instala uma competência do GenOffice nos agentes que encontra, levando a referência e os guias de design — o mesmo material que o `genoffice guide` imprime.

## O que não é

O servidor MCP lê e escreve ficheiros. Não é a janela: não existe painel de IA, e a janela de atualização da aplicação não se aplica. Se algum passo precisar da janela, abra o ficheiro.