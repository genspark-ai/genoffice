# Ligar um agente de programação

O GenOffice fala o Model Context Protocol, por isso um agente de programação consegue ler, escrever e renderizar os seus documentos através dos mesmos motores que a aplicação usa. O agente não está a adivinhar um formato de ficheiro: recebe os esquemas tipados das operações a partir das mesmas definições contra as quais o executor valida.

## Registá-lo a partir da aplicação

É em **Configurações ▸ Integrações** que isto se faz. O painel tem duas metades, e pode usar uma ou as duas.

**A skill.** Uma linha por cada agente de programação encontrado nesta máquina — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, cada uma com **Instalar**, **Atualizar** e **Desinstalar**, a que se juntam **Instalar em outra pasta…**, **Baixar skill (zip)** e **Copiar caminho**. Se o seu assistente não estiver na lista, indique ao GenOffice uma pasta de onde ele lê o `SKILL.md`, ou grave o zip e deixe o assistente instalá-lo. A skill e o MCP podem andar lado a lado: o assistente escolhe um dos dois, e fazem exatamente as mesmas coisas.

**MCP.** São oferecidas duas vias: **Iniciado pelo assistente (recomendado)**, em que acrescenta a configuração apresentada ao seu cliente e o assistente arranca o servidor sozinho, e o **Servidor HTTP local**, que a aplicação executa por si. De qualquer forma, o assistente acaba a falar com o GenOffice e nunca escreve um comando.

## Registá-lo a partir da linha de comando

A mesma coisa a partir de um terminal — esta é a via avançada, e a que deve usar quando o agente está num sítio onde o painel o não consegue encontrar:

```sh
genoffice mcp install all
```

Encontra os agentes de programação desta máquina e escreve a entrada do servidor stdio na configuração de cada um deles, deixando o resto desse ficheiro como o encontrou.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Um agente instalado num sítio fora do comum precisa de `--dir <path>`; e `--force` reescreve uma entrada que já lá está.

Tudo aquilo que o servidor aceita cabe num ecrã — as formas de instalar, remover e listar, a publicação por HTTP e as duas opções de esquema:

![A saída real de genoffice mcp --help: as formas install, uninstall e list, com as opções --http, --host, --token, --compact-schemas, --dir e --force](img/mcp.png)

## Executá-lo sem um assistente

Para um cliente noutra máquina, sirva-o por HTTP:

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

## Porque é que a skill

Um agente que não conhece o vocabulário das operações vai adivinhar. A skill traz a referência e os guias de design — o mesmo material que o `genoffice guide` imprime —, de modo que o assistente escreve operações cuja especificação leu mesmo. Instale-a a partir do painel acima, ou com `genoffice skill` a partir de um terminal.

## O que consegue alcançar na aplicação

O servidor não se limita aos ficheiros em disco. Enquanto o GenOffice estiver a correr, o agente também pode trabalhar através da janela:

- **`open_in_genoffice`** abre um ficheiro num separador e dá-lhe o foco.
- **`open_documents`** lista todos os documentos que tem abertos — id, tipo, caminho e se têm alterações por guardar —, e depois lê o conteúdo atual de um deles ou fecha-o, guardando primeiro, a não ser que lhe peça para descartar.
- **As ferramentas de conteúdo** recebem esse id (ou o caminho) como argumento `document`, pelo que uma alteração aterra no separador que já tem aberto e a janela muda para o mostrar.

Ficam fora do alcance duas coisas: não existe painel de IA, e a janela de atualização da aplicação não se aplica.

## O painel Servidor HTTP local

Em **Servidor HTTP local**, a aplicação executa o servidor por si própria em vez de o deixar ao assistente: um interruptor de ativação, um campo **Porta**, um indicador **Em execução / Parado** e a **Geração em segundo plano** (escrever os documentos diretamente num caminho sem abrir a interface), mais o **Exemplo de configuração do cliente** para copiar. Abrir **Avançado** acrescenta os dois URLs de ligação — Streamable HTTP e o mais antigo, de SSE —, um URL de **Verificação de integridade** e um interruptor de **Registo** que anota a atividade do servidor e das ferramentas num ficheiro local que pode **Abrir**, **Atualizar** ou **Limpar** a partir daí. Escuta apenas em localhost.
