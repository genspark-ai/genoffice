# Linha de comandos e agentes

Todas as instalações incluem um comando `genoffice` que conduz os mesmos motores que a janela usa — os mesmos analisadores, o mesmo escritor, o mesmo renderizador. Um ficheiro que a aplicação guarda e um ficheiro que o comando escreve são o mesmo ficheiro, e uma verificação que o painel de IA da aplicação ultrapassa é uma verificação que o comando também ultrapassa.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

É essa a superfície completa num único ecrã — cada comando com uma linha a dizer o que faz, depois as opções globais e os códigos de saída:

![A saída real de genoffice --help: a faixa de versão, cada comando com uma descrição de uma linha, e as opções globais e os códigos de saída](img/cli.png)

## Obter o comando

macOS e Windows incluem-no dentro do pacote da aplicação. Para o usar pelo nome, execute `genoffice install-cli` uma vez: cria uma ligação simbólica para o binário incluído em `/usr/local/bin`, ou no `PATH` do seu utilizador no Windows.

## Os comandos que vale a pena conhecer

| Comando           | O que faz                                                                                                                                                                                                                                                                                             |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Abre um documento na aplicação; arranca a aplicação se ela não estiver em execução.                                                                                                                                                                                                                   |
| `selection`       | O que o utilizador tem selecionado nesse ficheiro neste momento, enquanto a aplicação está em execução — o ponteiro para «este» / «aqui». Devolve um intervalo de blocos, um intervalo de células numa folha, elementos de um diapositivo ou uma página, conforme o editor em que o ficheiro estiver. |
| `convert`         | Converte entre formatos usando os motores da própria aplicação.                                                                                                                                                                                                                                       |
| `create`          | Cria um documento a partir de conteúdo estruturado.                                                                                                                                                                                                                                                   |
| `render`          | Um PNG por página, tal como o renderizador o dispõe.                                                                                                                                                                                                                                                  |
| `pdf`             | Lê a camada de texto de um PDF página a página, sem qualquer processo da aplicação.                                                                                                                                                                                                                   |
| `info`            | Metadados e um resumo da estrutura de um documento.                                                                                                                                                                                                                                                   |
| `search`          | Pesquisa na web ou em imagens através do fornecedor configurado na aplicação.                                                                                                                                                                                                                         |
| `image` / `media` | Gera uma imagem, ou descreve e faz perguntas sobre um ficheiro de imagem, vídeo ou áudio.                                                                                                                                                                                                             |
| `merge`           | Preenche os marcadores `{{key}}` num modelo `.docx`, `.pptx` ou `.xlsx`.                                                                                                                                                                                                                              |
| `capabilities`    | Indica que funcionalidades na nuvem estão configuradas nesta máquina.                                                                                                                                                                                                                                 |
| `guide`           | A referência de operações e os guias de design, gerados a partir das mesmas definições contra as quais o executor valida — por isso não pode divergir do que `apply` aceita. `--json` devolve-a com o esquema de cada operação.                                                                       |
| `install-cli`     | Coloca o `genoffice` no `PATH`.                                                                                                                                                                                                                                                                       |
| `skill`           | Lista os agentes de programação encontrados nesta máquina e instala ou atualiza neles a competência do GenOffice.                                                                                                                                                                                     |
| `mcp`             | Serve cada comando como uma ferramenta do Model Context Protocol. Consulte **Ligar um agente de programação**.                                                                                                                                                                                        |

## Edição: documentos, folhas, apresentações

`genoffice docs`, `genoffice sheet` e `genoffice slides` leem e editam pelo mesmo caminho de escrita que a aplicação usa, e partilham um vocabulário: uma **operação** (op) é uma única edição, e uma **especificação** (spec) é uma lista de operações aplicadas por ordem.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` relata o que um lote faria e não escreve nada, que é a forma barata de verificar uma especificação antes de a aplicar. O painel de IA da aplicação funciona exatamente com estas operações, por isso tudo o que lhe pedir também pode ser escrito em script.

## Model Context Protocol

`genoffice mcp` serve cada comando como uma ferramenta MCP, e `genoffice mcp install <agent|all>` regista-a na configuração do próprio agente de programação. Consulte **Ligar um agente de programação** para esse lado.

## O que o comando não faz

Ele lê e escreve o ficheiro. Não é a aplicação: não há janela, e a janela de atualização da aplicação não se aplica. Tudo o que precisa da janela — o painel de IA, a verificação de qualidade de uma apresentação renderizada — tem de esperar que abra o ficheiro. `genoffice render` dá-lhe os píxeis sem ela, e `genoffice slides` audita a disposição de uma apresentação por si só.
