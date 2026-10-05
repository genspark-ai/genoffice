# O editor Markdown

O editor Markdown abre .md / .markdown com a experiência de fonte + pré-visualização renderizada.

- **Abrir**: a partir do Início ou de Arquivo ▸ Abrir…; também funciona a partir da linha de comandos.
- **Editar**: edição de texto simples; as extensões GFM (tabelas, listas de tarefas, rasurado, ligações automáticas) são renderizadas na pré-visualização.
- **Pré-visualização**: em tempo real; os recursos relativos, como imagens, são resolvidos ao lado do documento.
- **Salvar**: fiel byte a byte — o BOM, o CRLF e a presença de uma quebra de linha no fim do ficheiro são preservados; um salvamento sem alterações não reescreve o ficheiro.
- **Pesquisar e substituir**: o ctrl+F pesquisa no código; Substituir grava o resultado de volta.
- **IA**: botões prontos deixam o assistente reescrever, desenvolver ou traduzir o documento.

## A barra de ferramentas

Uma fila de botões por cima do editor (passe o cursor para ver as dicas):

![A barra de ferramentas do Markdown](img/md-toolbar.png)

- **Ficheiro e histórico**: Salvar, Salvar Como…, Desfazer, Refazer, Localizar; o interruptor de **Salvamento automático** à direita grava as alterações periodicamente no disco.
- **Botão de IA**: abre o painel de IA; ao lado ficam as opções de reescrever / desenvolver / traduzir.
- **Estilo de parágrafo** (menu suspenso): alterna entre o texto do corpo e os vários níveis de título.
- **Formatação na linha**: **negrito**, _itálico_, ~~tachado~~, `código embutido`, ligação.
- **Listas**: lista com marcadores, lista numerada, lista de tarefas.
- **Inserir**: tabela, imagem, separador horizontal.
- **Propriedades**: insere ou salta para o bloco YAML front matter no topo do ficheiro.
- **Estrutura**: salta pela hierarquia de títulos.
- **Ortografia**: liga ou desliga a verificação ortográfica deste documento.

Três exemplos rápidos:

- **Título**: ponha o cursor nessa linha ▸ menu suspenso de estilo de parágrafo ▸ "Título 1".
- **Tabela**: clique **Inserir tabela** ▸ arraste para escolher o número de linhas e colunas ▸ escreva nas células; a pré-visualização renderiza-a de imediato.
- **Lista de tarefas**: selecione algumas linhas ▸ clique **Lista de tarefas** ▸ cada linha passa a `- [ ]` e aparece na pré-visualização como uma caixa de verificação.
