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

## Exportação

Menu Arquivo, tudo local e tudo pergunta onde colocar o resultado:

- **Exportar como Word…** e **Exportar como PDF…** escrevem um .docx ou .pdf a sério.
- **Exportar como imagens…** escreve um PNG por página numa pasta que escolher.
- **Converter e abrir no Docs** converte para .docx e abre-o no separador Docs incorporado aqui na aplicação — não é uma entrega a nada na nuvem, e a cópia convertida fica numa pasta de cache que é limpa ao fim de cerca de uma semana.

## Vista de código-fonte

O friso traz um interruptor **Código-fonte** (localizado com a aplicação). Ligue-o e o editor é substituído pelo Markdown em bruto: exatamente o texto que uma gravação escreve, nada embelezado, nada normalizado por baixo de si.

- **A edição é fiel byte a byte.** Uma gravação a partir da vista de código-fonte produz exatamente os mesmos bytes que uma gravação a partir do editor — o BOM, o CRLF e a presença de uma quebra de linha no fim sobrevivem todos.
- **É o mesmo documento.** Alterne para um lado e para o outro à vontade; a fonte é o próprio texto do editor, não uma cópia que seja preciso fundir.
- **A barra de formatação fica indisponível** enquanto a vista está aberta, porque a maioria desses botões insere construções do editor que só têm significado do lado renderizado. Reaparece quando fecha a vista.
- **Ficheiros JSON e outros ficheiros em modo de código-fonte** abrem aqui diretamente: não há nada para renderizar, por isso a fonte _é_ o documento.
