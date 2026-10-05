# O editor HTML

O editor HTML abre .html / .htm com dois modos: **Visualizar** (a página renderizada) e **Código**.

- **Visualizar**: renderização a sério; as folhas de estilo e imagens relativas são carregadas a partir de junto do ficheiro.
- **Inspetor de visualização**: clique para selecionar um elemento, clique duas vezes para editar o respetivo texto no local, elimine pela barra de ferramentas e pergunte à IA sobre a seleção.
- **Modo de código**: edita o HTML; ctrl+F procura e Substituir guarda a marcação reescrita.
- **Salvar**: fiel byte a byte (BOM/CRLF/quebra de linha final preservados); um salvamento sem alterações não reescreve.
- **Zoom**: ctrl+roda/pinça dimensiona a pré-visualização; ctrl+Z dentro da pré-visualização desfaz a última edição.

## A barra de ferramentas

Clique num elemento da pré-visualização e uma barra de ferramentas flutua por cima dele:

![A barra de ferramentas flutuante sobre um elemento selecionado](img/html-toolbar.png)

- **Ficheiro e histórico**: Salvar, Salvar Como…, Desfazer, Refazer, Localizar; o interruptor de **Salvamento automático** grava as alterações periodicamente.
- Interruptor **Visualizar / Código**; **Apresentar** mostra a página em ecrã inteiro.
- **Formatação**: negrito, itálico, aumentar/diminuir o tamanho da fonte; e o **painel de estilo** do elemento selecionado (cores e mais).
- **Inserir**: título, parágrafo, tabela, imagem (por ligação), Mais.
- **Ações de imagem** (com uma imagem selecionada): cortar, **remover o plano de fundo**, substituir, bloquear a proporção.
- **Ações de elemento** (com um elemento selecionado no inspetor de visualização): eliminar, duplicar, mover para cima/baixo.
- **Botão de IA**: abre o painel de IA; pergunte diretamente sobre o elemento selecionado.
