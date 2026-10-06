# Início rápido: a interface e o básico

O GenOffice é um pacote de escritório que funciona inteiramente na sua máquina: uma janela, uma faixa de guias, com seis editores — Docs (processamento de texto), Sheets (folhas de cálculo), Slides (apresentações), PDF, Markdown e HTML. Os ficheiros são .docx / .xlsx / .pptx / .pdf a sério, totalmente compatíveis com o Word, o Excel e o PowerPoint. Não é necessária rede.

## A interface em resumo

![O ecrã Início](img/home-screen.png)

A janela tem três partes:

- **Barra de guias (no topo)**: cada ficheiro aberto é uma guia. A guia Início mais à esquerda está sempre presente e não pode ser fechada; as restantes guias são os seus documentos. Faça clique duplo numa guia para renomear o ficheiro no local.
- **Área de conteúdo**: o editor (ou o ecrã Início) da guia ativa.
- **Barra de menus**: na barra de menus do sistema no macOS, no topo da janela no Windows/Linux. Os menus Ficheiro/Editar/Exibir mudam para corresponder ao editor ativo.

## Criar um documento

De qualquer uma destas formas:

- Clique num cartão de criação rápida na secção [Início rápido](help://getting-started) do **Início** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **Arquivo ▸ Novo**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML ou PDF.
- Arraste um ficheiro para a janela, ou clique duas vezes nele no gestor de ficheiros (se o GenOffice for a aplicação predefinida).

O documento novo abre sem título; o ficheiro no disco só é criado no primeiro salvamento.

## Abrir ficheiros

- O menu **Arquivo ▸ Abrir…** (⌘O/ctrl+O) abre o seletor de ficheiros do sistema: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Clique em qualquer item da lista **Recentes** no ecrã Início.
- `genoffice <ficheiro>` a partir de um terminal também abre ficheiros.

## Como funciona o salvamento

- **Salvamento manual**: ⌘S/ctrl+S, ou Arquivo ▸ Salvar / Salvar Como. O primeiro salvamento pede a localização e o nome.
- O **Salvamento automático** só é ativado depois de ter salvo o ficheiro manualmente pelo menos uma vez — um PDF que se limita a ler nunca é reescrito em silêncio. O salvamento automático dispara pouco depois de o conteúdo mudar.
- Fechar uma guia com alterações não salvas pergunta primeiro Salvar / Descartar / Cancelar.
- Cada escrita é atómica (ficheiro temporário + mudança de nome), pelo que uma falha de energia não deixa um ficheiro a meio.

## Atalhos frequentes

| Ação                       | macOS | Windows / Linux |
| -------------------------- | ----- | --------------- |
| Novo documento             | ⌘N    | ctrl+N          |
| Abrir                      | ⌘O    | ctrl+O          |
| Salvar                     | ⌘S    | ctrl+S          |
| Fechar guia                | ⌘W    | ctrl+W          |
| Abrir este manual          | F1    | F1              |
| Recolher a Faixa de Opções | ⌥⌘R   | Ctrl+F1         |

**Recolher a Faixa de Opções** funciona em todos os editores. A linha de guias mantém-se e a faixa de comandos por baixo desaparece; a guia selecionada serve também de controlo de recolha, pelo que enquanto a faixa estiver recolhida não há nenhuma guia selecionada e premir qualquer uma das guias traz a faixa de volta. Um duplo clique numa guia faz o mesmo. O sentido em que a deixou é guardado editor a editor.

Os atalhos dentro de cada editor (pincel de formatação, pesquisar e substituir, operações de tabela, ...) estão nos respetivos capítulos; o Docs traz ainda uma janela de atalhos de teclado pesquisável (**⌘/**) (consulte o seu capítulo).

## Os atalhos Option+Command

Option+Command é a camada que o Word reserva para saltos estruturados, e o GenOffice preenche-a da mesma forma. Tudo o que se segue é próprio do Docs:

| Atalho (macOS)  | O que faz              | Windows / Linux    |
| --------------- | ---------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3 | Título 1 / 2 / 3       | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0             | Normal                 | Ctrl+Alt+0         |
| ⌥⌘M             | Parágrafo              | Ctrl+Alt+M         |
| ⌥⌘A             | Novo Comentário        | Ctrl+Alt+A         |
| ⌥⌘F             | Inserir Nota de Rodapé | Ctrl+Alt+F         |
| ⌥⌘E             | Inserir Nota de Fim    | Ctrl+Alt+D         |
| ⌥⌘G             | Ir para                | Ctrl+G             |

Dois deles mudam no Windows, pela mesma razão que leva o Word a separá-los. **O macOS tem o ⌥⌘D** — mostra e oculta o Dock —, por isso a nota de fim é ⌥⌘E no Mac e Ctrl+Alt+D em qualquer outro lado. E **Ir para** dispensa o Alt: Ctrl+G, ao passo que o atalho do Mac o traz consigo.

Isso deixa o ⌥⌘D livre para o GenOffice usar no macOS, se um comando futuro precisar dele.

## Para onde ir a seguir

- Onde ficam os ficheiros: [O ecrã Início](help://home-screen).
- Gerir muitos ficheiros abertos: [Guias e gestão de janelas](help://tabs-and-windows).
- Deixar o trabalho à IA: [O painel do assistente de IA](help://ai-panel).
- Idioma, tema, aplicações predefinidas: [Configurações, idioma, tema e integrações MCP](help://settings-integrations).
