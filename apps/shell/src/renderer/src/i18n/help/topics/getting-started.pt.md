# Início rápido: a interface e o básico

O GenOffice é um pacote de escritório que funciona inteiramente na sua máquina: uma janela, uma faixa de guias, com seis editores — Docs (processamento de texto), Sheets (planilhas), Slides (apresentações), PDF, Markdown e HTML. Os arquivos são .docx / .xlsx / .pptx / .pdf de verdade, totalmente compatíveis com o Word, o Excel e o PowerPoint. Não é necessária rede.

## A interface em resumo

![A tela Início](img/home-screen.png)

A janela tem três partes:

- **Barra de guias (no topo)**: cada arquivo aberto é uma guia. A guia Início mais à esquerda está sempre presente e não pode ser fechada; as demais guias são os seus documentos. Faça clique duplo em uma guia para renomear o arquivo no local.
- **Área de conteúdo**: o editor (ou a tela Início) da guia ativa.
- **Barra de menus**: na barra de menus do sistema no macOS, no topo da janela no Windows/Linux. Os menus Arquivo/Editar/Exibir mudam para corresponder ao editor ativo.

## Criar um documento

De qualquer uma destas formas:

- Clique em um cartão de criação rápida na seção [Início rápido](help://getting-started) do **Início** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **Arquivo ▸ Novo**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML ou PDF.
- Arraste um arquivo para a janela, ou clique duas vezes nele no gerenciador de arquivos (se o GenOffice for o aplicativo padrão).

O documento novo abre sem título; o arquivo no disco só é criado no primeiro salvamento.

## Abrir arquivos

- O menu **Arquivo ▸ Abrir…** (⌘O/ctrl+O) abre o seletor de arquivos do sistema: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Clique em qualquer item da lista **Recentes** na tela Início.
- `genoffice <ficheiro>` a partir de um terminal também abre arquivos.

## Como funciona o salvamento

- **Salvamento manual**: ⌘S/ctrl+S, ou Arquivo ▸ Salvar / Salvar Como. O primeiro salvamento pede a localização e o nome.
- O **Salvamento automático** só é ativado depois de ter salvo o arquivo manualmente pelo menos uma vez — um PDF que se limita a ler nunca é reescrito em silêncio. O salvamento automático dispara pouco depois de o conteúdo mudar.
- Fechar uma guia com alterações não salvas pergunta primeiro Salvar / Descartar / Cancelar.
- Cada escrita é atômica (arquivo temporário + mudança de nome), por isso uma falha de energia não deixa um arquivo a meio.

## Atalhos frequentes

| Ação                       | macOS | Windows / Linux |
| -------------------------- | ----- | --------------- |
| Novo documento             | ⌘N    | ctrl+N          |
| Abrir                      | ⌘O    | ctrl+O          |
| Salvar                     | ⌘S    | ctrl+S          |
| Fechar guia                | ⌘W    | ctrl+W          |
| Abrir este manual          | F1    | F1              |
| Recolher a Faixa de Opções | ⌥⌘R   | Ctrl+F1         |

**Recolher a Faixa de Opções** funciona em todos os editores. A linha de guias mantém-se e a faixa de comandos por baixo desaparece; a guia selecionada serve também de controle de recolhimento, por isso enquanto a faixa estiver recolhida não há nenhuma guia selecionada e pressionar qualquer uma das guias traz a faixa de volta. Um duplo clique em uma guia faz o mesmo. O sentido em que a deixou é salvo editor a editor. O Docs acrescenta uma densidade intermediária: **⌥⌘K** / Ctrl+Alt+K compacta a faixa só em ícones, memorizada do mesmo jeito.

Os atalhos dentro de cada editor (pincel de formatação, pesquisar e substituir, operações de tabela, ...) estão nos respectivos capítulos; o Docs traz ainda uma janela de atalhos de teclado pesquisável (**⌘/**) (consulte o seu capítulo).

## Os atalhos Option+Command

Option+Command é a camada que o Word reserva para saltos estruturados, e o GenOffice preenche-a da mesma forma. O Docs fica com a maioria deles, o Sheets fica com dois seus para ficar em paridade com o Excel; um atalho funciona em todo o lado.

**Docs**

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

**Sheets**, com a grade com foco

| Atalho (macOS) | O que faz       | Windows / Linux |
| -------------- | --------------- | --------------- |
| ⌥⌘0            | Bordas externas | Ctrl+Shift+7    |
| ⌥⌘−            | Sem borda       | Ctrl+Shift+−    |

O Windows não é uma reescrita do par do Mac. O Excel para Mac dá ao Sheets **ambos** — ⌘⇧7 e ⌥⌘0 são duas teclas para as mesmas bordas externas —, por isso no Windows o comando mantém a posição Ctrl+Shift que já tinha, e a camada Option simplesmente não existe.

Repare que **⌥⌘0 significa Normal no Docs e Bordas externas no Sheets**. Nunca aparecem no mesmo editor, por isso na prática nada colide, mas ⌥⌘0 está ocupado e não está disponível como atalho global.

**Todos os editores**: **⌥⌘R / Ctrl+F1** recolhe a faixa de opções, tal como descrito acima.

Isso deixa o ⌥⌘D livre para o GenOffice usar no macOS, se um comando futuro precisar dele.

## Para onde ir a seguir

- Onde ficam os arquivos: [A tela Início](help://home-screen).
- Gerenciar muitos arquivos abertos: [Guias e gerenciamento de janelas](help://tabs-and-windows).
- Deixar o trabalho à IA: [O painel do assistente de IA](help://ai-panel).
- Idioma, tema, aplicativos padrão: [Configurações, idioma, tema e integrações MCP](help://settings-integrations).
