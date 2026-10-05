# O ecrã Início: onde ficam os seus arquivos

O ecrã Início é a página inicial do GenOffice: uma barra de navegação à esquerda e, à direita, as listas de arquivos e os cartões de criação rápida.

![O ecrã Início](img/home-screen.png)

## Navegação na barra lateral

- **Recentes**: os arquivos que você abriu recentemente, agrupados por período (esta semana / este mês / anteriores).
- **Favoritos**: os arquivos que você marcou com uma estrela. Passe o cursor sobre a linha do arquivo e clique na estrela para adicioná-lo ou removê-lo.
- **Genspark Projects**: depois de entrar na sua conta Genspark, mostra os projetos que você criou com o Genspark AI na web; clique em um deles para continuar editando no navegador. Há pesquisa, ordenação por data, atualização e carregar mais.
- **Pastas**: fixe diretórios usados com frequência na barra lateral (Adicionar pasta…) e salte até eles como quem usa marcadores. Raízes indisponíveis aparecem como indisponíveis e podem ser removidas da lista.
- **Lixeira**: aponta para a lixeira do sistema — os arquivos excluídos vão para lá e podem ser restaurados pelo sistema operacional.

## A lista de arquivos

Cada linha mostra um ícone, o nome do arquivo, a data de modificação e mais. O **menu ⋯** da linha oferece:

- **Renomear**: no próprio local, com a extensão preservada automaticamente.
- **Adicionar aos favoritos / Remover dos favoritos**
- **Duplicar**: cria uma cópia na mesma pasta.
- **Excluir**: move o arquivo para a lixeira do sistema — não é uma exclusão definitiva.
- **Mostrar na pasta**: localiza o arquivo no seu gestor de ficheiros.

## Pesquisa

A caixa de pesquisa no topo corresponde a duas coisas ao mesmo tempo:

- **Nomes de arquivos**: filtragem rápida por nome.
- **Conteúdo dos arquivos**: o GenOffice indexa seus arquivos em segundo plano (texto dentro de docx/xlsx/pptx/pdf/md/html, com recurso a OCR para PDFs digitalizados), pelo que pesquisar no corpo do texto também encontra arquivos. O âmbito e as opções ficam nas configurações de pesquisa.

## Cartões de início rápido

Os cartões acima das listas criam um documento novo num só passo. Clicar num cartão cria um arquivo desse tipo e abre o respetivo editor — comece a escrever de imediato, ou deixe o AI fazer o rascunho por si (todos os editores têm um **botão de IA** no friso e **Perguntar à IA** no menu de contexto da seleção).

O arquivo novo vai parar à pasta atualmente selecionada na barra lateral; se não houver nenhuma seleção, vai para a pasta predefinida.

O que faz cada cartão:

- **AI Docs** (.docx): um documento de texto em branco no editor Docs. O arquivo só é gravado no disco no **primeiro salvamento**; e os documentos novos abrem com o painel de IA expandido (desligue em Configurações → "Abrir o painel de IA em novos documentos").
- **AI Sheets** (.xlsx): uma folha de cálculo em branco no editor Sheets. Até você salvar, não existe arquivo no disco — o nome fica reservado para o primeiro salvamento; após a primeira geração pela IA, o arquivo também pode ser renomeado automaticamente a partir do seu conteúdo.
- **AI Slides** (.pptx): uma apresentação em branco no editor Slides.
- **AI Markdown** (.md): um documento Markdown em branco no editor Markdown.
- **AI HTML** (.html): uma página web em branco no editor HTML.
- **AI PDF** (.pdf): diferente dos restantes — cria **imediatamente** um PDF em branco verdadeiro de uma página na pasta de destino e abre-o como um ficheiro normal (o editor de PDF trabalha sobre ficheiros reais). Bom para anotar, ocultar partes do documento ou acrescentar texto; o arquivo pode ser renomeado automaticamente a partir do conteúdo no primeiro salvamento.
- **Abrir arquivo local**: um seletor de arquivos do sistema para Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) e páginas web (.html/.htm). É possível selecionar vários; cada arquivo recebe a sua própria guia.

> Dica: Arquivo ▸ Novo na barra de menus cria os mesmos tipos de documento (⌘N/Ctrl+N usa por omissão um documento de texto); arrastar um arquivo para a janela abre-o.

## Projetos na nuvem (Genspark Projects)

- O primeiro uso exige entrar na sua conta Genspark (fluxo de código de dispositivo: o GenOffice mostra um código e você conclui o início de sessão no navegador).
- A lista de projetos fica sincronizada com a web; Abrir no navegador leva lá para continuar.
- Não entrar não afeta nenhuma das funcionalidades locais.
