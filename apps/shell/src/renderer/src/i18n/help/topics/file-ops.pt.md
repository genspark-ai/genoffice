# Operações de arquivo: renomear, excluir, exportar

Este capítulo cobre as operações de arquivo comuns a todos os editores; as opções de exportação de cada editor estão no respectivo capítulo.

O **menu ⋯** da linha do arquivo (passe o cursor sobre a linha do arquivo) reúne estas ações:

![O menu ⋯ de uma linha de arquivo](img/file-ops.png)

## Renomear

Dois pontos de entrada, um conjunto de verificações:

- **⋯ ▸ Renomear** a partir de uma linha na tela Início.
- **Clique duas vezes em uma guia de arquivo** para renomear no local (consulte [Guias e gerenciamento de janelas](help://tabs-and-windows)).

As regras: a extensão é preservada automaticamente; caracteres proibidos, pontos no final e nomes reservados (CON/NUL e afins) são recusados com uma mensagem; e um nome em conflito na mesma pasta também é bloqueado. O arquivo real no disco é renomeado, e a lista de recentes e os favoritos acompanham.

## Excluir

- **⋯ ▸ Excluir** no Início: pergunta quais os arquivos e depois move-os para a **lixeira do sistema**, onde o sistema operacional os pode restaurar. O GenOffice não mantém um histórico próprio para desfazer essa operação — a recuperação é responsabilidade da lixeira, não de um aviso.

## Duplicar

**⋯ ▸ Duplicar** cria uma cópia de <nome> na mesma pasta; se houver colisão de nome, um contador é acrescentado automaticamente. A cópia entra nos **Recentes** em vez de ser aberta por si — a um clique de distância, não à sua frente.

## Salvar e Salvar Como

- **⌘S / ctrl+S** salva o arquivo atual; um arquivo sem título pergunta primeiro pela localização e pelo nome.
- **Salvar Como…** grava um arquivo novo e deixa o original intacto; as edições seguintes passam a incidir sobre o arquivo novo.
- Cada gravação é atômica (arquivo temporário + mudança de nome), por isso sair durante a gravação não corrompe o arquivo.
- O Salvamento automático só entra em vigor depois do primeiro salvamento manual (consulte [Início rápido](help://getting-started)).

## Exportar para PDF

- **Docs**: Arquivo ▸ Exportar como PDF… (ou o botão na faixa de opções), paginado conforme a disposição atual.
- **Slides**: a exportação rasteriza página a página, com progresso para apresentações grandes.
- **Sheets**: a exportação segue a paginação de impressão.
- As exportações são renderizadas em uma janela oculta e o resultado fica onde você escolher.

## Exportar para Word / imagens

- **PDF ▸ PDF para Word**: transforma o PDF em um .docx (convertido localmente; layouts complexos são preservados tanto quanto possível).
- O **Docs** pode exportar páginas como imagens (PNG por página).

## Imprimir

Arquivo ▸ Imprimir… em cada editor (⌘P/ctrl+P) abre a caixa de diálogo de impressão do sistema; os PDFs imprimem com a ordem de páginas e as rotações atuais.

## Onde vivem os arquivos sem título

A localização que escolher no primeiro salvamento é a casa dele; antes disso o documento só existe em memória. O Salvamento automático só assume a partir desse primeiro salvamento.
