# Operações de arquivo: renomear, excluir, exportar

Este capítulo cobre as operações de arquivo comuns a todos os editores; as opções de exportação de cada editor estão no respetivo capítulo.

O **menu ⋯** da linha do arquivo (passe o cursor sobre a linha do arquivo) reúne estas ações:

![O menu ⋯ de uma linha de arquivo](img/file-ops.png)

## Renomear

Dois pontos de entrada, um conjunto de verificações:

- **⋯ ▸ Renomear** a partir de uma linha no ecrã Início.
- **Clique duas vezes numa guia de ficheiro** para renomear no local (consulte [Guias e gestão de janelas](help://tabs-and-windows)).

As regras: a extensão é preservada automaticamente; caracteres proibidos, pontos no final e nomes reservados (CON/NUL e afins) são recusados com uma mensagem; e um nome em conflito na mesma pasta também é bloqueado. O ficheiro real no disco é renomeado, e a lista de recentes e os favoritos acompanham.

## Excluir

- **⋯ ▸ Excluir** no Início: move o arquivo para a **lixeira do sistema**, de onde pode ser restaurado pelo sistema operacional.
- Após a exclusão aparece um aviso com opção de desfazer durante alguns segundos — desfazer devolve o arquivo ao lugar onde estava.

## Duplicar

**⋯ ▸ Duplicar** cria uma cópia de <nome> na mesma pasta e abre-a numa guia nova; se houver colisão de nome, um contador é acrescentado automaticamente.

## Salvar e Salvar Como

- **⌘S / ctrl+S** salva o ficheiro atual; um ficheiro sem título pergunta primeiro pela localização e pelo nome.
- **Salvar Como…** grava um ficheiro novo e deixa o original intacto; as edições seguintes passam a incidir sobre o ficheiro novo.
- Cada gravação é atómica (ficheiro temporário + mudança de nome), pelo que sair a meio da escrita não corrompe o ficheiro.
- O Salvamento automático só entra em vigor depois do primeiro salvamento manual (consulte [Início rápido](help://getting-started)).

## Exportar para PDF

- **Docs**: Arquivo ▸ Exportar como PDF… (ou o botão no friso), paginado conforme a disposição atual.
- **Slides**: a exportação rasteriza página a página, com progresso para apresentações grandes.
- **Sheets**: a exportação segue a paginação de impressão.
- As exportações são renderizadas numa janela oculta e o resultado fica onde você escolher.

## Exportar para Word / imagens

- **PDF ▸ PDF para Word**: transforma o PDF num .docx (convertido localmente;layouts complexos são preservados tanto quanto possível).
- O **Docs** pode exportar páginas como imagens (PNG por página).

## Imprimir

Arquivo ▸ Imprimir… em cada editor (⌘P/ctrl+P) abre a caixa de diálogo de impressão do sistema; os PDFs imprimem com a ordem de páginas e as rotações atuais.

## Onde vivem os ficheiros sem título

A localização que escolher no primeiro salvamento é a casa dele; antes disso o documento só existe em memória. O Salvamento automático só assume a partir desse primeiro salvamento.
