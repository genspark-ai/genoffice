# Guias e gerenciamento de janelas

Todos os arquivos abertos compartilham uma janela; a barra de guias no topo muda entre eles, ao estilo de um navegador.

![A barra de guias: Início e dois documentos](img/tabs.png)

## Noções básicas

- **Mudar**: clique em uma guia, ou gire a roda do mouse sobre a barra para percorrer as guias.
- **Fechar**: o × na guia, ou ⌘W/ctrl+W. Se houver alterações não salvas, o GenOffice pergunta primeiro.
- **Nova**: o + no extremo direito da barra.
- **Ativação imediata**: pressionar uma guia muda para ela logo que é pressionada, sem esperar que o clique termine.

## Renomear: clique duas vezes em uma guia

**Clique duas vezes em qualquer guia de arquivo** e o título transforma-se em um campo de escrita no local: Enter confirma, Escape cancela, e sair do campo também confirma; o Enter que confirma uma sugestão do IME não é confundido com uma confirmação. A renomeação passa pelas mesmas verificações de segurança que a renomeação de uma linha na tela Início (caracteres proibidos, nomes em conflito), renomeia o arquivo no disco e atualiza a lista de recentes.

## Arrastar para reordenar

Pressione e arraste uma guia para o lado para reordenar; as guias vizinhas abrem espaço de imediato e a ordem é definida quando você solta o mouse. Uma zona morta de 4 pixels impede que cliques simples façam as guias oscilar.

## Menu de contexto

Clique com o botão direito em uma guia para **Abrir em nova janela** — disponível em cada guia de documento que possa destacar — e **Fechar**, desativado em uma guia que não possa ser fechada. Ambas seguem o idioma do aplicativo.

## Destacar em uma janela

**Arraste uma guia para fora da barra** (ou use Abrir em nova janela) e ela passa a ser uma janela própria com o documento ativo; arraste-a de volta para a encaixar na barra. Uma janela destacada é um par em tudo — pode continuar a editar e a salvar lá.

## A lista de todas as guias

Quando as guias não cabem na barra, o pequeno ícone da barra de guias no extremo direito abre a lista completa (um menu do sistema, nunca tapado pela área de conteúdo); escolha com as setas.

## A barra de ferramentas de cada editor

Cada guia de editor tem uma barra de ferramentas no topo (a disposição exata varia um pouco de editor para editor):

- **Salvar** (⌘S/ctrl+S) e **Salvar Como…**.
- **Desfazer / Refazer**: vários níveis; cada editor mantém seu próprio histórico.
- **Localizar** (ctrl+F): abre o painel de pesquisa desse editor.
- Interruptor de **Salvamento automático**: quando está ligado, as alterações são gravadas no disco periodicamente; quando está desligado, só o salvamento manual grava (e o marcador de não salvo na guia ou no título lembra-lhe disso).

## A guia Início

A guia Início mais à esquerda não pode ser fechada; para voltar de qualquer editor, clique nela ou use Arquivo ▸ **Voltar ao início**.
