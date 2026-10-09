# Docs: processamento de texto

O Docs é o processador de texto parecido com o Word: lê e escreve .docx a sério, com uma paginação WYSIWYG verdadeira.

## O friso

Separadores: **Página Inicial / Inserir / Desenhar / Layout / Design / Referências / Revisão / Exibir**, mais separadores contextuais para o objeto selecionado (design e layout da tabela, imagens, cabeçalho e rodapé).

- **Página Inicial**: área de transferência; fonte (incluindo tamanhos CJK e marcas de ênfase) com **Limpar Toda a Formatação** e um interruptor para **Mostrar/ocultar marcas de formatação**; parágrafo (alinhamento/recuo/espaçamento/listas) mais **Definir novo marcador / Definir novo formato de número / Lista de Vários Níveis**, que guardam os seus próprios estilos de lista no documento; estilos (Título 1-6/Normal/Citação, modificáveis) com um **Painel de Estilos** para a lista completa.
- **Inserir**: quebras de página e de secção; tabelas (uma grelha de linhas × colunas, ou **Inserir Tabela…** para um tamanho exato); imagens, formas, caixas de texto; **Folha de Rosto** e **Página em Branco** a partir de uma galeria predefinida; **Gráfico**; **Letra Capitular**; **WordArt**; campos (data, hora, número de página, total de páginas, nome do arquivo); hiperligações, **Indicador** e referências cruzadas; comentários; cabeçalho e rodapé e números de página; símbolos e equações.
  - **Gráfico** insere um verdadeiro objeto de gráfico com os seus próprios dados — de barras, de linhas ou circular — e não uma imagem. _Editar dados_ do Word abre os números que estão por trás.
- **Desenhar**: tinta sobre a página, no grupo **Ferramentas de Desenho** — **Selecionar** volta à edição de texto, depois **Caneta**, **Marca-Texto** e **Borracha** (um clique ou um gesto apaga o traço inteiro). Ao lado, **Estilo da Caneta** / **Estilo do Marca-Texto** é um único controlo com amostras de cor e uma fila de espessuras; o respetivo rótulo segue a ferramenta ativa. A tinta é guardada no documento como uma anotação que flutua por cima do texto, por isso sobrevive a guardar e reabrir, e **Limpar Tudo**, no grupo seguinte, remove-a toda.
- **Layout**: margens, orientação e tamanho do papel, colunas, recuos e espaçamento dos parágrafos.
- **Design**: temas, conjuntos de cores, marca de água, bordas da página.
- **Referências**: sumário (atualizável), notas de rodapé e de fim, legendas, referências cruzadas.

  ![O separador Referências](img/docs-references.png)

- **Revisão**: **Editor** revê a ortografia, a gramática e a pontuação de todo o documento; **Traduzir**; verificação ortográfica; comentários (**Comentários com IA** trata os que estão abertos); controlar alterações com as vistas Toda a Marcação / Marcação Simples, aceitar/rejeitar e **Resumo de revisões com IA**; contagem de palavras; **Comparar** com outro ficheiro; **Proteger Documento**.
- **Exibir**: cinco formas de ver o ficheiro — **Layout de Impressão**, **Layout da Web**, **Estrutura de Tópicos**, **Modo de Leitura** e **Visualização de Páginas**; reduzir/ampliar/100 %/largura da página/uma página; **Painel de IA**; **Modo Escuro**; régua, linhas de grade e o painel de navegação; **Nova Guia**, **Dividir** e **Alternar Guias**; a janela pesquisável de **atalhos de teclado**.
  - O **Modo Escuro** escurece a página e a tela à volta dela, nunca o friso — a divisão do Word entre uma superfície de edição escura e uma janela escura. A escolha fica guardada e, em qualquer caso, prevalece sobre o tema da aplicação.
  - **Dividir** abre um segundo painel por baixo, que desliza de forma independente e espelha o primeiro; feche-o com o × no seu limite.

## O painel de navegação

**Exibir ▸ Painel de Navegação** abre um painel lateral com a estrutura de títulos do documento, uma caixa de pesquisa sobre todo o documento e uma miniatura por página. O facto de estar aberto é memorizado entre execuções, pelo que um documento em que se navega pela estrutura continua navegável.

**A estrutura** é a árvore de títulos. Clique com o botão direito num título da estrutura para o dobrar ou reestruturar, e não apenas para navegar:

- **Recolher / Expandir** num título dobra toda a sua subárvore — o capítulo desaparece, o texto dele continua no documento.
- **Recolher tudo / Expandir tudo** dobra ou desdobra tudo de uma vez. Num relatório longo é a diferença entre uma estrutura legível e uma parede de texto.
- **Mostrar níveis de título** filtra a árvore até às profundidades que lhe interessam, deixando-lhe com _Mostrar título 1_ um índice que se consegue mesmo folhear.
- **Promover / Rebaixar** mudam o nível do título, e com ele o nível que todos os títulos abaixo herdam — é assim que um capítulo passa a secção.
- **Novo título antes / Novo título depois** inserem um no ponto de inserção, sem sair do painel.
- **Excluir** remove o título _e tudo o que está por baixo dele_, e é este o item a ter cuidado: é uma eliminação de subárvore, não uma eliminação de linha.
- **Selecionar título e conteúdo** seleciona do título até ao fim da sua subárvore, pronto para uma edição de toda a secção.

## Menu de contexto

Clique com o botão direito em qualquer ponto do texto — o menu acompanha aquilo em que clicou. Os principais grupos:

- **Área de transferência**: recortar / copiar / colar / **colar como texto simples**.
- **Fonte, parágrafo**: mudar a família e o tamanho, negrito/itálico/sublinhado, alinhamento/recuo/espaçamento sem passar pelo friso.
- **Sinônimos**: lista sinônimos da palavra selecionada; clique num deles para substituir.
- **Traduzir** (IA): traduz a seleção para o idioma de destino (inglês, chinês simplificado, japonês, coreano, francês, alemão, espanhol, ...) pelo painel de IA.
- **Novo Comentário**: anexa um comentário à seleção.
- **Ortografia** (numa palavra escrita errado): substituições sugeridas, ignorar tudo, adicionar ao dicionário, definir o idioma de verificação.
- **Hiperligação**: abrir / editar / copiar ligação / remover hiperligação.
- **Imagem**: ver imagem, salvar imagem como…, **quebrar texto automaticamente** (na linha / quadrado à esquerda e à direita / acima e abaixo / atrás do texto / à frente do texto), ordem de disposição.
- **Campos** (sumário, números de página): atualizar campo / alternar códigos de campo / editar campo.
- **Numeração de lista** (dentro de uma lista): reiniciar numeração / continuar numeração / alterar nível da lista / definir valor de numeração.
- **Tabela** (cursor dentro de uma tabela): inserir linhas/colunas, mesclar / dividir células, dividir tabela, ajuste automático, alinhamento da célula, distribuir linhas/colunas, propriedades da tabela, menu excluir, selecionar.

## Edição

- Pesquisar e substituir (ctrl+F / ctrl+H): diferenciar maiúsculas de minúsculas, palavra inteira, expressões regulares.
- Pincel de formatação; desfazer/refazer profundo; opções de colagem.
- Tabelas: mesclar/dividir células, operações de linhas e colunas, bordas e sombreamento, classificação, fórmulas.
- Imagens: quebra de texto, corte, compactação; tela de desenho.

## Tipografia CJK

- A compactação de pontuação e as regras de quebra de linha kinsoku correspondem às do Word; há conversão entre largura plena e meia largura.
- Os candidatos de fonte cobrem os nomes de família CJK mais comuns no Windows e no macOS.

## IA

- O botão de IA no friso e o painel lateral: reescrever, expandir, traduzir, resumir, inserir tabela, além de instruções livres.
- Cada turno da IA guarda um instantâneo antes de começar; é possível voltar a qualquer um deles na lista de versões, e o próprio retrocesso pode ser desfeito.

## Salvar e exportar

- Salvar em .docx reescreve apenas os parágrafos que mudaram — o conteúdo não tocado permanece byte a byte igual.
- Exporta PDF (com a paginação atual) e imagens de cada página.

## Imprimir

ctrl+P pela caixa de diálogo do sistema, com as páginas exatamente como serão impressas.
