# PDF: leitura, anotações e ocultação

O editor de PDF tem cinco separadores no friso: **Página Inicial / Anotar / Editar / Páginas / Exibir**. Ele lê e escreve: o texto pode ser editado, o conteúdo pode ser ocultado, assinado e os formulários podem ser preenchidos.

## Leitura e navegação

- Barra lateral esquerda: **Miniaturas** (clique para saltar, o intervalo visível fica destacado) ou **Estrutura** (os marcadores, quando existem).
- Zoom: o controlo de proporção no canto inferior direito; ctrl+roda ajusta o zoom por degraus.
- Rotação: por página ou em todas as páginas a partir do menu Páginas; as rotações são gravadas ao salvar.
- Pesquisar: ctrl+F para procurar em todo o texto com todas as ocorrências realçadas.
- PDFs encriptados: aparece um pedido de palavra-passe (uma pequena janela própria) para os abrir; a palavra-passe só é usada nesta sessão.

## Selecionar texto e marcar (Anotar)

Experimente num parágrafo qualquer:

1. **Arraste o rato por cima de uma frase** — ao largar, uma barra de anotação surge por cima do texto:

![A barra de anotação depois de selecionar texto](img/pdf-highlight.png)

2. Escolha **Destacar** (o quadrado amarelo abre uma paleta de cores), **Sublinhar** ou **Tachar**; **Perguntar à IA** envia a seleção com a sua pergunta para o painel de IA.
3. Para desfazer uma anotação, volte a arrastar-seleccionar o mesmo trecho e clique no botão ativo da barra (um interruptor ao estilo do Word), ou selecione-o e prima Delete.

A ter em conta:

- Arrastar sobre o texto faz aparecer uma barra: **Destacar / Sublinhar / Tachar / Copiar / Perguntar à IA**.
- As cores vêm da paleta; e **aplicar a mesma marcação a um intervalo já marcado remove-a** (interruptor ao estilo do Word).
- As marcações já gravadas no ficheiro podem ser selecionadas e eliminadas (menu ⋯ ou Delete).
- **Nota**: enquanto uma ferramenta de desenho estiver armada, a camada de texto não pode ser selecionada — a ferramenta desarma-se sozinha após cada colocação, ficando outra vez em modo de seleção para a ação seguinte.

## Ferramentas de desenho (Anotar)

Seis ferramentas: **Desenho livre, Retângulo, Elipse, Seta, Nota**, mais a **Censurar área** no separador Anotar.

- Cada ferramenta é um interruptor: clique para a armar; **ela desarma-se assim que uma forma é colocada** (clique outra vez na ferramenta para continuar); clicar na ferramenta armada também a desarma.
- O traço segue a espessura do pincel; o retângulo, a elipse e a seta desenham-se arrastando; as cores vêm da paleta de desenho.
- As formas colocadas podem ser selecionadas, eliminadas, arrastadas e (retângulo/elipse) redimensionadas.
- **Ocultar conteúdo, o fluxo completo** (esconder uma linha de texto):

  1. Separador Anotar ▸ clique **Censurar área** (a ferramenta arma-se).
  2. **Arraste uma caixa sobre o conteúdo** — fica coberto por um traço hachurado, e a barra de ferramentas ganha os botões **Limpar marcações** e **Aplicar censura**:

  ![A página depois de marcar uma área para censurar](img/pdf-redact.png)

  3. Clique **Aplicar censura** e confirme — o resultado é uma cópia de trabalho em que o texto e as imagens cobertas são removidos a sério (não apenas tapados), e isso não pode ser anulado; o documento original fica intacto.

  Errou? O botão de limpar marcações apaga as marcações atuais para poder desenhar de novo.

## Notas adesivas e fio de comentários

- A **ferramenta de nota** coloca um alfinete e abre um cartão na margem para o texto (o nome do autor é configurável); depois de confirmar, guarda-o como uma anotação de Texto PDF padrão.
- Clique num alfinete para abrir o fio: **Responder** (fios planos ao estilo do WPS/Acrobat), **Editar** o seu comentário, **Eliminar** um comentário ou o fio inteiro.
- As edições em curso sobrevivem até o gravação escrever o novo texto na mesma anotação do ficheiro, mantendo a cadeia de respostas intacta.

## Editar o conteúdo do PDF (Editar)

- **Editar texto**: clique no texto para o editar por blocos (motor pdfium; as fontes são aproximadas tanto quanto possível).
- **Inserir texto**: coloque texto pesquisável com a fonte, o tamanho e a cor que quiser.
- **Inserir imagem / carimbo**.
- Formulários: os campos AcroForm preenchem-se diretamente; os valores são escritos ao salvar.

## Assinaturas

- **Assinatura manuscrita**: desenhe-a; pode ser associada a um campo de assinatura de um formulário.
- **Assinatura por imagem**: coloque uma imagem como assinatura.
- As assinaturas guardadas podem ser reutilizadas.

## Operações de página (Páginas)

- **Girar / Excluir página / Reordenar**: arraste as miniaturas para reordenar; a exclusão pede confirmação.
- **Importar páginas**: traga páginas de outro PDF para o documento. **Inserir página em branco** adiciona uma página vazia.
- **Substituir páginas** troca um intervalo por páginas oriundas de outro documento; **Recortar páginas** apara as margens, com a opção de aplicar a todas as páginas.
- **Tamanho da página** redimensiona todas as páginas para um único tamanho de papel; **Inverter ordem** inverte o documento de trás para a frente.
- **Extrair página**: exporta as páginas selecionadas para um PDF novo.
- **Dividir PDF**: duas formas — dividir por intervalos em vários ficheiros, ou cortar cada página numa grelha de páginas mais pequenas.
- **Mesclar PDF**: duas formas — acrescentar outros PDFs, ou combinar várias páginas numa só folha. Os tamanhos são somados **antes** de se ler seja o que for, e um total **acima de 1 GiB é recusado** com uma mensagem legível (para manter a memória limitada).
- As alterações ao nível da página são gravadas no salvamento seguinte; Salvar Como deixa o original intacto.

## Exportar e imprimir

- **Exportar como Word… / PowerPoint… / Excel…** no menu Arquivo, ou os mesmos três a partir de **Converter PDF** no friso — tudo local, sem envio. O .pptx sai com um diapositivo por página e o .xlsx com uma folha de cálculo por página. Cada um pergunta onde guardar.
- **Imprimir**: a ordem e as rotações atuais através da caixa de diálogo do sistema; são suportados intervalos de páginas.

## Salvar

- O salvamento normal/automático grava as anotações e as edições de volta no ficheiro (de forma atómica).
- **A ocultação passa pelo seu próprio fluxo Aplicar censura**, que produz uma cópia e deixa o original intacto, para que o conteúdo sensível não fique nele.
