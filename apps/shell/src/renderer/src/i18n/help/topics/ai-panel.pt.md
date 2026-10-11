# O painel do assistente de IA

Todos os editores podem chamar o painel de IA: selecione algo, dê uma instrução e veja o resultado a chegar.

## Abrir e usar

![O painel de IA no Docs](img/ai-panel.png)

- Pontos de entrada: o **botão de IA** no friso de cada editor, **Perguntar à IA** nos menus de contexto, ou Perguntar à IA na barra de marcação.
- Descreva a tarefa por palavras normais (reescrever isto / transformar esta coluna em percentagens / refazer a disposição desta página...) e prima Enter.
- As respostas são renderizadas **em fluxo**; quando a IA precisa de ferramentas (ler o documento, editá-lo), executa-as e continua até terminar.
- **Parar**: interrompe a volta em curso a qualquer momento.
- **Fila de espera**: enquanto uma resposta chega, Enter coloca sua mensagem na fila em vez de descartá-la — uma faixa acima do campo de entrada conta quantas; abra-a para editar, remover, reordenar, pausar ou limpar. A mensagem seguinte parte sozinha ao fim do turno; Esc continua parando o turno.

## O que sabe fazer

- **Docs**: reescrever, desenvolver, traduzir, resumir, inserir tabelas e imagens, ajustar a formatação; cada volta guarda um instantâneo primeiro.
- **Sheets**: fórmulas, preenchimento de dados, transformações em massa, formatação.
- **Slides**: gerar a apresentação inteira, ajustar o layout, reescrever os textos.
- **PDF**: perguntas e respostas e resumos com base no texto ou nas páginas selecionadas.
- **Markdown / HTML**: reescrever, desenvolver, traduzir.
- **Início ▸ Ask AI** responde a partir dos seus arquivos favoritados e os cita — o painel de base de conhecimento tem a própria seção no capítulo Início, e o painel de IA do Docs carrega a mesma capacidade.

## Reversão e segurança

- O painel do Docs mantém uma **lista de versões**: um instantâneo por volta, pode voltar a qualquer um deles, e a própria reversão pode ser anulada com Ctrl+Z. Os instantâneos sobrevivem a fechar e reabrir o documento.
- As edições da IA passam pelo mesmo fluxo de edição das manuais (anuláveis e sujeitas a salvamento) — nada contorna a sua confirmação de salvamento.

## Ocultar uma seleção da IA

- Em Slides, HTML e Markdown uma seleção pode sair da vista do modelo: clique direito ▸ **Ocultar a seleção da IA**, dê nome ao espaço reservado, e o modelo lê `{{nome}}` no lugar. (Markdown arma o item com a chave de olho riscado da faixa; a do HTML vive no painel de código.)
- O arquivo guarda o texto real — salvar, exportar e imprimir ficam intactos, e a marca sobrevive à reabertura. O mesmo item desfaz (Markdown ganha também **Mostrar a seleção à IA novamente**).
- HTML oculta conteúdos de elementos, valores de atributos e literais de string dos scripts; a página renderizada nunca muda.
- Slides também oculta imagens, vídeos e áudios inteiros, e um slide com conteúdo oculto nunca é enviado como captura: a checagem automática de layout pula a página e avisa.
- Isso vale só para a IA do aplicativo — a CLI, as ferramentas MCP e a exportação headless leem o arquivo diretamente e continuam vendo o texto.

## Privacidade

- As instruções e o conteúdo relevante do documento vão para **o serviço de modelo que configurou** (Genspark ou um endpoint personalizado, ver o capítulo seguinte); sem configuração, nada é enviado.
- Os ficheiros locais não são carregados para lado nenhum; as chaves BYOK ficam guardadas no ficheiro de definições da aplicação nesta máquina e são enviadas apenas nos cabeçalhos dos pedidos.
