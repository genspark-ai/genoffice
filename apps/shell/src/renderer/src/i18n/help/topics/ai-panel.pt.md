# O painel do assistente de IA

Todos os editores podem chamar o painel de IA: selecione algo, dê uma instrução e veja o resultado a chegar.

## Abrir e usar

![O painel de IA no Docs](img/ai-panel.png)

- Pontos de entrada: o **botão de IA** no friso de cada editor, **Perguntar à IA** nos menus de contexto, ou Perguntar à IA na barra de marcação.
- Descreva a tarefa por palavras normais (reescrever isto / transformar esta coluna em percentagens / refazer a disposição desta página...) e prima Enter.
- As respostas são renderizadas **em fluxo**; quando a IA precisa de ferramentas (ler o documento, editá-lo, executar um script), executa-as e continua até terminar.
- **Parar**: interrompe a volta em curso a qualquer momento.

## O que sabe fazer

- **Docs**: reescrever, desenvolver, traduzir, resumir, inserir tabelas e imagens, ajustar a formatação; cada volta guarda um instantâneo primeiro.
- **Sheets**: fórmulas, preenchimento de dados, transformações em massa, formatação.
- **Slides**: gerar a apresentação inteira, ajustar o layout, reescrever os textos.
- **PDF**: perguntas e respostas e resumos com base no texto ou nas páginas selecionadas.
- **Markdown / HTML**: reescrever, desenvolver, traduzir.

## Reversão e segurança

- O painel do Docs mantém uma **lista de versões**: um instantâneo por volta, pode voltar a qualquer um deles, e a própria reversão pode ser anulada com Ctrl+Z. Os instantâneos sobrevivem a fechar e reabrir o documento.
- As edições da IA passam pelo mesmo fluxo de edição das manuais (anuláveis e sujeitas a salvamento) — nada contorna a sua confirmação de salvamento.

## Privacidade

- As instruções e o conteúdo relevante do documento vão para **o serviço de modelo que configurou** (Genspark ou um endpoint personalizado, ver o capítulo seguinte); sem configuração, nada é enviado.
- Os ficheiros locais não são carregados para lado nenhum; as chaves BYOK vivem apenas nos cabeçalhos das pedidos — nunca em disco nem nos registos.
