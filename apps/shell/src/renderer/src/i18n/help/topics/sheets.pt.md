# Sheets: folhas de cálculo

O Sheets é o editor parecido com o Excel; os cálculos correm num processo separado do motor em Rust (se esse processo falhar, a aplicação não cai). Abre e guarda .xlsx a sério; .csv e .tsv abrem-se como tabelas.

## A interface

- **Friso**: oito separadores, descritos um a um abaixo.
- **Barra de Fórmulas**: mostra e edita a fórmula da célula ativa; suporta as funções mais comuns.
- **Guias de folha** (em baixo): adicionar / renomear / eliminar / mover folhas.
- **Edição de célula**: clique duas vezes ou escreva diretamente; Enter confirma e desce, Tab vai para a direita, Escape cancela (os hábitos do Excel).
- **Atalhos**: alinhados com a família Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Separadores do friso

- **Página Inicial**: fonte, preenchimento, bordas, formatos de número (moeda/percentagem/milhares, aumentar/diminuir decimais), alinhamento, merger, inserir linhas/colunas e dimensioná-las, formatação condicional, formatar como tabela, estilos de célula, área de transferência e pincel de formatação, classificar e filtrar.
- **Inserir**: imagens, formas, caixas de texto, ligações, comentários, caixa de verificação, Cabeçalho e Rodapé, símbolos, equações; **nove tipos de gráfico** (colunas, barras, linhas, área, pizza, dispersão, radar, rosca e combinação de colunas + linhas), além de **Gráficos Recomendados**, que ordena os tipos para a seleção e mostra a pré-visualização; **Gráfico Dinâmico** a partir da célula onde está; **Minigráficos** (linhas, colunas, vitória/derrota); **Segmentação de Dados** e **Linha do Tempo** para filtrar uma Tabela Dinâmica.
  - Os gráficos, os gráficos dinâmicos e os minigráficos são objetos reais no livro guardado.
  - **Segmentação de Dados** e **Linha do Tempo** são controlos de sessão: a filtragem que fazem é guardada e o Excel mostra a mesma tabela dinâmica filtrada, mas o próprio botão de segmentação não faz parte do ficheiro.
- **Layout da Página**: cores e fontes do tema, interruptores de linhas de grade e títulos ao imprimir, visualização de quebra de página.
- **Fórmulas**: AutoSoma e inserção de funções, definir nomes (também a partir da seleção), rastrear precedentes/dependentes, janela de inspeção, recalcular folha/livro.
- **Dados**: classificar e filtrar (incluindo filtro avançado e limpar o filtro), texto para colunas, mesclar pastas de trabalho, atualizar tudo.
- **Revisão**: navegar pelos comentários (mostrar, anterior/seguinte), traduzir e um grupo **Proteção** — **Proteger Planilha**, **Proteger Pasta de Trabalho** e **Permitir Edição de Intervalos**.
  - A proteção é escrita no .xlsx e o que esta aplicação aplica não tem palavra-passe, por isso o mesmo botão passa a **Desproteger…** e anula-a. A proteção vinda de outro programa que _tem mesmo_ palavra-passe não pode ser removida daqui.
  - Os dois botões de proteção aplicam-se de imediato — não há nenhuma caixa de diálogo para cancelar, apenas uma nota na barra de estado a indicar que será escrito ao guardar.
  - **Permitir Edição de Intervalos** marca as células que continuam editáveis enquanto o resto da folha fica bloqueada.
- **Exibir**: interruptores de linhas de grade, **Barra de Fórmulas**, títulos e realce da linha e coluna ativas; zoom; **Normal** e **Visualização de Quebra de Página**. As linhas de grade e os títulos são guardados com a folha; o realce é uma preferência sua.
- **Design do Gráfico**: aparece com um gráfico selecionado — tipo de gráfico, estilos e cores, editar o intervalo de dados.

O separador Dados, botão a botão (da esquerda para a direita na figura):

![O separador Dados](img/sheets-data.png)

- **Tabela Dinâmica**: constrói uma tabela dinâmica a partir do intervalo atual; arraste campos para agregar.
- **Atualizar**: recalcula os dados da tabela dinâmica atual.
- **De Texto/CSV**: importa um .csv/.txt como uma folha nova, separando pelo delimitador.
- **Mesclar pastas de trabalho**: traz folhas de outros ficheiros .xlsx para este.
- **Atualizar Tudo**: recalcula todas as tabelas dinâmicas e todos os conjuntos de dados externos.
- **Classificar** (menu suspenso): crescente / decrescente / classificação personalizada (regras de várias colunas).
- **Filtro**: acrescenta menus ▼ à linha de cabeçalho; assinale os valores a manter.
- Os pequenos botões empilhados ao lado: **Limpar** (trazer todas as linhas de volta), **Reaplicar** (executar outra vez o filtro atual), **Avançado** (filtrar com um intervalo de critérios).
- **Texto para Colunas** (menu suspenso): divide uma coluna em várias pelo delimitador ou pela largura fixa.
- **Preenchimento Rápido**: dê um exemplo e o resto da coluna preenche-se por padrão (ctrl+E).
- **Remover Duplicatas**: elimina as linhas repetidas segundo as colunas selecionadas.
- **Validação de Dados** (menu suspenso): regras de introdução para a seleção (listas suspensas, intervalos numéricos, ...).
- **Consolidar**: reúne vários intervalos num só sítio por categoria.
- **Teste de Hipóteses** (menu suspenso): Atingir meta — resolve uma célula de entrada para que uma célula com fórmula chegue ao valor de destino.
- **Agrupar / Desagrupar** (menu suspenso): grupos de linhas e colunas com ocultar e expandível.
- **Subtotal**: insere linhas de subtotal por categoria.

O separador Fórmulas, botão a botão:

![O separador Fórmulas](img/sheets-formulas.png)

- **Inserir Função** (fx): procure funções com um assistente de argumentos.
- **AutoSoma** (menu suspenso): SUM com um clique, mais média/contagem/máximo/mínimo.
- **Usadas Recentemente / Financeira / Lógica / Texto / Data e Hora / Pesquisa e Referência / Matemática e Trigonometria / Mais**: percorra e insira funções por categoria.
- **Gerenciador de Nomes**: ver, criar e eliminar intervalos com nome.
- **Definir Nome** (menu suspenso): dê um nome à seleção; **Usar em Fórmula** insere um nome já existente; **Criar a partir da Seleção** dá nome aos intervalos a partir da respetiva linha ou coluna de cabeçalho.
- **Rastrear Precedentes / Rastrear Dependentes**: setas azuis mostram de onde vem o dados de uma fórmula e para onde esta alimenta; **Remover Setas** limpa-as.
- **Mostrar Fórmulas**: as células mostram a própria fórmula em vez do resultado.
- **Verificação de Erros**: localiza e explica erros de fórmula.
- **Janela de Inspeção**: fixe as células que lhe interessam e veja os valores ao vivo.
- **Opções de Cálculo** (menu suspenso): recálculo automático ou manual; no modo manual, **Calcular Agora** e **Calcular Planilha** disparam-no à mão.

## Números e formatação

- Formatos de número: geral, número, moeda, percentagem, data/hora, fração, científico e mais.
- Alinhamento, quebra de texto, células fundidas, bordas e preenchimento.
- Altura das linhas e largura das colunas por arrasto; clique duas vezes no limite para ajustar ao conteúdo.

## Dados

**Classificar e filtrar** (exemplo: ordenar uma coluna por ordem descendente):

1. Clique em **qualquer célula dessa coluna** (não precisa de selecionar a coluna inteira).
2. Separador Página Inicial ▸ **Classificar e Filtrar** ▸ **Decrescente**; as linhas inteiras reordenam-se em conjunto (a área é ordenada como um todo).
3. Para as suas próprias regras (várias colunas, por cor): o mesmo caminho, escolha **Classificação Personalizada**.
4. Filtro: selecione a linha de cabeçalho e clique **Classificar e Filtrar ▸ Filtro** — cada cabeçalho recebe um menu ▼ onde assina os valores a manter; limpar o filtro traz tudo de volta.

- Classificar e filtrar.
- Congelar painéis.
- .csv / .tsv: abrem-se de imediato como tabela (um tsv separado por tabulações é analisado como uma só tabela); ao guardar, o formato original é reescrito.

## Menus de contexto

- **Na grelha**: o menu do próprio editor (Univer) — recortar/copiar/colar, inserir e eliminar linhas/colunas, ocultar, fundir células, congelar painéis e outros itens do dia a dia.
- **Na barra de estado inferior**: escolha que estatísticas a barra de estado mostra (média / contagem / soma, ...); a escolha fica guardada.
- **Num separador de folha em baixo**: adicionar / renomear / eliminar / dar cor / ocultar folhas (o menu de separadores do Univer).
- O menu de contexto da barra de guias superior está em [Guias e gestão de janelas](help://tabs-and-windows).

## IA

- O painel de IA lateral: selecione um intervalo e dê instruções em linguagem corrente (reformatar, gerar dados, escrever fórmulas).
- Anexe ficheiros a um pedido com o botão 📎, ou arraste-os para o painel; acompanham a questão e as imagens voltam como miniaturas.
- Uma resposta pode citar uma célula — clique na referência e a grelha salta para lá.
- As alterações feitas pela IA podem ser revertidas a partir do painel.

## Salvar e exportar

- Guardar em .xlsx (as fórmulas e a formatação são preservadas); Salvar Como…; a exportação para PDF segue a paginação de impressão.
- O salvamento automático segue a regra global (liga-se após o primeiro salvamento manual).

## Estabilidade

- O componente de cálculo em Rust está isolado da interface ao nível do processo: se forem dados extremos a matá-lo, recebe uma mensagem e uma tentativa de recuperar a sessão — não uma falha da aplicação.
