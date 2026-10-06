# O editor HTML

O editor HTML abre .html / .htm com dois modos: **Visualizar** (a página renderizada) e **Código**.

- **Visualizar**: renderização a sério; as folhas de estilo e imagens relativas são carregadas a partir de junto do ficheiro.
- **Inspetor de visualização**: clique para selecionar um elemento, clique duas vezes para editar o respetivo texto no local, elimine pela barra de ferramentas e pergunte à IA sobre a seleção.
- **Modo de código**: edita o HTML; ctrl+F procura e Substituir guarda a marcação reescrita.
- **Salvar**: fiel byte a byte (BOM/CRLF/quebra de linha final preservados); um salvamento sem alterações não reescreve.
- **Zoom**: ctrl+roda/pinça dimensiona a pré-visualização; ctrl+Z dentro da pré-visualização desfaz a última edição.

## A barra de ferramentas

Clique num elemento da pré-visualização e uma barra de ferramentas flutua por cima dele:

![A barra de ferramentas flutuante sobre um elemento selecionado](img/html-toolbar.png)

- **Ficheiro e histórico**: Salvar, Salvar Como…, Desfazer, Refazer, Localizar; o interruptor de **Salvamento automático** grava as alterações periodicamente.
- Interruptor **Visualizar / Código**; **Apresentar** mostra a página em ecrã inteiro.
- **Formatação**: negrito, itálico, aumentar/diminuir o tamanho da fonte; e o **painel de estilo** do elemento selecionado (cores e mais).
- **Inserir**: título, parágrafo, tabela, imagem (por ligação), Mais.
- **Ações de imagem** (com uma imagem selecionada): cortar, **remover o plano de fundo**, substituir, bloquear a proporção.
- **Ações de elemento** (com um elemento selecionado no inspetor de visualização): eliminar, duplicar, mover para cima/baixo.
- **Botão de IA**: abre o painel de IA; pergunte diretamente sobre o elemento selecionado.

## Exportação

Menu Arquivo, tudo local e tudo pergunta onde colocar o resultado:

- **Exportar como Word…** e **Exportar como PDF…** escrevem um .docx ou .pdf a sério.
- **Exportar como HTML de arquivo único…** escreve um único .html com as imagens incorporadas. Não substitui o ficheiro que tem aberto neste momento e indica-lhe quantas imagens não conseguiu incorporar.

## Inserir esqueleto

Para uma página em branco, **Inserir ▸ Inserir esqueleto** escreve um documento mínimo em modo de normas:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Cada parte está lá por uma razão, e é por isso que isto é um comando e não algo que se escreva à mão:

- o **doctype**, ou a pré-visualização corre em modo de quirks, onde as dimensões das caixas e a disposição das tabelas seguem regras diferentes das que espera;
- o **`lang`**, ou um leitor de ecrã fica sem nenhuma linguagem para ler a página e o navegador escolhe uma fonte e um corretor ortográfico para a linguagem errada;
- o **charset**, ou uma página de texto não latino pode descodificar-se como texto ilegível.

A meta de viewport está deliberadamente ausente: isto é renderizado num painel de ambiente de trabalho, sem nenhuma viewport móvel para que afete.

O `lang` segue o idioma da interface da aplicação, por isso o esqueleto que insere é aquele para o qual as suas ferramentas já estão preparadas. Edite-o livremente depois.

O item só aparece no modo de edição, e apenas enquanto o documento está vazio — não há nada para inserir um esqueleto _dentro_ quando já há conteúdo.
