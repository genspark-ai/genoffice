# Configurações, idioma, tema e integrações MCP

## Abrir as configurações

A linha da conta no canto inferior esquerdo do ecrã Início abre o painel de configurações (mostra Entrar quando não tem sessão iniciada); e as opções relacionadas com a IA ficam na sua secção Modelo de IA.

![A janela de configurações](img/settings-integrations.png) — a configuração dos modelos está no capítulo Modelos de IA e configurações.

## Idioma

- As configurações oferecem **21 idiomas de interface**: inglês, chinês simplificado, japonês, coreano, francês, alemão, espanhol, tailandês, indonésio, russo, árabe, português, italiano, polaco, tcheco, neerlandês, malaio, hebraico, híndi, chinês tradicional e vietnamita.
- A troca aplica-se de imediato e persiste; e a barra de menus do sistema é reconstruída com o novo idioma.

## Tema

Claro / Escuro / Seguir o Sistema. O modo Seguir o Sistema acompanha o aspeto do sistema operativo, e os editores mudam de aspeto em sincronia sem cintilação.

## Associações de aplicações predefinidas

As Configurações podem registar o GenOffice como a aplicação que abre ficheiros .docx / .xlsx / .pptx / .pdf e semelhantes (registo de aplicação predefinida ao nível do sistema; confirme quando for solicitado).

## Avisos de software de terceiros e atualizações

- Ajuda ▸ Avisos de software de terceiros: o inventário completo das licenças de código aberto que acompanha a aplicação.
- Ajuda ▸ Procurar atualizações…: inicia uma verificação manual; e se houver uma versão mais recente, o GenOffice propõe a instalação.

## Entrar na Genspark

- O acesso (nas Configurações ou na lista de projetos na nuvem) usa um fluxo de **código de dispositivo**: o GenOffice mostra um código e abre o início de sessão no navegador; e a partir daí o processo prossegue sozinho.
- O início de sessão serve apenas para: a lista de projetos na nuvem e os modelos alojados da Genspark. Sem ele, todas as funcionalidades locais e os modelos personalizados continuam a funcionar.
- Terminar sessão é um clique nas Configurações.

## Integração MCP (para utilizadores avançados / clientes de IA)

O GenOffice integra um **servidor MCP** local para que clientes de IA externos (Claude Desktop, Cursor, ...) possam ler e escrever diretamente os seus documentos:

- Arranque: `genoffice mcp` na linha de comandos (a porta e o token de autenticação são configuráveis; só loopback por omissão).
- Capacidades: criar/abrir/editar .docx, .xlsx e .pptx, ler conteúdos, converter formatos, exportar PDF e mais — o mesmo conjunto de ferramentas que usam as aplicações de ambiente de trabalho.
- Segurança: a autenticação por token é opcional mas recomendada; e o ouvinte fica na máquina local por omissão; consulte `genoffice mcp --help`.

## Folha de referência da linha de comandos

| Comando            | O que faz                    |
| ------------------ | ---------------------------- |
| `genoffice <file>` | abrir um ficheiro            |
| `genoffice mcp`    | iniciar o servidor MCP local |
| `genoffice --help` | todos os comandos e opções   |
