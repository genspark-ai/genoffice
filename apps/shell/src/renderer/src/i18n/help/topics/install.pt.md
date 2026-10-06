# Instalar o GenOffice

Todas as versões são publicadas na [página de lançamentos](https://github.com/genspark-ai/genoffice/releases/latest). Todas vêm de `main`; e os instaladores para macOS e Windows estão assinados.

## Escolher o ficheiro para a sua máquina

| Plataforma                           | Requisitos                                            | Ficheiro                |
| ------------------------------------ | ----------------------------------------------------- | ----------------------- |
| **macOS** — Apple Silicon            | macOS 11+                                             | `.dmg` (arm64)          |
| **macOS** — Intel                    | macOS 11+                                             | `.dmg` (x64)            |
| **Windows** — a maioria dos PC       | Windows 10+, Intel/AMD                                | instalador `-x64.exe`   |
| **Windows** em Arm                   | Windows 11 em Arm (Snapdragon X e semelhantes)        | instalador `-arm64.exe` |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 ou mais recente)    | `.deb`                  |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm`                  |
| **Linux** — qualquer outra           | x86_64, glibc 2.34+, FUSE 2                           | `.AppImage`             |

As versões mais antigas permanecem na página de lançamentos.

## No macOS

Abra o `.dmg` e arraste o **GenOffice** para Aplicações. Numa compilação Intel, o primeiro arranque pede-lhe que confirme a abertura: clique com o botão direito na aplicação em Aplicações ▸ Abrir. Isso é o Gatekeeper a encontrar uma imagem de disco sem assinatura aparente, não uma transferência danificada.

## No Windows

Execute o instalador `.exe`. Ele coloca o GenOffice no menu Iniciar e regista `.docx`, `.xlsx`, `.pptx`, `.pdf` e afins, para que um clique duplo num ficheiro o abra.

## No Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — corre de onde está, sem passo de instalação:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Precisa do runtime FUSE 2. Se preferir não o instalar, execute-o com `--appimage-extract-and-run` e a sandbox funciona a partir do diretório extraído.

## A versão que está a executar

**Configurações ▸ Sobre** é onde a aplicação comunica a versão que tem, o canal que segue e onde vive o projeto. Quando é publicada uma versão mais recente nesse canal, o mesmo painel oferece a atualização; uma instalação por Flatpak, Nix ou Docker não o faz, porque essas são substituídas pela ferramenta que as lá colocou.

![Configurações ▸ Sobre, a mostrar a versão instalada, o canal de atualização que a aplicação segue e a ligação GitHub do projeto](img/install.png)

## Mais três vias

As três embrulham os artefactos publicados em vez de recompilar a aplicação, por isso seguem os binários oficiais.

- **Flatpak** — instala o `.deb` através do mecanismo `extra-data` do Flatpak e executa-o sob o Zypak, que adapta a sandbox própria do Electron. A árvore `/app` é só de leitura, por isso as atualizações passam pelo Flatpak e não pela janela de atualização da aplicação.
- **Nix** — uma expressão de pacote que compila a partir de uma cópia do código com `nix-build packaging/nix`; consulte `packaging/nix`.
- **Docker** — uma imagem de conversão em lote sem interface (`packaging/docker`). Converte uma árvore inteira de documentos Office, Markdown e HTML para PDF através da própria cadeia de exportação da aplicação, por isso o resultado é exatamente o que Arquivo ▸ Exportar produziria. É um utilitário CLI, não um servidor.

## A linha de comandos vem de brinde

Todas as instalações incluem um comando `genoffice` que fala com os mesmos motores que a janela usa, por isso os dois nunca discordam sobre um ficheiro. Em macOS e Windows está dentro do pacote da aplicação; e `genoffice install-cli` coloca-o no seu `PATH`. Consulte **Linha de comandos e agentes** para saber o que ele sabe fazer.
