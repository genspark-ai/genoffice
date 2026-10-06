# Instalar GenOffice

Cada versión publicada está en la [página de versiones](https://github.com/genspark-ai/genoffice/releases/latest). Todas salen de `main`; los instaladores de macOS y Windows están firmados.

## Elija el archivo para su equipo

| Plataforma | Requisitos | Archivo |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — la mayoría de los PC | Windows 10+, Intel/AMD | instalador `-x64.exe` |
| **Windows** en Arm | Windows 11 en Arm (Snapdragon X y similares) | instalador `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 o posterior) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — cualquier otro | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Las versiones anteriores siguen en la página de versiones.

## En macOS

Abra el `.dmg` y arrastre **GenOffice** a Aplicaciones. En una compilación para Intel, el primer arranque le pide confirmar la apertura: haga clic derecho en la aplicación dentro de Aplicaciones ▸ Abrir. Es Gatekeeper:topándose con una imagen de disco que parece no estar firmada, no una descarga dañada.

## En Windows

Ejecute el instalador `.exe`. Coloca GenOffice en el menú Inicio y registra `.docx`, `.xlsx`, `.pptx`, `.pdf` y formatos similares como aplicaciones predeterminadas, de modo que al hacer doble clic en un archivo este se abra en GenOffice.

## En Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — se ejecuta donde esté, sin ningún paso de instalación:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Necesita el entorno de ejecución de FUSE 2. Si prefiere no instalarlo, ejecútelo con `--appimage-extract-and-run` y el sandbox funcionará desde el directorio extraído.

## Tres rutas más

Las tres envuelven los artefactos publicados en lugar de recompilar la aplicación, así que van a juego con los binarios oficiales.

- **Flatpak** — instala el `.deb` a través del mecanismo `extra-data` de Flatpak y lo ejecuta bajo Zypak, que adapta el propio sandbox de Electron. El árbol `/app` es de solo lectura, así que las actualizaciones pasan por Flatpak y no por el diálogo de actualización de la aplicación.
- **Nix** — una expresión de paquete que usted compila desde una copia local del repositorio con `nix-build packaging/nix`; véase `packaging/nix`.
- **Docker** — una imagen de conversión por lotes sin interfaz (`packaging/docker`). Convierte un árbol entero de documentos de Office, Markdown y HTML a PDF mediante la propia canalización de exportación de la aplicación, así que la salida es la misma que produciría Archivo ▸ Exportar. Es una utilidad de línea de comandos, no un servidor.

## La línea de comandos viene incluida

Cada instalación trae un comando `genoffice` que habla con los mismos motores que la ventana, de modo que ambos nunca discrepan sobre un archivo. En macOS y Windows está dentro del paquete de la aplicación; `genoffice install-cli` lo pone en su `PATH`. Vea **Línea de comandos y agentes** para saber qué puede hacer.
