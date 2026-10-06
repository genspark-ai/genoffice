{
  description = "GenOffice - AI-native office suite";

  inputs.nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";

  outputs =
    { self, nixpkgs }:
    let
      forAllSystems = nixpkgs.lib.genAttrs [
        "x86_64-linux"
        "aarch64-linux"
      ];
    in
    {
      packages = forAllSystems (
        system:
        let
          pkgs = nixpkgs.legacyPackages.${system};
          nodejs = pkgs.nodejs_22;
          lock = pkgs.lib.importJSON ./package-lock.json;
          electron = pkgs."electron_${pkgs.lib.versions.major lock.packages."node_modules/electron".version}";
          fonts = pkgs.symlinkJoin {
            name = "fonts";
            paths = [
              pkgs.dejavu_fonts
              pkgs.liberation_ttf
            ];
          };
        in
        {
          default = pkgs.buildNpmPackage {
            pname = "genoffice";
            version = (pkgs.lib.importJSON ./apps/shell/package.json).version;
            src = self;
            inherit nodejs;
            npmDepsHash = "sha256-6M0PPZ1ox0YNiG86HSAkSDp/HrNf00LwiudPM2YKCq0=";

            cargoRoot = "apps/sheets/native/xlsx-engine";
            cargoDeps = pkgs.rustPlatform.importCargoLock {
              lockFile = ./apps/sheets/native/xlsx-engine/Cargo.lock;
            };

            nativeBuildInputs = with pkgs; [
              cargo
              rustc
              rustPlatform.cargoSetupHook
              makeWrapper
            ];
            nativeCheckInputs = with pkgs; [
              git
              libxml2
            ];

            npmRebuildFlags = [ "--ignore-scripts" ];
            env = {
              ELECTRON_OVERRIDE_DIST_PATH = "${electron}/bin";
              CHROME_PATH = "${pkgs.chromium}/bin/chromium";
              FONTCONFIG_FILE = pkgs.makeFontsConf { fontDirectories = [ fonts ]; };
            };

            npmBuildScript = "build:all";
            postBuild = ''
              npm run notices
            '';

            doCheck = true;
            checkPhase = ''
              runHook preCheck
              (
                unset CARGO_HOME GIT_CONFIG_GLOBAL GIT_CONFIG_SYSTEM XDG_CACHE_HOME XDG_CONFIG_HOME XDG_DATA_HOME XDG_STATE_HOME
                HOME=$(mktemp -d) || exit
                export HOME
                trap 'chmod -R u+w "$HOME"; rm -rf "$HOME"' EXIT
                trap 'exit 130' INT
                trap 'exit 143' TERM
                trap 'exit 129' HUP
                gitvars=$(git rev-parse --local-env-vars) || exit
                unset $gitvars
                root=$(pwd -P) || exit
                isroot() { grep -qs '"name": "@genoffice/shell"' "$1/apps/shell/package.json"; }
                while ! isroot "$root" && [ "$root" != / ]; do root=$(dirname "$root"); done
                isroot "$root" || { echo "checkPhase: not inside a genoffice source tree" >&2; exit 1; }
                cd "$root" || exit
                [ -e .git ] || [ -L .git ] || { git init -q && git add -A; } || exit
                ln -s ${fonts}/share/fonts "$HOME/.fonts" || exit
                npm run fixtures -w @genoffice/sheets || exit
                export CI=1
                export VITEST_MAX_WORKERS=$((NIX_BUILD_CORES < 4 ? NIX_BUILD_CORES : 4))
                workspaces=$(node -p 'require("./package.json").scripts.test.match(/@genoffice\/[\w-]+/g).join(" ")') || exit
                for w in $workspaces; do
                  npm run test -w $w -- --testTimeout=120000 --hookTimeout=120000 || exit
                done
              ) && runHook postCheck || return $?
            '';

            installPhase = ''
              runHook preInstall
              mkdir -p "$out/lib"
              cp -r . "$out/lib/genoffice"
              (
                cd "$out/lib/genoffice" || exit 1
                npm prune --omit=dev
                sidecar=apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar
                mv "$sidecar" "$TMPDIR/xlsx-sidecar"
                rm -rf .git .task apps/sheets/native/xlsx-engine/target apps/sheets/fixtures/generated \
                  node_modules/.vite-temp {apps,packages}/*/node_modules/.vite
                find apps packages -maxdepth 2 -type d -empty -delete
                install -Dm755 "$TMPDIR/xlsx-sidecar" "$sidecar"
              )
              makeWrapper ${electron}/bin/electron "$out/bin/genoffice-app" \
                --add-flags "$out/lib/genoffice/apps/shell"
              cli="$out/lib/genoffice/packages/cli"
              makeWrapper ${nodejs}/bin/node "$cli/bin/genoffice" \
                --add-flags "$cli/dist/genoffice.cjs" \
                --set-default GENOFFICE_APP_BIN "$out/bin/genoffice-app"
              ln -s "$cli/bin/genoffice" "$out/bin/genoffice"
              runHook postInstall
            '';

            meta = {
              description = "AI-native office suite (docs, sheets, slides, pdf, markdown, html)";
              license = pkgs.lib.licenses.asl20;
              platforms = pkgs.lib.platforms.linux;
              mainProgram = "genoffice";
            };
          };
        }
      );

      checks = self.packages;

      apps = forAllSystems (system: {
        genoffice-app = {
          type = "app";
          program = "${self.packages.${system}.default}/bin/genoffice-app";
          meta.description = "GenOffice desktop app";
        };
      });
    };
}
