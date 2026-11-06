#!/bin/sh
# GenOffice Flatpak launcher (packaging/flatpak, genoffice#1858). The first
# launch unpacks the extra-data deb into /app/extra/app, then Zypak runs the
# bundled Electron so its own sandbox works inside the Flatpak sandbox.
# ELECTRON_OZONE_PLATFORM_HINT lets Electron pick Wayland when the session
# offers it and fall back to X11 otherwise.
if [ ! -x /app/extra/app/opt/GenOffice/genoffice ]; then
  /app/bin/apply_extra || exit 1
fi
export ELECTRON_OZONE_PLATFORM_HINT="${ELECTRON_OZONE_PLATFORM_HINT:-auto}"
exec /app/bin/zypak-wrapper.sh /app/extra/app/opt/GenOffice/genoffice "$@"
