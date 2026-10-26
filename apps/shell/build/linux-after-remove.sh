#!/bin/sh
# deb/rpm post-remove: drop the genoffice symlink only on a real uninstall.
# rpm runs the old package's %postun after the new %post during an upgrade
# (with $1 = 1); deb passes "upgrade" there. Removing then would kill the
# link the new version just created.
case "$1" in
  0|remove|purge) ;;
  *) exit 0 ;;
esac
# fpm already removed our MIME file; rebuild the database so the icons we
# overrode come back without a relogin.
command -v update-mime-database >/dev/null 2>&1 && update-mime-database /usr/share/mime >/dev/null 2>&1 || true
profile="/etc/apparmor.d/genoffice"
if [ -f "$profile" ]; then
  if apparmor_status --enabled >/dev/null 2>&1 && ! { [ -x /usr/bin/ischroot ] && /usr/bin/ischroot; } && command -v apparmor_parser >/dev/null 2>&1; then
    apparmor_parser --remove "$profile" >/dev/null 2>&1 || true
  fi
  rm -f "$profile"
fi
link="/usr/bin/genoffice"
# same ownership rule as the post-install: only a link into our install dir is ours
if [ -L "$link" ]; then
  case "$(readlink "$link")" in
    /opt/GenOffice/*) rm -f "$link" ;;
  esac
fi
exit 0
