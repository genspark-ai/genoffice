#!/bin/sh
# deb/rpm post-install. Replaces electron-builder's default template, so its
# chrome-sandbox and AppArmor steps are repeated here (genoffice#2150).
set -e
app="/opt/GenOffice"
# Chromium aborts when the SUID helper exists without 4755; Chrome's deb sets it unconditionally.
[ -f "$app/chrome-sandbox" ] && { chown root:root "$app/chrome-sandbox" 2>/dev/null || true; chmod 4755 "$app/chrome-sandbox" 2>/dev/null || true; }
# Ubuntu 24.04+ restricts unprivileged userns; --skip-kernel-load doubles as the abi/4.0 check.
profile_src="$app/resources/apparmor-profile"
profile_dst="/etc/apparmor.d/genoffice"
if [ -f "$profile_src" ] && apparmor_status --enabled >/dev/null 2>&1; then
  if apparmor_parser --skip-kernel-load --debug "$profile_src" >/dev/null 2>&1; then
    cp -f "$profile_src" "$profile_dst"
    if ! { [ -x /usr/bin/ischroot ] && /usr/bin/ischroot; } && command -v apparmor_parser >/dev/null 2>&1; then
      apparmor_parser --replace --write-cache --skip-read-cache "$profile_dst" || true
    fi
  fi
fi
# electron-builder stamps every MIME type it registers with
# <icon name="x-office-document"/>, which overrides the system's per-type
# icons (spreadsheet, presentation, text, pdf...) for all 21 types we claim
# (genoffice#2126). Dropping the line lets each type fall back to its
# freedesktop generic-icon; the glob/MimeType associations are untouched.
mime="/usr/share/mime/packages/genoffice.xml"
if [ -f "$mime" ] && grep -q '<icon name="x-office-document"' "$mime"; then
  sed -i '/<icon name="x-office-document"/d' "$mime"
fi
command -v update-mime-database >/dev/null 2>&1 && update-mime-database /usr/share/mime >/dev/null 2>&1 || true
command -v update-desktop-database >/dev/null 2>&1 && update-desktop-database /usr/share/applications >/dev/null 2>&1 || true
# Expose the genoffice command line shipped inside the app. Only an absent
# name, a dead link or a link into our own install dir is taken over; anything
# else at /usr/bin/genoffice belongs to another program (genoffice#893).
launcher="$app/resources/cli/genoffice"
link="/usr/bin/genoffice"
[ -x "$launcher" ] || exit 0
if [ -L "$link" ]; then
  case "$(readlink "$link")" in
    /opt/GenOffice/*) ;;
    *) [ -e "$link" ] && { echo "genoffice: $link is another program, left as is; run: ln -s $launcher $link" >&2; exit 0; } ;;
  esac
elif [ -e "$link" ]; then
  echo "genoffice: $link is another program, left as is; run: ln -s $launcher $link" >&2
  exit 0
fi
ln -sfn "$launcher" "$link"
exit 0
