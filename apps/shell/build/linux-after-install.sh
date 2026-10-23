#!/bin/sh
# deb/rpm post-install: expose the genoffice command line shipped inside the app.
# Only an absent name, a dead link or a link into our own install dir is taken
# over; anything else at /usr/bin/genoffice belongs to another program (genoffice#893).
set -e
# electron-builder stamps every MIME type it registers with
# <icon name="x-office-document"/>, which overrides the system's per-type
# icons (spreadsheet, presentation, text, pdf...) for all 21 types we claim
# (genoffice#2126). Dropping the line lets each type fall back to its
# freedesktop generic-icon; the glob/MimeType associations are untouched.
mime="/usr/share/mime/packages/genoffice.xml"
if [ -f "$mime" ] && grep -q '<icon name="x-office-document"' "$mime"; then
  sed -i '/<icon name="x-office-document"/d' "$mime"
  command -v update-mime-database >/dev/null 2>&1 && update-mime-database /usr/share/mime >/dev/null 2>&1 || true
fi
launcher="/opt/GenOffice/resources/cli/genoffice"
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
