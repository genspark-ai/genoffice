#!/bin/bash
# End-to-end marketplace flow test
# Verifies: upload → search → install → pi loader picks up → uninstall → cleanup
set -e

BASE="http://127.0.0.1:18081/api/ipc"
H="-H Content-Type:application/json"

# Use a unique ID so we don't conflict with existing data
ID="e2e-w34-verify-$RANDOM"
echo "=== E2E marketplace flow test: id=$ID ==="

echo ""
echo "[1/7] Upload new skill"
curl -s -X POST "$BASE/home:marketplace-upload" $H \
  -d "{\"args\":[{\"kind\":\"skill\",\"payload\":{
    \"id\":\"$ID\",
    \"name\":\"E2E W34 Verify\",
    \"description\":\"End-to-end test of marketplace upload → install → pi loader → uninstall flow\",
    \"author\":\"E2E Test\",
    \"version\":\"1.0.0\",
    \"tools\":[\"e2e_test_tool\"],
    \"scopes\":[\"files:read\"],
    \"category\":\"dev\",
    \"tags\":[\"e2e\",\"test\"]
  }}]}" | python3 -c "import sys,json; d=json.load(sys.stdin)['result']; print(f'  upload ok={d.get(\"ok\")} reviewStatus={d.get(\"reviewStatus\")}')"

echo ""
echo "[2/7] Search marketplace for our test id"
HITS=$(curl -s -X POST "$BASE/home:marketplace-search" $H \
  -d "{\"args\":[{\"q\":\"$ID\"}]}" \
  | python3 -c "import sys,json; d=json.load(sys.stdin)['result']; print(len(d['skills'])+len(d['plugins']))")
echo "  hits: $HITS (expect >= 1)"

echo ""
echo "[3/7] Install via IPC"
curl -s -X POST "$BASE/home:install-skill" $H \
  -d "{\"args\":[{\"id\":\"$ID\"}]}" \
  | python3 -c "import sys,json; d=json.load(sys.stdin)['result']; print(f'  install ok={d.get(\"ok\")} piInstalled={d.get(\"piInstalled\")}')"

echo ""
echo "[4/7] Verify SKILL.md on disk"
DISK_FILE="/tmp/genoffice-data/pi-skills/$ID/SKILL.md"
if [ -f "$DISK_FILE" ]; then
  HAS_NAME=$(grep -c "^name: $ID$" "$DISK_FILE")
  HAS_DESC=$(grep -c "^description:" "$DISK_FILE")
  HAS_DISPLAY=$(grep -c "^display_name: E2E W34 Verify$" "$DISK_FILE")
  echo "  ✓ SKILL.md exists"
  echo "  ✓ has frontmatter 'name: $ID' (pi slug): $HAS_NAME (expect 1)"
  echo "  ✓ has frontmatter 'description:' (required by pi): $HAS_DESC (expect 1)"
  echo "  ✓ has frontmatter 'display_name: E2E W34 Verify' (human name): $HAS_DISPLAY (expect 1)"
else
  echo "  ✗ SKILL.md MISSING"
fi

echo ""
echo "[5/7] Verify pi loader actually picks it up"
curl -s -X POST "$BASE/home:list-pi-skills" $H -d '{"args":[]}' \
  | python3 -c "
import sys, json
d = json.load(sys.stdin)['result']
match = [s for s in d['piSkills'] if s['name'] == '$ID']
diag_match = [x for x in d['diagnostics'] if '$ID' in x.get('path','')]
print(f'  piSkills matched: {len(match)} (expect 1)')
print(f'  diagnostics for our id: {len(diag_match)} (expect 0)')
"

echo ""
echo "[6/7] Uninstall via IPC"
curl -s -X POST "$BASE/home:uninstall-skill" $H \
  -d "{\"args\":[{\"id\":\"$ID\"}]}" \
  | python3 -c "import sys,json; d=json.load(sys.stdin)['result']; print(f'  uninstall ok={d.get(\"ok\")}')"

echo ""
echo "[7/7] Verify cleanup"
if [ -f "$DISK_FILE" ]; then
  echo "  ✗ SKILL.md still on disk (uninstall failed to clean up)"
else
  echo "  ✓ SKILL.md removed from disk"
fi
IN_INDEX=$(grep -l "\"$ID\"" /tmp/genoffice-data/pi-skills/.index.json 2>/dev/null || echo "")
if [ -n "$IN_INDEX" ]; then
  echo "  ✗ $ID still in .index.json"
else
  echo "  ✓ $ID removed from .index.json"
fi
PI_REMAIN=$(curl -s -X POST "$BASE/home:list-pi-skills" $H -d '{"args":[]}' \
  | python3 -c "import sys,json; d=json.load(sys.stdin)['result']; print(len([s for s in d['piSkills'] if s['name'] == '$ID']))")
echo "  pi loader sees our id: $PI_REMAIN (expect 0)"

echo ""
echo "=== E2E marketplace flow test complete ==="
