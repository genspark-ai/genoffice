/**
 * What the session may do with the document's protection. The owner password
 * is the only thing that may change or lift it: the /P limits bind a
 * user-password session (§7.6.3.2), so even one allowed to modify the content
 * must not strip the owner's protection.
 */
export function mayChangeProtection(
  readOnly: boolean,
  lock: { unlockedAsOwner: boolean } | null,
): boolean {
  // a read-only session cannot save at all, so a protection change could never land
  if (readOnly) return false
  // a plain document has no protection to lose: adding one is allowed
  return lock === null || lock.unlockedAsOwner
}
