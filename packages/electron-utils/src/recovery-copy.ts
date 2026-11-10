/**
 * A crash-recovery copy is offered whenever it still exists on open: a newer
 * original is not proof the copy is obsolete (the file may have been edited
 * elsewhere after the crash), it only changes the prompt wording.
 */
export function recoveryCopyDiskNewer(copyMtimeMs: number, originalMtimeMs: number): boolean {
  return originalMtimeMs >= copyMtimeMs
}
