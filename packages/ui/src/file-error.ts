/**
 * A short, honest form of a file error for a toast.
 *
 * The full Node message (`EBUSY: resource busy or locked, open 'D:\...'`) is
 * both too long for a toast and unreadable. The errno alone (`EBUSY`,
 * `ENOSPC`, `EACCES`) keeps every failure distinguishable — a full disk must
 * not read as "in use" — while staying on one line. Messages without a code
 * are truncated, so serialization failures still say something.
 */
const MAX_MESSAGE = 60

export function shortFileError(error: unknown): string {
  const code = (error as NodeJS.ErrnoException | null)?.code
  if (typeof code === 'string' && code) return code
  const message = error instanceof Error ? error.message : String(error ?? '')
  // An error that crossed IPC is a string: the code survives only as its prefix.
  const errno = /^(?:Error: )?(E[A-Z0-9]{2,}): /.exec(message)
  if (errno?.[1]) return errno[1]
  return message.length > MAX_MESSAGE ? `${message.slice(0, MAX_MESSAGE)}…` : message
}
