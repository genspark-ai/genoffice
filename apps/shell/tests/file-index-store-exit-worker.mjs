// Store-worker test double for the unopenable-DB case: exits before answering
// anything, the way a FileIndexStore construction failure kills the real
// worker at startup.
import process from 'node:process'

process.exit(1)
