/**
 * The ceilings a delimited import (a .csv/.tsv opened as a workbook) may
 * reach.
 *
 * The import streams the grid and the worksheet (see csv-import), so what a
 * file costs is the compression of its archive rather than its size in
 * memory. The ceiling is still there — a text file of any size converts into a
 * workbook of up to a million rows — but it is the user's to set, and it
 * starts where the streaming import can afford it.
 */

/**
 * The steps the setting offers, in MB. The ladder tops out at 256 on
 * purpose: the open path still decodes the whole file into one JS string
 * (sheets-main's prepareWorkbookForOpen), and a string caps at
 * 0x1fffffe8 = 536,870,888 characters — a 512 MB ASCII file is already
 * 536,870,912, so a bigger ceiling would admit files that fail inside
 * TextDecoder.decode with a raw RangeError instead of converting.
 */
export const CSV_IMPORT_LIMIT_CHOICES_MB = [64, 128, 256] as const

/** the ceiling a delimited import starts at, in MB */
export const CSV_IMPORT_DEFAULT_LIMIT_MB = 256
