# Vendored: MicroType Express (MTX) decoder

TypeScript port of the decompression half of
[libeot](https://github.com/umanwizard/libeot) (Brennan T. Vincent, MPL-2.0;
see `LICENSE`). libeot relicensed Monotype's original MTX reference code under
the MPL and relies on the royalty-free patent grants Microsoft and Monotype
attached to the [W3C EOT submission](http://www.w3.org/Submission/2008/01/).

What it does: an EOT payload with `TTEMBED_TTCOMPRESSED` (flag 0x4) holds three
LZ-compressed streams (the CTF table container, glyph push data, glyph
instructions). `lzcomp.ts` inflates them (adaptive Huffman over an LZ77 stream,
optional run-length layer), `ctf.ts` rebuilds `glyf`/`loca`/`cvt ` from the
compact table format and writes a plain sfnt with fresh checksums.

Ported files: `lzcomp/{ahuff,bitio,lzcomp,liblzcomp}.c` → `lzcomp.ts`;
`ctf/{parseCTF,parseTTF,SFNTContainer}.c`, `triplet_encodings.c`,
`util/stream.c` → `ctf.ts`. Compression, hdmx/VDMX reconstruction and the
size-limited copy window are not ported (unused by the decoder path).
