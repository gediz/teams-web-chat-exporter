# Bundled PDF fonts

These fonts are embedded (subsetted at runtime via HarfBuzz) into exported
PDFs. They are full, unsubsetted Noto faces shipped as TTF, under the SIL
Open Font License 1.1 (see `OFL.txt`).

| File | Coverage |
|------|----------|
| `NotoSans-Regular.ttf` | Latin, Cyrillic, Greek, Hebrew, basic Arabic |
| `NotoSans-Bold.ttf` | as above, bold |
| `NotoSansSC-Regular.ttf` | Han ideographs (all common: URO + Ext A) + Hiragana + Katakana |
| `NotoSansKR-Regular.ttf` | full Noto Sans KR (Hangul + Hanja + kana + Latin) |

`pickFontForChar` in `src/background/pdf.ts` routes each codepoint to one of
these: Hangul/Jamo to KR, Han and kana to SC, everything else to Noto Sans.

## Why TTF and not WOFF2

WOFF2 would cut the bundle by ~3 MB at identical coverage, but the service
worker has no native WOFF2 decoder, and the only viable JS decoder (wawoff2)
relies on dynamic code evaluation that the MV3 extension CSP
(`script-src 'self' 'wasm-unsafe-eval'`) blocks. It throws at worker load and
fails service-worker registration, so WOFF2 is not usable here. The TTFs are
compressed by the store's package zip anyway (the CJK font ships at ~6 MB).

## Provenance / regenerating

All four are Noto (OFL), from github.com/google/fonts. Noto Sans KR is the
variable font instantiated to a static Regular:

```
fonttools varLib.instancer "NotoSansKR[wght].ttf" wght=400 -o NotoSansKR-Regular.ttf
```
