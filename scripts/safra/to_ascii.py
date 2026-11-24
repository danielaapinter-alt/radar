"""Converte o fonte para ASCII puro (acentos em strings viram \\uXXXX)."""
import io, sys, tokenize

src = open(sys.argv[1], encoding="utf-8").read()
src = src.replace("# -*- coding: utf-8 -*-\n", "", 1)
lines = src.splitlines(keepends=True)
offs = [0]
for l in lines:
    offs.append(offs[-1] + len(l))
pos = lambda rc: offs[rc[0] - 1] + rc[1]
spans = []
for tok in tokenize.generate_tokens(io.StringIO(src).readline):
    if any(ord(c) > 127 for c in tok.string):
        if tok.type not in (tokenize.STRING, tokenize.FSTRING_MIDDLE):
            raise SystemExit(f"nao-ASCII fora de string, linha {tok.start[0]}: {tok.string!r}")
        spans.append((pos(tok.start), pos(tok.end)))
for a, b in reversed(spans):
    seg = "".join(c if ord(c) < 128 else f"\\u{ord(c):04x}" for c in src[a:b])
    src = src[:a] + seg + src[b:]
assert all(ord(c) < 128 for c in src)
open(sys.argv[2], "w", encoding="ascii", newline="\r\n").write(src)
