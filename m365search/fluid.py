"""Loop / Fluid documents cached by New Teams (meeting notes, facilitator notes, Loop components).

Teams keeps the Loop documents it has open in the `fluidDriverCache` IndexedDB: a snapshot (merge-tree
text per channel at some sequence number, plus catch-up ops) and the ops that came after it. The text is
rebuilt by replaying the ops in sequence order. Concurrent edits are applied in arrival order, which is
exact for the usual single-writer case (the AI facilitator) and close enough for search otherwise.
"""
import json
import re

from . import teams

PARAGRAPH = "\n"


def _json(x):
    if isinstance(x, bytes):  # V8 one-byte strings are Latin-1, not UTF-8
        try:
            x = x.decode("utf-8")
        except UnicodeDecodeError:
            x = x.decode("latin-1")
    return json.loads(x) if isinstance(x, str) else x


class _Text:
    """A merge-tree string as a list of segments: str for text, None for a marker (one position)."""

    def __init__(self, segments):
        self.segs = segments

    def _split(self, pos):
        """Index of the segment that starts at pos, splitting one if needed."""
        at = 0
        for i, s in enumerate(self.segs):
            n = 1 if s is None else len(s)
            if at == pos:
                return i
            if at < pos < at + n:  # markers have length 1, so only text gets here
                self.segs[i:i + 1] = [s[:pos - at], s[pos - at:]]
                return i + 1
            at += n
        return len(self.segs)

    def apply(self, op):
        kind = op.get("type")
        if kind == 3:  # group
            for o in op.get("ops") or []:
                self.apply(o)
        elif kind == 0:  # insert
            seg = op.get("seg")
            if isinstance(seg, str):
                seg = {"text": seg}
            new = [None] if "marker" in (seg or {}) else [seg.get("text", "")] if seg else []
            self.segs[self._split(op.get("pos1", 0)):0] = new
        elif kind == 1:  # remove
            a = self._split(op.get("pos1", 0))
            b = self._split(op.get("pos2", op.get("pos1", 0)))
            del self.segs[a:b]

    def text(self) -> str:
        out = "".join(PARAGRAPH if s is None else s for s in self.segs)
        return re.sub(r"\n{3,}", "\n\n", out).strip()


def _segments(header: dict) -> list:
    segs = []
    for s in header.get("segmentTexts") or header.get("segments") or []:
        if isinstance(s, str):
            segs.append(s)
        elif "marker" in s:
            segs.append(None)
        else:
            segs.append(s.get("text", ""))
    return segs


def _channel_ops(c):
    """Unwrap container op contents down to (component address, channel address, merge-tree op)."""
    c = _json(c)
    if isinstance(c, list):  # grouped batch
        for x in c:
            yield from _channel_ops(x.get("contents") if isinstance(x, dict) and "contents" in x else x)
        return
    if not isinstance(c, dict):
        return
    if c.get("type") == "groupedBatch":
        yield from _channel_ops(c.get("contents"))
        return
    if c.get("type") == "component":
        c = c.get("contents") or {}
    comp, inner = c.get("address"), c.get("contents")
    if not isinstance(inner, dict):
        return
    content = inner.get("content")
    if not isinstance(content, dict):
        return
    chan, mt = content.get("address"), content.get("contents")
    if comp and chan and isinstance(mt, dict):
        yield comp, chan, mt


def _document(snapshot: dict, extra_ops: list[dict]) -> str:
    blobs = snapshot.get("blobContents") or {}

    def blob(bid):
        return _json(blobs.get(bid)) if bid in blobs else None

    texts, done = {}, {}
    channels = (snapshot.get("snapshotTree") or {}).get("trees", {}).get(".app", {}).get("trees", {}) \
        .get(".channels", {}).get("trees", {})
    for comp, tree in channels.items():
        for chan, ct in (tree.get("trees", {}).get(".channels", {}).get("trees", {})).items():
            content = ct.get("trees", {}).get("content", {}).get("blobs", {})
            header = blob(content.get("header")) if content.get("header") else None
            if not header:
                continue
            t = _Text(_segments(header))
            seq = header.get("chunkSequenceNumber", 0)
            catchup = [o for o in blob(content.get("catchupOps")) or [] if isinstance(o, dict)]
            for op in sorted(catchup, key=lambda o: o.get("sequenceNumber", 0)):
                if op.get("sequenceNumber", 0) > seq:
                    t.apply(op.get("contents") or {})
                    seq = op["sequenceNumber"]
            texts[(comp, chan)], done[(comp, chan)] = t, seq
    base = snapshot.get("sequenceNumber", 0)
    seen = set()
    ops = [o for o in list(snapshot.get("ops") or []) + extra_ops if isinstance(o, dict)]
    for op in sorted(ops, key=lambda o: o.get("sequenceNumber", 0)):
        sn = op.get("sequenceNumber", 0)
        if sn <= base or sn in seen:
            continue
        seen.add(sn)
        if op.get("type") not in ("op", "groupedBatch", None):
            continue
        for comp, chan, mt in _channel_ops(op.get("contents")):
            if (comp, chan) in texts and sn > done[(comp, chan)]:
                texts[(comp, chan)].apply(mt)
    # the main text first (largest), then the smaller components (table cells, list items, titles)
    parts = sorted((t.text() for t in texts.values()), key=len, reverse=True)
    return "\n\n".join(p for p in parts if p)


def documents() -> list[dict]:
    """One dict per cached Loop document: id, text, cached_ms, first_ms (earliest edit stamp)."""
    out = []
    for ldb, blob in teams.stores():
        w = teams.idb.WrappedIndexDB(ldb, blob)
        for d in w.database_ids:
            if d.name != "fluidDriverCache":
                continue
            snaps, ops = {}, {}
            for r in w[d.dbid_no]["driverStorage.V3"].iterate_records(errors_to_stdout=False):
                v = teams._v(r.value)
                if not isinstance(v, dict):
                    continue
                key = str(teams._v(getattr(r.key, "value", r.key)))
                doc, _, kind = key.partition("_")
                try:
                    obj = _json(teams._v(v.get("cachedObject")))
                except (ValueError, TypeError):
                    continue
                value = _json(obj.get("value")) if isinstance(obj, dict) else None
                created = teams._v(v.get("createdTimeMs")) or 0
                if kind.startswith("snapshot") and isinstance(value, dict):
                    if created >= snaps.get(doc, (0, None))[0]:
                        snaps[doc] = (created, value)
                elif kind.startswith("ops") and isinstance(value, list):
                    ops.setdefault(doc, []).extend(value)
            for doc, (created, snap) in snaps.items():
                try:
                    text = _document(snap, ops.get(doc, []))
                except Exception as e:  # one odd document must not stop the build
                    print(f"  ! fluid {doc[:12]}: {type(e).__name__}: {e}")
                    continue
                stamps = [int(t) for t in re.findall(r'"timestamp\\?":(\d{13})', json.dumps(snap, default=str)[:2_000_000])]
                if text:
                    out.append(dict(id=doc, text=text, cached_ms=created, first_ms=min(stamps) if stamps else created))
    return out
