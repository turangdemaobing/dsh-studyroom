# -*- coding: utf-8 -*-
"""Minimal BIFF8 (.xls) reader: OLE compound file + records. Stdlib only."""
import struct


class CFB:
    def __init__(self, data):
        if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("not an OLE compound file")
        self.data = data
        self.ss = 1 << struct.unpack_from("<H", data, 30)[0]
        self.mss = 1 << struct.unpack_from("<H", data, 32)[0]
        num_fat = struct.unpack_from("<I", data, 44)[0]
        dir_start = struct.unpack_from("<I", data, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 56)[0]
        minifat_start = struct.unpack_from("<I", data, 60)[0]
        difat_start = struct.unpack_from("<I", data, 68)[0]
        num_difat = struct.unpack_from("<I", data, 72)[0]
        difat = list(struct.unpack_from("<109I", data, 76))
        sect, n = difat_start, num_difat
        while sect < 0xFFFFFFFA and n > 0:
            vals = struct.unpack_from("<%dI" % (self.ss // 4), data, 512 + sect * self.ss)
            difat.extend(vals[:-1])
            sect = vals[-1]
            n -= 1
        fat = []
        for s in [x for x in difat if x < 0xFFFFFFFA][:num_fat]:
            fat.extend(struct.unpack_from("<%dI" % (self.ss // 4), data, 512 + s * self.ss))
        self.fat = fat
        self.minifat = []
        if minifat_start < 0xFFFFFFFA:
            raw = self._chain_bytes(minifat_start, None)
            self.minifat = list(struct.unpack_from("<%dI" % (len(raw) // 4), raw, 0)) if raw else []
        self.entries = {}
        self.root = None
        self._read_dir(dir_start)

    def _chain(self, start, table):
        out, s = [], start
        while s < 0xFFFFFFFA and len(out) < 1 << 20:
            out.append(s)
            s = table[s] if s < len(table) else 0xFFFFFFFE
        return out

    def _chain_bytes(self, start, size, mini=False):
        if mini:
            chain = self._chain(start, self.minifat)
            out = b"".join(self.ministream[s * self.mss:(s + 1) * self.mss] for s in chain)
        else:
            chain = self._chain(start, self.fat)
            out = b"".join(self.data[512 + s * self.ss:512 + (s + 1) * self.ss] for s in chain)
        return out[:size] if size is not None else out

    def _read_dir(self, dir_start):
        raw = self._chain_bytes(dir_start, None)
        for i in range(len(raw) // 128):
            e = raw[i * 128:(i + 1) * 128]
            name_len = struct.unpack_from("<H", e, 64)[0]
            if not name_len:
                continue
            name = e[:max(0, name_len - 2)].decode("utf-16-le", "ignore")
            etype = e[66]
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<I", e, 120)[0]
            rec = {"name": name, "type": etype, "start": start, "size": size}
            if etype == 5:
                self.root = rec
            else:
                self.entries[name] = rec
        if self.root and self.root["start"] < 0xFFFFFFFA:
            self.ministream = self._chain_bytes(self.root["start"], self.root["size"])

    def stream(self, name):
        e = self.entries[name]
        if e["size"] < self.mini_cutoff:
            return self._chain_bytes(e["start"], e["size"], mini=True)
        return self._chain_bytes(e["start"], e["size"])


def _rk(v):
    if v & 0x02:
        num = float(struct.unpack("<i", struct.pack("<I", v & 0xFFFFFFFC))[0] >> 2)
    else:
        num = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return num / 100.0 if v & 0x01 else num


def _xlunicode(buf, pos):
    cch = struct.unpack_from("<H", buf, pos)[0]
    grbit = buf[pos + 2]
    pos += 3
    rich = struct.unpack_from("<H", buf, pos)[0] if grbit & 0x08 else 0
    if grbit & 0x08:
        pos += 2
    ext = struct.unpack_from("<I", buf, pos)[0] if grbit & 0x04 else 0
    if grbit & 0x04:
        pos += 4
    if grbit & 0x01:
        s = buf[pos:pos + cch * 2].decode("utf-16-le", "ignore")
        pos += cch * 2
    else:
        s = buf[pos:pos + cch].decode("latin-1")
        pos += cch
    pos += rich * 4 + ext
    return s, pos


class _SSTR:
    """Shared string table reader that understands CONTINUE boundaries."""

    def __init__(self, segments):
        self.segs = segments
        self.si = 0
        self.off = 0

    def _avail(self):
        return len(self.segs[self.si]) - self.off

    def read(self, n):
        out = b""
        while n > 0:
            if self._avail() == 0:
                self.si += 1
                self.off = 0
            take = min(n, self._avail())
            out += self.segs[self.si][self.off:self.off + take]
            self.off += take
            n -= take
        return out

    def at_seg_start(self):
        return self.off == 0

    def next_seg(self):
        self.si += 1
        self.off = 0

    def parse_string(self):
        cch = struct.unpack("<H", self.read(2))[0]
        grbit = self.read(1)[0]
        rich = struct.unpack("<H", self.read(2))[0] if grbit & 0x08 else 0
        ext = struct.unpack("<I", self.read(4))[0] if grbit & 0x04 else 0
        wide = bool(grbit & 0x01)
        chars = []
        remaining = cch
        while remaining > 0:
            if self._avail() == 0:
                self.next_seg()
                wide = bool(self.read(1)[0] & 0x01)
            avail_chars = self._avail() // (2 if wide else 1)
            if avail_chars == 0:
                self.next_seg()
                wide = bool(self.read(1)[0] & 0x01)
                continue
            take = min(remaining, avail_chars)
            raw = self.read(take * (2 if wide else 1))
            chars.append(raw.decode("utf-16-le", "ignore") if wide else raw.decode("latin-1"))
            remaining -= take
        if rich:
            self.read(rich * 4)
        if ext:
            self.read(ext)
        return "".join(chars)


def read_xls(path):
    data = open(path, "rb").read()
    cfb = CFB(data)
    wb = cfb.stream("Workbook" if "Workbook" in cfb.entries else "Book")

    recs, pos = [], 0
    while pos + 4 <= len(wb):
        rt, rl = struct.unpack_from("<HH", wb, pos)
        body = wb[pos + 4:pos + 4 + rl]
        recs.append((pos, rt, body))
        pos += 4 + rl

    sheets, sst, sst_segments = [], [], None
    for i, (p, rt, body) in enumerate(recs):
        if rt == 0x0085:
            bof_pos = struct.unpack_from("<I", body, 0)[0]
            cch = body[6]
            wide = body[7] & 0x01
            name = (body[8:8 + cch * 2].decode("utf-16-le", "ignore") if wide
                    else body[8:8 + cch].decode("latin-1"))
            sheets.append({"name": name, "pos": bof_pos})
        elif rt == 0x00FC:
            segs = [body]
            j = i + 1
            while j < len(recs) and recs[j][1] == 0x003C:
                segs.append(recs[j][2])
                j += 1
            total, unique = struct.unpack_from("<II", body, 0)
            r = _SSTR(segs)
            r.read(8)
            for _ in range(unique):
                sst.append(r.parse_string())
            break

    def cells_at(start):
        out = {}
        k = next(i for i, r in enumerate(recs) if r[0] == start)
        depth = 0
        while k < len(recs):
            p, rt, body = recs[k]
            k += 1
            if rt == 0x0809:
                depth += 1
                if depth > 1:
                    continue
            if rt == 0x000A:
                break
            try:
                if rt == 0x00FD:
                    row, col, _xf, isst = struct.unpack_from("<HHHI", body, 0)
                    out[(row, col)] = sst[isst] if isst < len(sst) else ""
                elif rt == 0x0204:
                    row, col = struct.unpack_from("<HH", body, 0)
                    s, _ = _xlunicode(body, 6)
                    out[(row, col)] = s
                elif rt == 0x0203:
                    row, col = struct.unpack_from("<HH", body, 0)
                    out[(row, col)] = struct.unpack_from("<d", body, 6)[0]
                elif rt == 0x027E:
                    row, col = struct.unpack_from("<HH", body, 0)
                    out[(row, col)] = _rk(struct.unpack_from("<I", body, 6)[0])
                elif rt == 0x00BD:
                    row = struct.unpack_from("<H", body, 0)[0]
                    c0 = struct.unpack_from("<H", body, 2)[0]
                    n = (len(body) - 6) // 6
                    for x in range(n):
                        out[(row, c0 + x)] = _rk(struct.unpack_from("<I", body, 4 + x * 6 + 2)[0])
                elif rt == 0x0006:
                    row, col = struct.unpack_from("<HH", body, 0)
                    res = body[6:14]
                    if res[6:8] == b"\xff\xff":
                        if res[0] == 0:
                            nxt = recs[k] if k < len(recs) else None
                            if nxt and nxt[1] == 0x0207:
                                s, _ = _xlunicode(nxt[2], 0)
                                out[(row, col)] = s
                    else:
                        out[(row, col)] = struct.unpack("<d", res)[0]
            except Exception:
                pass
        return out

    return [(s["name"], cells_at(s["pos"])) for s in sheets]


if __name__ == "__main__":
    import json
    import sys
    for name, cells in read_xls(sys.argv[1]):
        print("=== SHEET:", name, "cells:", len(cells))
        if not cells:
            continue
        maxr = max(r for r, c in cells)
        maxc = max(c for r, c in cells)
        for r in range(maxr + 1):
            row = [cells.get((r, c), "") for c in range(maxc + 1)]
            if any(str(x).strip() for x in row):
                print(r, json.dumps([str(x) for x in row], ensure_ascii=False))
