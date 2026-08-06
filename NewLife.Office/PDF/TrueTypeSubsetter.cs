using System.Text;

namespace NewLife.Office.Pdf;

/// <summary>TrueType 字体子集化器（P08，对标 PdfSharp）</summary>
/// <remarks>
/// 解析 TrueType/TTC 字体，仅保留指定 Unicode 字符对应的字形并重建子集 TTF，
/// 可显著减小嵌入字体的 PDF 体积（完整中文字体 ~10MB → 数百 KB）。
/// <para>支持 cmap Format 4（BMP）、简单字形与复合字形（递归收集组件字形并重映射 GlyphID）。</para>
/// </remarks>
public static class TrueTypeSubsetter
{
    /// <summary>生成子集字体</summary>
    /// <param name="fontData">原始字体字节（TTF 或 TTC）</param>
    /// <param name="sfOffset">字体偏移（TTC 内偏移，TTF 为 0）</param>
    /// <param name="unicodes">需保留的 Unicode 码点（BMP 内）</param>
    /// <returns>子集 TTF 字节与新的 Unicode → GlyphID 映射</returns>
    public static (Byte[] Data, Dictionary<UInt16, UInt16> GlyphMap) Subset(
        Byte[] fontData, Int32 sfOffset, ICollection<Int32> unicodes)
    {
        if (fontData == null || fontData.Length == 0) throw new ArgumentNullException(nameof(fontData));

        // 1. 表目录
        var tables = ReadTableDirectory(fontData, sfOffset);

        // 2. cmap（Unicode → 旧 GlyphID）
        var cmap = ParseCmap(fontData, GetTableOffset(tables, "cmap"));

        // 3. 收集用到的 glyph（含复合字形组件，递归）
        var used = new SortedSet<Int32> { 0 }; // .notdef 始终保留
        foreach (var code in unicodes)
        {
            if (code < 0 || code > 0xFFFF) continue;
            if (cmap.TryGetValue((UInt16)code, out var gid) && gid != 0)
                CollectGlyphTree(fontData, tables, gid, used);
        }

        // 4. 旧 GlyphID → 新 GlyphID
        var remap = new Dictionary<Int32, Int32>(used.Count);
        var idx = 0;
        foreach (var g in used) remap[g] = idx++;

        // 5. 重建各表
        var head = SubsetHead(fontData, GetTableOffset(tables, "head"), 1);
        var maxp = SubsetMaxp(fontData, GetTableOffset(tables, "maxp"), (UInt16)remap.Count);
        var hhea = SubsetHhea(fontData, GetTableOffset(tables, "hhea"), (UInt16)remap.Count);
        var hmtx = SubsetHmtx(fontData, tables, remap, (UInt16)remap.Count);
        var (loca, glyf) = SubsetGlyfLoca(fontData, tables, remap);
        var (newCmap, newCmapMap) = SubsetCmap(cmap, remap);
        var name = CopyTable(fontData, GetTableOffset(tables, "name"), GetTableLength(tables, "name"));
        var post = CopyTable(fontData, GetTableOffset(tables, "post"), GetTableLength(tables, "post"));
        var os2 = CopyTable(fontData, GetTableOffset(tables, "OS/2"), GetTableLength(tables, "OS/2"));
        var cvt = CopyTable(fontData, GetTableOffset(tables, "cvt "), GetTableLength(tables, "cvt "));
        var fpgm = CopyTable(fontData, GetTableOffset(tables, "fpgm"), GetTableLength(tables, "fpgm"));
        var prep = CopyTable(fontData, GetTableOffset(tables, "prep"), GetTableLength(tables, "prep"));

        // 6. 组装新 sfnt
        var newTables = new Dictionary<String, Byte[]>(StringComparer.Ordinal)
        {
            ["head"] = head,
            ["hhea"] = hhea,
            ["maxp"] = maxp,
            ["hmtx"] = hmtx,
            ["cmap"] = newCmap,
            ["loca"] = loca,
            ["glyf"] = glyf,
            ["name"] = name,
            ["post"] = post,
        };
        // 可选表（有则保留，无则跳过）
        if (os2.Length > 0) newTables["OS/2"] = os2;
        if (cvt.Length > 0) newTables["cvt "] = cvt;
        if (fpgm.Length > 0) newTables["fpgm"] = fpgm;
        if (prep.Length > 0) newTables["prep"] = prep;

        return (BuildSfnt(newTables), newCmapMap);
    }

    #region 表解析
    private static Dictionary<String, (Int32 Offset, Int32 Length)> ReadTableDirectory(Byte[] data, Int32 sfOffset)
    {
        var result = new Dictionary<String, (Int32, Int32)>(StringComparer.Ordinal);
        if (sfOffset + 12 > data.Length) return result;
        var numTables = ReadU16(data, sfOffset + 4);
        for (var i = 0; i < numTables; i++)
        {
            var pos = sfOffset + 12 + i * 16;
            if (pos + 16 > data.Length) break;
            var tag = Encoding.ASCII.GetString(data, pos, 4);
            var off = (Int32)ReadU32(data, pos + 8);
            var len = (Int32)ReadU32(data, pos + 12);
            result[tag] = (off, len);
        }
        return result;
    }

    private static Int32 GetTableOffset(Dictionary<String, (Int32 Offset, Int32 Length)> tables, String tag)
        => tables.TryGetValue(tag, out var t) ? t.Offset : -1;

    private static Int32 GetTableLength(Dictionary<String, (Int32 Offset, Int32 Length)> tables, String tag)
        => tables.TryGetValue(tag, out var t) ? t.Length : 0;

    /// <summary>解析 cmap Format 4，返回 Unicode → 旧 GlyphID</summary>
    private static Dictionary<UInt16, UInt16> ParseCmap(Byte[] data, Int32 cmapOff)
    {
        var map = new Dictionary<UInt16, UInt16>();
        if (cmapOff < 0 || cmapOff + 4 > data.Length) return map;
        var numSubtables = ReadU16(data, cmapOff + 2);
        var format4 = -1;
        for (var i = 0; i < numSubtables; i++)
        {
            var base2 = cmapOff + 4 + i * 8;
            if (base2 + 8 > data.Length) break;
            var platformId = ReadU16(data, base2);
            var encodingId = ReadU16(data, base2 + 2);
            var subOff = cmapOff + (Int32)ReadU32(data, base2 + 4);
            if (subOff + 2 > data.Length || ReadU16(data, subOff) != 4) continue;
            if (platformId == 3 && encodingId == 1) { format4 = subOff; break; }
            if (platformId == 0 && (encodingId == 3 || encodingId == 4) && format4 < 0) format4 = subOff;
        }
        if (format4 < 0) return map;

        var segCount = ReadU16(data, format4 + 6) / 2;
        var endCodeBase = format4 + 14;
        var startCodeBase = endCodeBase + segCount * 2 + 2;
        var idDeltaBase = startCodeBase + segCount * 2;
        var idRangeOffBase = idDeltaBase + segCount * 2;
        for (var i = 0; i < segCount; i++)
        {
            var endCode = ReadU16(data, endCodeBase + i * 2);
            var startCode = ReadU16(data, startCodeBase + i * 2);
            var idDelta = ReadS16(data, idDeltaBase + i * 2);
            var idRangeOff = ReadU16(data, idRangeOffBase + i * 2);
            if (startCode == 0xFFFF) break;
            for (var code = (UInt32)startCode; code <= endCode; code++)
            {
                UInt16 gid;
                if (idRangeOff == 0)
                {
                    gid = (UInt16)((code + (UInt32)(UInt16)idDelta) & 0xFFFF);
                }
                else
                {
                    var idxPos = idRangeOffBase + i * 2 + idRangeOff + (code - startCode) * 2;
                    if (idxPos + 1 >= data.Length) continue;
                    gid = ReadU16(data, (Int32)idxPos);
                    if (gid != 0) gid = (UInt16)((gid + (UInt32)(UInt16)idDelta) & 0xFFFF);
                }
                if (gid != 0) map[(UInt16)code] = gid;
            }
        }
        return map;
    }

    /// <summary>递归收集字形及其复合组件字形</summary>
    private static void CollectGlyphTree(Byte[] data, Dictionary<String, (Int32 Offset, Int32 Length)> tables,
        Int32 gid, SortedSet<Int32> used)
    {
        if (gid <= 0 || !used.Add(gid)) return;

        var (glyphData, _) = GetGlyphBytes(data, tables, gid);
        if (glyphData.Length < 10) return;
        var numberOfContours = ReadS16(glyphData, 0);
        if (numberOfContours >= 0) return; // 简单字形

        // 复合字形：遍历组件
        var pos = 10;
        while (true)
        {
            if (pos + 4 > glyphData.Length) break;
            var flags = ReadU16(glyphData, pos);
            var compGid = ReadU16(glyphData, pos + 2);
            pos += 4;
            if (compGid > 0) CollectGlyphTree(data, tables, compGid, used);

            // 跳过参数
            pos += (flags & 0x0001) != 0 ? 4 : 2;
            // 跳过变换
            if ((flags & 0x0008) != 0) pos += 2;
            else if ((flags & 0x0040) != 0) pos += 4;
            else if ((flags & 0x0080) != 0) pos += 8;

            if ((flags & 0x0020) == 0) break;
        }
    }

    /// <summary>获取指定 GlyphID 的原始字节（基于 loca/glyf）</summary>
    private static (Byte[] Data, Int32 OldGid) GetGlyphBytes(Byte[] data, Dictionary<String, (Int32 Offset, Int32 Length)> tables, Int32 gid)
    {
        var locaOff = GetTableOffset(tables, "loca");
        var glyfOff = GetTableOffset(tables, "glyf");
        if (locaOff < 0 || glyfOff < 0) return (new Byte[0], gid);

        // indexToLocFormat 在 head 表 +50
        var headOff = GetTableOffset(tables, "head");
        var longFmt = headOff >= 0 && headOff + 52 <= data.Length && ReadS16(data, headOff + 50) == 1;

        Int32 start, end;
        if (longFmt)
        {
            if (locaOff + (gid + 1) * 4 + 4 > data.Length) return (new Byte[0], gid);
            start = (Int32)ReadU32(data, locaOff + gid * 4);
            end = (Int32)ReadU32(data, locaOff + (gid + 1) * 4);
        }
        else
        {
            if (locaOff + (gid + 1) * 2 + 2 > data.Length) return (new Byte[0], gid);
            start = ReadU16(data, locaOff + gid * 2) * 2;
            end = ReadU16(data, locaOff + (gid + 1) * 2) * 2;
        }

        if (start < 0 || end < start || glyfOff + end > data.Length) return (new Byte[0], gid);
        var len = end - start;
        var buf = new Byte[len];
        Array.Copy(data, glyfOff + start, buf, 0, len);
        return (buf, gid);
    }

    #endregion

    #region 表重建

    private static Byte[] SubsetHead(Byte[] data, Int32 headOff, Int32 indexToLocFormat)
    {
        var head = new Byte[54];
        if (headOff >= 0 && headOff + 54 <= data.Length)
            Array.Copy(data, headOff, head, 0, 54);
        else
        {
            // 空 head：写版本/魔法数/unitsPerEm 默认
            WriteU32(head, 0, 0x00010000);
            WriteU32(head, 12, 0x5F0F3CF5);
            WriteU16(head, 18, 1000);
        }
        // indexToLocFormat（+50），checkSumAdjustment（+8）= 0
        WriteU32(head, 8, 0);
        WriteS16(head, 50, (Int16)indexToLocFormat);
        return head;
    }

    private static Byte[] SubsetMaxp(Byte[] data, Int32 maxpOff, UInt16 numGlyphs)
    {
        var maxp = new Byte[32];
        if (maxpOff >= 0 && maxpOff + 32 <= data.Length)
            Array.Copy(data, maxpOff, maxp, 0, 32);
        else
        {
            WriteU32(maxp, 0, 0x00010000);
            WriteU16(maxp, 4, 0);
        }
        WriteU16(maxp, 4, numGlyphs);
        return maxp;
    }

    private static Byte[] SubsetHhea(Byte[] data, Int32 hheaOff, UInt16 numHMetrics)
    {
        var hhea = new Byte[36];
        if (hheaOff >= 0 && hheaOff + 36 <= data.Length)
            Array.Copy(data, hheaOff, hhea, 0, 36);
        else
        {
            WriteU32(hhea, 0, 0x00010000);
        }
        WriteU16(hhea, 34, numHMetrics);
        return hhea;
    }

    /// <summary>重建 hmtx：仅保留用到的字形（每个字形独立 advanceWidth）</summary>
    private static Byte[] SubsetHmtx(Byte[] data, Dictionary<String, (Int32 Offset, Int32 Length)> tables,
        Dictionary<Int32, Int32> remap, UInt16 numGlyphs)
    {
        var hmtxOff = GetTableOffset(tables, "hmtx");
        var hheaOff = GetTableOffset(tables, "hhea");
        var maxpOff = GetTableOffset(tables, "maxp");

        var hMetrics = 1;
        if (hheaOff >= 0 && hheaOff + 36 <= data.Length)
            hMetrics = ReadU16(data, hheaOff + 34);
        var oldGlyphs = 1;
        if (maxpOff >= 0 && maxpOff + 6 <= data.Length)
            oldGlyphs = ReadU16(data, maxpOff + 4);

        var hmtx = new Byte[numGlyphs * 4];
        foreach (var kv in remap)
        {
            var oldGid = kv.Key;
            var newGid = kv.Value;
            Int32 advance = 0, lsb = 0;
            if (oldGid < hMetrics && hmtxOff >= 0)
            {
                advance = ReadS16(data, hmtxOff + oldGid * 4);
                lsb = ReadS16(data, hmtxOff + oldGid * 4 + 2);
            }
            else if (hMetrics > 0 && hmtxOff >= 0)
            {
                // 最后一个 hMetric 的 advanceWidth + 后续 lsb
                advance = ReadS16(data, hmtxOff + (hMetrics - 1) * 4);
                var lsbOff = hmtxOff + hMetrics * 4 + (oldGid - hMetrics) * 2;
                if (oldGid < oldGlyphs && lsbOff + 2 <= data.Length)
                    lsb = ReadS16(data, lsbOff);
            }
            WriteS16(hmtx, newGid * 4, (Int16)advance);
            WriteS16(hmtx, newGid * 4 + 2, (Int16)lsb);
        }
        return hmtx;
    }

    /// <summary>重建 glyf/loca：拼接用到的字形数据，复合字形组件 GlyphID 重映射</summary>
    private static (Byte[] Loca, Byte[] Glyf) SubsetGlyfLoca(Byte[] data,
        Dictionary<String, (Int32 Offset, Int32 Length)> tables, Dictionary<Int32, Int32> remap)
    {
        var glyfOff = GetTableOffset(tables, "glyf");
        var glyfMax = glyfOff >= 0 ? GetTableLength(tables, "glyf") : 0;

        // 先按新 gid 收集原始 glyph 字节
        var ordered = remap.OrderBy(e => e.Value).ToList();
        var glyphBuffers = new List<Byte[]>(ordered.Count);
        foreach (var kv in ordered)
        {
            var (raw, _) = GetGlyphBytes(data, tables, kv.Key);
            glyphBuffers.Add(raw);
        }

        // 用 long loca（4 字节偏移）
        var loca = new Byte[(glyphBuffers.Count + 1) * 4];
        var glyf = new MemoryStream();
        for (var i = 0; i < glyphBuffers.Count; i++)
        {
            WriteU32(loca, i * 4, (UInt32)glyf.Length);

            var raw = glyphBuffers[i];
            if (raw.Length >= 2)
            {
                var numberOfContours = ReadS16(raw, 0);
                if (numberOfContours < 0)
                {
                    // 复合字形：重映射组件 GlyphID
                    var rewritten = RewriteComposite(raw, remap);
                    glyf.Write(rewritten, 0, rewritten.Length);
                }
                else
                {
                    glyf.Write(raw, 0, raw.Length);
                }
            }
        }
        WriteU32(loca, glyphBuffers.Count * 4, (UInt32)glyf.Length);
        return (loca, glyf.ToArray());

        // 局部函数：重写复合字形中的组件 GlyphID
        Byte[] RewriteComposite(Byte[] raw, Dictionary<Int32, Int32> map)
        {
            var sb = new MemoryStream();
            sb.Write(raw, 0, 10); // header（含 numberOfContours=负数）
            var pos = 10;
            while (pos + 4 <= raw.Length)
            {
                var flags = ReadU16(raw, pos);
                var compGid = ReadU16(raw, pos + 2);
                var newGid = map.TryGetValue(compGid, out var ng) ? (UInt16)ng : (UInt16)0;
                sb.WriteByte((Byte)(flags >> 8));
                sb.WriteByte((Byte)(flags & 0xFF));
                sb.WriteByte((Byte)(newGid >> 8));
                sb.WriteByte((Byte)(newGid & 0xFF));
                pos += 4;
                pos += (flags & 0x0001) != 0 ? 4 : 2;
                if ((flags & 0x0008) != 0) pos += 2;
                else if ((flags & 0x0040) != 0) pos += 4;
                else if ((flags & 0x0080) != 0) pos += 8;
                if ((flags & 0x0020) == 0) break;
            }
            // 指令区
            if (pos < raw.Length)
                sb.Write(raw, pos, raw.Length - pos);
            return sb.ToArray();
        }
    }

    /// <summary>重建 cmap（Format 4）：仅保留用到的 Unicode → 新 GlyphID</summary>
    private static (Byte[] Cmap, Dictionary<UInt16, UInt16> Map) SubsetCmap(
        Dictionary<UInt16, UInt16> oldCmap, Dictionary<Int32, Int32> remap)
    {
        var map = new Dictionary<UInt16, UInt16>();
        foreach (var kv in oldCmap)
        {
            if (remap.TryGetValue(kv.Value, out var newGid) && newGid != 0)
                map[kv.Key] = (UInt16)newGid;
        }

        var codes = map.Keys.OrderBy(c => c).ToList();
        var segCount = codes.Count + 1; // +1 终止段 0xFFFF
        var pow = 1;
        while (pow * 2 <= segCount) pow *= 2;

        using var ms = new MemoryStream();
        // cmap 头：version + numTables + encoding record（子表偏移 12）
        WriteU16(ms, 0);
        WriteU16(ms, 1);
        WriteU16(ms, 3);
        WriteU16(ms, 1);
        WriteU32(ms, 12);

        // Format 4 子表
        WriteU16(ms, 4);            // format
        WriteU16(ms, 0);            // length（占位，稍后回填）
        WriteU16(ms, 0);            // language
        WriteU16(ms, (UInt16)(segCount * 2));
        WriteU16(ms, (UInt16)(pow * 2));
        WriteU16(ms, (UInt16)Math.Log(pow, 2));
        WriteU16(ms, (UInt16)(segCount * 2 - pow * 2));

        // endCode[]
        foreach (var c in codes) WriteU16(ms, (UInt16)c);
        WriteU16(ms, 0xFFFF);
        WriteU16(ms, 0);            // reservedPad
        // startCode[]
        foreach (var c in codes) WriteU16(ms, (UInt16)c);
        WriteU16(ms, 0xFFFF);
        // idDelta[] = 0（用 idRangeOffset 定位）
        foreach (var _ in codes) WriteU16(ms, 0);
        WriteU16(ms, 1);
        // idRangeOffset[] = 2 * segCount（每段一个字形，glyph 紧随 idRangeOffset 数组）
        foreach (var _ in codes) WriteU16(ms, (UInt16)(segCount * 2));
        WriteU16(ms, 0);
        // glyphIdArray
        foreach (var c in codes) WriteU16(ms, map[c]);

        var bytes = ms.ToArray();
        // 回填子表 length（子表从偏移 12 开始，length 字段位于 14）
        var subLen = bytes.Length - 12;
        bytes[14] = (Byte)(subLen >> 8);
        bytes[15] = (Byte)(subLen & 0xFF);
        return (bytes, map);
    }

    #endregion

    #region sfnt 组装

    private static Byte[] BuildSfnt(Dictionary<String, Byte[]> tables)
    {
        var tags = tables.Keys.OrderBy(t => t).ToList();
        var numTables = tags.Count;
        var entrySelector = 0;
        var pow = 1;
        while (pow * 2 <= numTables) { pow *= 2; entrySelector++; }
        var searchRange = pow * 16;
        var rangeShift = numTables * 16 - searchRange;

        var headerLen = 12 + numTables * 16;
        // 表数据 4 字节对齐
        var totalLen = headerLen;
        foreach (var tag in tags)
            totalLen += (tables[tag].Length + 3) / 4 * 4;

        var data = new Byte[totalLen];
        // sfnt 头
        WriteU32(data, 0, 0x00010000);
        WriteU16(data, 4, (UInt16)numTables);
        WriteU16(data, 6, (UInt16)searchRange);
        WriteU16(data, 8, (UInt16)entrySelector);
        WriteU16(data, 10, (UInt16)rangeShift);

        var offset = headerLen;
        for (var i = 0; i < tags.Count; i++)
        {
            var tag = tags[i];
            var table = tables[tag];
            // 目录项
            var dirPos = 12 + i * 16;
            var tagBytes = Encoding.ASCII.GetBytes(tag);
            Array.Copy(tagBytes, 0, data, dirPos, 4);
            var checksum = CalcChecksum(table);
            WriteU32(data, dirPos + 4, checksum);
            WriteU32(data, dirPos + 8, (UInt32)offset);
            WriteU32(data, dirPos + 12, (UInt32)table.Length);
            // 表数据
            Array.Copy(table, 0, data, offset, table.Length);
            offset += (table.Length + 3) / 4 * 4;
        }

        // head.checkSumAdjustment（head 表偏移 +8）
        var headPos = FindTableOffset(data, "head");
        if (headPos >= 0)
        {
            var sum = CalcChecksum(data, 0, data.Length);
            var adjust = (0xB1B0AFBAu - sum) & 0xFFFFFFFFu;
            WriteU32(data, headPos + 8, (UInt32)adjust);
        }
        return data;
    }

    private static Int32 FindTableOffset(Byte[] data, String tag)
    {
        var numTables = ReadU16(data, 4);
        for (var i = 0; i < numTables; i++)
        {
            var pos = 12 + i * 16;
            if (pos + 16 > data.Length) break;
            var t = Encoding.ASCII.GetString(data, pos, 4);
            if (t == tag) return (Int32)ReadU32(data, pos + 8);
        }
        return -1;
    }

    private static UInt32 CalcChecksum(Byte[] table)
    {
        var padded = (table.Length + 3) / 4 * 4;
        var sum = 0u;
        for (var i = 0; i < padded; i += 4)
        {
            var b0 = i < table.Length ? table[i] : (Byte)0;
            var b1 = i + 1 < table.Length ? table[i + 1] : (Byte)0;
            var b2 = i + 2 < table.Length ? table[i + 2] : (Byte)0;
            var b3 = i + 3 < table.Length ? table[i + 3] : (Byte)0;
            sum += ((UInt32)b0 << 24) | ((UInt32)b1 << 16) | ((UInt32)b2 << 8) | b3;
        }
        return sum;
    }

    private static UInt32 CalcChecksum(Byte[] data, Int32 start, Int32 end)
    {
        var sum = 0u;
        for (var i = start; i < end; i += 4)
        {
            var b0 = i < end ? data[i] : (Byte)0;
            var b1 = i + 1 < end ? data[i + 1] : (Byte)0;
            var b2 = i + 2 < end ? data[i + 2] : (Byte)0;
            var b3 = i + 3 < end ? data[i + 3] : (Byte)0;
            sum += ((UInt32)b0 << 24) | ((UInt32)b1 << 16) | ((UInt32)b2 << 8) | b3;
        }
        return sum;
    }

    private static Byte[] CopyTable(Byte[] data, Int32 off, Int32 len)
    {
        if (off < 0 || len <= 0 || off + len > data.Length) return new Byte[0];
        var buf = new Byte[len];
        Array.Copy(data, off, buf, 0, len);
        return buf;
    }

    #endregion

    #region 二进制读取

    private static UInt16 ReadU16(Byte[] data, Int32 off) =>
        (UInt16)((data[off] << 8) | data[off + 1]);

    private static UInt16 ReadU16(Byte[] data, Int32 baseOff, Int32 off) =>
        ReadU16(data, baseOff + off);

    private static Int16 ReadS16(Byte[] data, Int32 off) =>
        (Int16)ReadU16(data, off);

    private static Int16 ReadS16(Byte[] data, Int32 baseOff, Int32 off) =>
        ReadS16(data, baseOff + off);

    private static UInt32 ReadU32(Byte[] data, Int32 off) =>
        ((UInt32)data[off] << 24) | ((UInt32)data[off + 1] << 16) |
        ((UInt32)data[off + 2] << 8) | data[off + 3];

    private static void WriteU16(Byte[] data, Int32 off, UInt16 val)
    {
        data[off] = (Byte)(val >> 8);
        data[off + 1] = (Byte)(val & 0xFF);
    }

    private static void WriteS16(Byte[] data, Int32 off, Int16 val) => WriteU16(data, off, (UInt16)val);

    private static void WriteU32(Byte[] data, Int32 off, UInt32 val)
    {
        data[off] = (Byte)(val >> 24);
        data[off + 1] = (Byte)((val >> 16) & 0xFF);
        data[off + 2] = (Byte)((val >> 8) & 0xFF);
        data[off + 3] = (Byte)(val & 0xFF);
    }

    private static void WriteU16(MemoryStream ms, UInt16 val)
    {
        ms.WriteByte((Byte)(val >> 8));
        ms.WriteByte((Byte)(val & 0xFF));
    }

    private static void WriteU32(MemoryStream ms, UInt32 val)
    {
        ms.WriteByte((Byte)(val >> 24));
        ms.WriteByte((Byte)((val >> 16) & 0xFF));
        ms.WriteByte((Byte)((val >> 8) & 0xFF));
        ms.WriteByte((Byte)(val & 0xFF));
    }

    #endregion
}
