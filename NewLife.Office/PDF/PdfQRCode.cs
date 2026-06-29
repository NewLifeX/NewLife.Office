using System.Text;
using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>纯 C# QR 码生成器（零外部依赖），支持版本 1-6（21x21 至 41x41），字节模式 + ECC M 级</summary>
public static class PdfQRCode
{
    #region 公共方法
    /// <summary>生成 QR 码 PNG 字节数组</summary>
    /// <param name="text">要编码的文本（URL 或简短字符串）</param>
    /// <param name="moduleSize">每模块像素数（默认 4，生成合理尺寸）</param>
    /// <returns>PNG 格式字节</returns>
    public static Byte[] Generate(String text, Int32 moduleSize = 4)
    {
        if (text.IsNullOrEmpty()) throw new ArgumentNullException(nameof(text));

        var data = Encoding.UTF8.GetBytes(text);
        var version = ChooseVersion(data.Length);
        var modules = (version - 1) * 4 + 21; // 每边模块数
        var matrix = new Boolean[modules, modules];

        // 1. 放置功能图案
        PlaceFinders(matrix);
        PlaceTiming(matrix);
        PlaceAlignment(matrix, version);
        ReserveFormat(matrix);

        // 2. 编码数据
        var codewords = EncodeData(data, version);

        // 3. 放置数据模块
        PlaceData(matrix, codewords);

        // 4. 应用掩码（选择最优）
        ApplyBestMask(matrix);

        // 5. 写入格式信息
        WriteFormatInfo(matrix);

        // 6. 渲染为 PNG
        return RenderPng(matrix, moduleSize);
    }
    #endregion

    #region 版本选择
    private static Int32 ChooseVersion(Int32 byteCount)
    {
        // ECC Level M (Medium): 各版本最大字节容量
        var caps = new[] { 0, 14, 26, 42, 62, 84, 106 }; // V1-V6
        for (var v = 1; v < caps.Length; v++)
        {
            if (byteCount <= caps[v]) return v;
        }
        if (byteCount <= 106) return 6;
        throw new ArgumentException($"文本过长（{byteCount} 字节），最大支持 106 字节。请缩短文本。");
    }
    #endregion

    #region 功能图案放置
    private static void PlaceFinders(Boolean[,] m)
    {
        var size = m.GetLength(0);
        // 三个定位图案：左上、右上、左下
        PlaceFinder(m, 0, 0);
        PlaceFinder(m, 0, size - 7);
        PlaceFinder(m, size - 7, 0);
    }

    private static void PlaceFinder(Boolean[,] m, Int32 row, Int32 col)
    {
        for (var r = 0; r < 7; r++)
        {
            for (var c = 0; c < 7; c++)
            {
                // 外框 7x7 黑，内部 5x5 白，中心 3x3 黑
                var outer = r == 0 || r == 6 || c == 0 || c == 6;
                var inner = r >= 2 && r <= 4 && c >= 2 && c <= 4;
                m[row + r, col + c] = outer || inner;
            }
        }
    }

    private static void PlaceTiming(Boolean[,] m)
    {
        var size = m.GetLength(0);
        for (var i = 8; i < size - 8; i++)
        {
            m[6, i] = i % 2 == 0; // 水平
            m[i, 6] = i % 2 == 0; // 垂直
        }
    }

    private static void PlaceAlignment(Boolean[,] m, Int32 version)
    {
        if (version < 2) return;
        // 版本 2-6 对齐图案位置表
        var positions = version switch
        {
            2 => new[] { 6, 18 },
            3 => new[] { 6, 22 },
            4 => new[] { 6, 26 },
            5 => new[] { 6, 30 },
            6 => new[] { 6, 34 },
            _ => new[] { 6, 18 }
        };

        foreach (var r in positions)
        {
            foreach (var c in positions)
            {
                // 跳过与定位图案重叠的位置
                if ((r == 6 && c == 6) || (r == 6 && c == positions[^1]) || (r == positions[^1] && c == 6))
                    continue;
                PlaceAlignmentBlock(m, r - 2, c - 2);
            }
        }
    }

    private static void PlaceAlignmentBlock(Boolean[,] m, Int32 row, Int32 col)
    {
        for (var r = 0; r < 5; r++)
        {
            for (var c = 0; c < 5; c++)
            {
                var outer = r == 0 || r == 4 || c == 0 || c == 4;
                var center = r == 2 && c == 2;
                m[row + r, col + c] = outer || center;
            }
        }
    }

    private static void ReserveFormat(Boolean[,] m)
    {
        var size = m.GetLength(0);
        // 格式信息保留区域（标记为 true 占位，后续写入）
        for (var i = 0; i < 9; i++)
        {
            if (i != 6) { m[8, i] = true; m[i, 8] = true; } // 左上
        }
        m[8, 8] = true; // 暗模块参考
        // 右上
        for (var i = size - 8; i < size; i++)
            m[8, i] = true;
        // 左下
        for (var i = size - 7; i < size; i++)
            m[i, 8] = true;
    }
    #endregion

    #region 数据编码
    private static Byte[] EncodeData(Byte[] data, Int32 version)
    {
        // ECC M: 每版本 (总码字, 数据码字, EC码字数, 块数, 块1数据, 块1EC)
        var eccTable = new (Int32 Total, Int32 Data, Int32 EC, Int32 Blocks, Int32 B1Data, Int32 B1EC)[]
        {
            (0,0,0,0,0,0), // v0 (unused)
            (26, 16, 10, 1, 16, 10),  // v1
            (44, 28, 16, 1, 28, 16),  // v2
            (70, 44, 26, 1, 44, 26),  // v3
            (100, 64, 36, 2, 32, 18), // v4: 2 blocks, each 32 data + 18 EC
            (134, 86, 48, 2, 43, 24), // v5
            (172, 108, 64, 2, 54, 32),// v6
        };

        var ecc = eccTable[version];
        var dataCodewords = new Byte[ecc.Total];

        // 模式指示符: 0100 (字节模式)
        var pos = 0;
        WriteBits(dataCodewords, ref pos, 0x4, 4);  // 模式

        // 字符计数（8位）
        WriteBits(dataCodewords, ref pos, data.Length, 8);

        // 数据字节
        foreach (var b in data)
            WriteBits(dataCodewords, ref pos, b, 8);

        // 终止符（0000）
        WriteBits(dataCodewords, ref pos, 0, 4);

        // 填充到 8 位边界
        pos = (pos + 7) / 8 * 8;

        // 填充字节 0xEC 0x11 交替直到填满数据码字
        var pad = false;
        var dataBytes = ecc.Data;
        while (pos / 8 < dataBytes)
        {
            WriteBits(dataCodewords, ref pos, pad ? 0x11 : 0xEC, 8);
            pad = !pad;
        }

        // 分割成块并计算 ECC
        if (ecc.Blocks == 1)
        {
            // 单块：直接计算 ECC
            var dataBlock = new Byte[ecc.B1Data];
            Array.Copy(dataCodewords, 0, dataBlock, 0, ecc.B1Data);
            var ecBlock = ComputeECC(dataBlock, ecc.B1EC);
            var result = new Byte[ecc.Total];
            Array.Copy(dataBlock, 0, result, 0, dataBlock.Length);
            Array.Copy(ecBlock, 0, result, dataBlock.Length, ecBlock.Length);
            return result;
        }
        else
        {
            // 多块：分别计算后再交织
            var blocks = ecc.Blocks;
            var b1Data = ecc.B1Data;
            var b1EC = ecc.B1EC;

            var dataBlocks = new Byte[blocks][];
            var ecBlocks = new Byte[blocks][];
            var offset = 0;
            for (var b = 0; b < blocks; b++)
            {
                var blockData = new Byte[b1Data];
                Array.Copy(dataCodewords, offset, blockData, 0, b1Data);
                offset += b1Data;
                dataBlocks[b] = blockData;
                ecBlocks[b] = ComputeECC(blockData, b1EC);
            }

            // 交织
            var result = new Byte[ecc.Total];
            var ri = 0;
            for (var i = 0; i < b1Data; i++)
                for (var b = 0; b < blocks; b++)
                    if (i < dataBlocks[b].Length) result[ri++] = dataBlocks[b][i];
            for (var i = 0; i < b1EC; i++)
                for (var b = 0; b < blocks; b++)
                    result[ri++] = ecBlocks[b][i];

            return result;
        }
    }

    private static void WriteBits(Byte[] buf, ref Int32 bitPos, Int32 value, Int32 numBits)
    {
        for (var i = numBits - 1; i >= 0; i--)
        {
            var byteIdx = bitPos / 8;
            var bitIdx = 7 - (bitPos % 8);
            if (((value >> i) & 1) != 0)
                buf[byteIdx] |= (Byte)(1 << bitIdx);
            bitPos++;
        }
    }

    /// <summary>Reed-Solomon ECC 计算（GF(256)）</summary>
    private static Byte[] ComputeECC(Byte[] data, Int32 ecCount)
    {
        // 生成多项式系数（硬编码，支持最多 36 个 ECC 码字）
        var generator = new Byte[] { 1 };
        for (var i = 0; i < ecCount; i++)
        {
            // 乘以 (x + α^i)
            var next = new Byte[generator.Length + 1];
            for (var j = 0; j < generator.Length; j++)
            {
                next[j] = GfMul(generator[j], GfExp(i));
                if (j > 0) next[j] = (Byte)(next[j] ^ generator[j - 1]);
            }
            next[generator.Length] = generator[^1];
            generator = next;
        }

        // 多项式除法
        var result = new Byte[ecCount];
        var msg = new Byte[data.Length + ecCount];
        Array.Copy(data, 0, msg, 0, data.Length);

        for (var i = 0; i < data.Length; i++)
        {
            var factor = msg[i];
            if (factor == 0) continue;
            for (var j = 0; j < generator.Length; j++)
                msg[i + j] = (Byte)(msg[i + j] ^ GfMul(generator[j], factor));
        }

        Array.Copy(msg, data.Length, result, 0, ecCount);
        return result;
    }

    // GF(256) 运算：使用原多项式 x^8 + x^4 + x^3 + x^2 + 1 (0x11D)
    private static Byte GfMul(Byte a, Byte b)
    {
        var result = 0;
        for (var i = 0; i < 8; i++)
        {
            if ((b & 1) != 0) result ^= a;
            var high = a & 0x80;
            a <<= 1;
            if (high != 0) a ^= 0x1D;
            b >>= 1;
        }
        return (Byte)result;
    }

    private static Byte GfExp(Int32 power)
    {
        // α^0 = 1, α^1 = 2, ... 预计算
        var result = 1;
        for (var i = 0; i < power; i++)
        {
            result <<= 1;
            if ((result & 0x100) != 0) result ^= 0x11D;
        }
        return (Byte)result;
    }
    #endregion

    #region 数据模块放置
    private static void PlaceData(Boolean[,] m, Byte[] codewords)
    {
        var size = m.GetLength(0);
        var totalBits = codewords.Length * 8;

        // 从右下角开始，之字形向上放置
        var col = size - 1;
        var row = size - 1;
        var up = true;
        var bitIdx = 0;

        while (col > 0)
        {
            if (col == 6) col--; // 跳过垂直时序图案列

            for (var r = 0; r < size; r++)
            {
                var rr = up ? size - 1 - r : r;
                for (var dc = 0; dc < 2; dc++)
                {
                    var cc = col - dc;
                    if (cc < 0) continue;

                    // 检查是否为功能模块保留位
                    if (!IsDataModule(m, rr, cc)) continue;

                    if (bitIdx < totalBits)
                    {
                        var byteIdx = bitIdx / 8;
                        var bitInByte = 7 - (bitIdx % 8);
                        m[rr, cc] = ((codewords[byteIdx] >> bitInByte) & 1) != 0;
                        bitIdx++;
                    }
                    else
                    {
                        m[rr, cc] = false;
                    }
                }
            }

            col -= 2;
            up = !up;
        }
    }

    private static Boolean IsDataModule(Boolean[,] m, Int32 row, Int32 col)
    {
        // 定位图案区域、时序线、格式信息区域不可覆盖
        // 简化判断：检查是否已被写入 true（功能图案）
        // 功能图案（finder/timing/alignment）在放置时已设为 true
        // 所以这里只需要检查是否为 false（空位）
        return !m[row, col];
    }
    #endregion

    #region 掩码与格式
    private static void ApplyBestMask(Boolean[,] m)
    {
        var size = m.GetLength(0);
        var bestMatrix = (Boolean[,])m.Clone();
        var bestScore = Int32.MaxValue;
        var bestMask = 0;

        for (var mask = 0; mask < 8; mask++)
        {
            var test = (Boolean[,])m.Clone();
            ApplyMask(test, mask);
            var score = EvaluateMask(test);
            if (score < bestScore)
            {
                bestScore = score;
                bestMask = mask;
                bestMatrix = test;
            }
        }

        Array.Copy(bestMatrix, m, m.Length);
        // 存储掩码编号用于格式信息（通过静态字段）
        _bestMask = bestMask;
    }

    private static Int32 _bestMask;

    private static void ApplyMask(Boolean[,] m, Int32 mask)
    {
        var size = m.GetLength(0);
        for (var r = 0; r < size; r++)
        {
            for (var c = 0; c < size; c++)
            {
                if (!IsDataModule(m, r, c)) continue;
                var invert = mask switch
                {
                    0 => (r + c) % 2 == 0,
                    1 => r % 2 == 0,
                    2 => c % 3 == 0,
                    3 => (r + c) % 3 == 0,
                    4 => ((r / 2) + (c / 3)) % 2 == 0,
                    5 => (r * c) % 2 + (r * c) % 3 == 0,
                    6 => ((r * c) % 2 + (r * c) % 3) % 2 == 0,
                    _ => ((r + c) % 2 + (r * c) % 3) % 2 == 0,
                };
                if (invert) m[r, c] = !m[r, c];
            }
        }
    }

    private static Int32 EvaluateMask(Boolean[,] m)
    {
        // 简化惩罚评分：连续相同颜色行/列 + 1:1:3:1:1 比例模式
        var size = m.GetLength(0);
        var score = 0;

        // 连续同色（水平）
        for (var r = 0; r < size; r++)
        {
            var run = 1;
            for (var c = 1; c < size; c++)
            {
                if (m[r, c] == m[r, c - 1]) { run++; }
                else
                {
                    if (run >= 5) score += run - 2;
                    run = 1;
                }
            }
            if (run >= 5) score += run - 2;
        }

        // 连续同色（垂直）
        for (var c = 0; c < size; c++)
        {
            var run = 1;
            for (var r = 1; r < size; r++)
            {
                if (m[r, c] == m[r - 1, c]) { run++; }
                else
                {
                    if (run >= 5) score += run - 2;
                    run = 1;
                }
            }
            if (run >= 5) score += run - 2;
        }

        // 2x2 块
        for (var r = 0; r < size - 1; r++)
            for (var c = 0; c < size - 1; c++)
                if (m[r, c] == m[r, c + 1] && m[r, c] == m[r + 1, c] && m[r, c] == m[r + 1, c + 1])
                    score += 3;

        return score;
    }

    private static void WriteFormatInfo(Boolean[,] m)
    {
        var size = m.GetLength(0);
        // ECC M(00) + mask pattern
        var formatBits = (0 << 3) | (_bestMask & 0x7); // ECC M = 00
        var format = EncodeFormat(formatBits);

        // 15 位格式信息写入固定位置
        // 左上角分离（绕开定位图案）
        var pos = 14;
        for (var i = 0; i <= 8; i++) if (i != 6) m[8, i] = ((format >> pos--) & 1) != 0;
        for (var i = 7; i >= 0; i--) if (i != 6) m[i, 8] = ((format >> pos--) & 1) != 0;

        // 右上角（在定位图案下方）
        pos = 14;
        for (var i = size - 1; i >= size - 8; i--) { m[8, i] = ((format >> pos--) & 1) != 0; if (pos < 0) break; }

        // 左下角（在定位图案右侧）
        pos = 14;
        for (var i = size - 8; i < size; i++) { m[i, 8] = ((format >> pos--) & 1) != 0; if (pos < 0) break; }
    }

    private static Int32 EncodeFormat(Int32 data)
    {
        // BCH(15,5) with generator polynomial x^10 + x^8 + x^5 + x^4 + x^2 + x + 1 (0x537)
        var code = data << 10;
        const Int32 gen = 0x537;
        for (var i = 4; i >= 0; i--)
        {
            if ((code & (1 << (i + 10))) != 0)
                code ^= gen << i;
        }
        var result = ((data << 10) | (code & 0x3FF)) ^ 0x5412; // XOR with mask 101010000010010
        return result;
    }
    #endregion

    #region PNG 渲染
    private static Byte[] RenderPng(Boolean[,] matrix, Int32 moduleSize)
    {
        var modules = matrix.GetLength(0);
        var quietZone = moduleSize * 4; // 每边空白区
        var imgSize = modules * moduleSize + quietZone * 2;

        // 构建原始 RGBA 像素
        var pixels = new Byte[imgSize * imgSize * 4];
        for (var r = 0; r < imgSize; r++)
        {
            for (var c = 0; c < imgSize; c++)
            {
                var idx = (r * imgSize + c) * 4;
                // 检查是否在 QR 码区域内
                var qr = r >= quietZone && r < imgSize - quietZone &&
                         c >= quietZone && c < imgSize - quietZone;
                var mr = (r - quietZone) / moduleSize;
                var mc = (c - quietZone) / moduleSize;
                var isBlack = qr && mr >= 0 && mr < modules && mc >= 0 && mc < modules && matrix[mr, mc];

                pixels[idx] = (Byte)(isBlack ? 0 : 255);     // R
                pixels[idx + 1] = (Byte)(isBlack ? 0 : 255); // G
                pixels[idx + 2] = (Byte)(isBlack ? 0 : 255); // B
                pixels[idx + 3] = 255;                        // A
            }
        }

        // 构建最小 PNG
        return EncodePng(pixels, imgSize, imgSize);
    }

    private static Byte[] EncodePng(Byte[] rgba, Int32 width, Int32 height)
    {
        using var ms = new MemoryStream();
        // PNG 签名
        ms.Write(new Byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, 0, 8);

        // IHDR
        WritePngChunk(ms, "IHDR", w =>
        {
            WriteUInt32BE(w, (UInt32)width);
            WriteUInt32BE(w, (UInt32)height);
            w.WriteByte(8);  // bit depth
            w.WriteByte(6);  // color type: RGBA
            w.WriteByte(0);  // compression
            w.WriteByte(0);  // filter
            w.WriteByte(0);  // interlace
        });

        // IDAT
        WritePngChunk(ms, "IDAT", w =>
        {
            // 构建原始扫描线（每行前加 filter=0）
            var raw = new Byte[height * (1 + width * 4)];
            for (var r = 0; r < height; r++)
            {
                raw[r * (1 + width * 4)] = 0; // filter: None
                rgba.AsSpan(r * width * 4, width * 4).CopyTo(raw.AsSpan(r * (1 + width * 4) + 1, width * 4));
            }

            // zlib 压缩（Deflate）
            using var compressed = new MemoryStream();
            using (var deflate = new System.IO.Compression.DeflateStream(compressed, System.IO.Compression.CompressionLevel.Optimal, true))
            {
                deflate.Write(raw, 0, raw.Length);
            }
            var compressedBytes = compressed.ToArray();

            // zlib 包装：2字节头 + 压缩数据 + 4字节 Adler32
            w.WriteByte(0x78); // CMF
            w.WriteByte(0xDA); // FLG: max compression, 32K window
            w.Write(compressedBytes, 0, compressedBytes.Length);
            var adler = Adler32(raw);
            WriteUInt32BE(w, (UInt32)((adler >> 24) & 0xFF | (adler >> 8) & 0xFF00 | (adler << 8) & 0xFF0000 | (adler << 24) & 0xFF000000));
        });

        // IEND
        WritePngChunk(ms, "IEND", _ => { });

        return ms.ToArray();
    }

    private static void WritePngChunk(Stream ms, String type, Action<MemoryStream> writeData)
    {
        var dataMs = new MemoryStream();
        writeData(dataMs);
        var data = dataMs.ToArray();

        // 长度（4 字节，大端）
        WriteUInt32BE(ms, (UInt32)data.Length);
        // 类型（4 字节 ASCII）
        var typeBytes = Encoding.ASCII.GetBytes(type);
        ms.Write(typeBytes, 0, 4);
        // 数据
        ms.Write(data, 0, data.Length);
        // CRC32（类型 + 数据）
        Span<Byte> crcInput = stackalloc Byte[4 + data.Length];
        typeBytes.AsSpan().CopyTo(crcInput);
        data.AsSpan().CopyTo(crcInput.Slice(4));
        var crc = Crc32(crcInput);
        WriteUInt32BE(ms, crc);
    }

    private static void WriteUInt32BE(Stream ms, UInt32 value)
    {
        var buf = new Byte[4];
        var writer = new SpanWriter(buf) { IsLittleEndian = false };
        writer.Write(value);
        ms.Write(buf, 0, 4);
    }

    private static UInt32 Crc32(ReadOnlySpan<Byte> data)
    {
        var crc = 0xFFFFFFFFu;
        for (var i = 0; i < data.Length; i++)
        {
            crc ^= data[i];
            for (var j = 0; j < 8; j++)
                crc = (crc >> 1) ^ ((crc & 1) != 0 ? 0xEDB88320u : 0);
        }
        return crc ^ 0xFFFFFFFFu;
    }

    private static Int32 Adler32(Byte[] data)
    {
        var a = 1;
        var b = 0;
        for (var i = 0; i < data.Length; i++)
        {
            a = (a + data[i]) % 65521;
            b = (b + a) % 65521;
        }
        return (b << 16) | a;
    }
    #endregion
}
