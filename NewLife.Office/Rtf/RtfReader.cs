using System;
using System.Collections.Generic;
using System.Text;

namespace NewLife.Office.Rtf;

/// <summary>RTF 解析器</summary>
/// <remarks>
/// 将 RTF 文本解析为 RtfDocument 对象模型。
/// 支持：控制字/符号、分组嵌套、字体表、颜色表、\u Unicode、\'XX ANSI、段落格式、表格结构。
/// </remarks>
internal sealed class RtfReader
{
    #region 字段
    private String _rtf = "";
    private Int32 _pos;

    // 字体表 index → name
    private readonly List<String> _fontTable = [];
    // 颜色表 index → RGB（0 = auto = -1）
    private readonly List<Int32> _colorTable = [];
    // 文档属性
    private String _title = "";
    private String _author = "";
    private String _subject = "";
    #endregion

    #region 入口
    /// <summary>解析 RTF 文本，返回文档对象</summary>
    /// <param name="rtf">RTF 源文本</param>
    /// <returns>文档对象</returns>
    public RtfDocument Read(String rtf)
    {
        _rtf = rtf;
        _pos = 0;

        var doc = new RtfDocument();
        if (!rtf.TrimStart().StartsWith("{\\rtf")) return doc;

        // Skip opening {
        SkipWhitespace();
        if (_pos < _rtf.Length && _rtf[_pos] == '{') _pos++;

        // Parse top-level group content
        ParseGroup(doc);

        doc.Title = _title;
        doc.Author = _author;
        doc.Subject = _subject;
        return doc;
    }
    #endregion

    #region 解析状态
    /// <summary>当前格式状态（可入栈/出栈）</summary>
    private sealed class FormatState
    {
        public Int32 FontIndex { get; set; }
        public Int32 FontSize { get; set; } = 24;       // 半磅，默认 12pt
        public Boolean Bold { get; set; }
        public Boolean Italic { get; set; }
        public Boolean Underline { get; set; }
        public Boolean Strikethrough { get; set; }
        public Int32 ForeColorIndex { get; set; }       // 颜色表索引（0=auto）
        public Int32 BackColorIndex { get; set; }
        public RtfAlignment Alignment { get; set; }
        public Int32 LeftIndent { get; set; }
        public Int32 RightIndent { get; set; }
        public Int32 FirstLineIndent { get; set; }
        public Int32 SpaceBefore { get; set; }
        public Int32 SpaceAfter { get; set; }
        public Int32 LineSpacing { get; set; }
        public Boolean InTable { get; set; }
        public Boolean Hidden { get; set; }

        /// <summary>克隆当前状态</summary>
        public FormatState Clone() => new()
        {
            FontIndex = FontIndex,
            FontSize = FontSize,
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Strikethrough = Strikethrough,
            ForeColorIndex = ForeColorIndex,
            BackColorIndex = BackColorIndex,
            Alignment = Alignment,
            LeftIndent = LeftIndent,
            RightIndent = RightIndent,
            FirstLineIndent = FirstLineIndent,
            SpaceBefore = SpaceBefore,
            SpaceAfter = SpaceAfter,
            LineSpacing = LineSpacing,
            InTable = InTable,
            Hidden = Hidden,
        };
    }
    #endregion

    #region 解析主体
    private void ParseGroup(RtfDocument doc)
    {
        var stateStack = new Stack<FormatState>();
        var currentState = new FormatState();

        // 当前构建中的文档块
        var currentPara = new RtfParagraph();
        var currentRun = new StringBuilder();
        var currentTable = (RtfTable?)null;
        var currentRow = (RtfTableRow?)null;
        var currentCell = (RtfTableCell?)null;
        var cellBoundaries = new List<Int32>();

        void FlushRun()
        {
            if (currentRun.Length == 0) return;
            var text = currentRun.ToString();
            currentRun.Clear();
            if (currentState.Hidden) return;
            var run = BuildRun(text, currentState, false);
            if (currentCell != null)
            {
                if (currentCell.Paragraphs.Count == 0) currentCell.Paragraphs.Add(currentPara);
                currentPara.Runs.Add(run);
            }
            else
            {
                currentPara.Runs.Add(run);
            }
        }

        void CommitParagraph()
        {
            FlushRun();
            ApplyParaFormat(currentPara, currentState);
            if (currentCell != null)
            {
                if (!currentCell.Paragraphs.Contains(currentPara))
                    currentCell.Paragraphs.Add(currentPara);
                currentPara = new RtfParagraph { InTable = true };
            }
            else if (currentPara.Runs.Count > 0 || currentPara.Alignment != RtfAlignment.Left)
            {
                doc.Blocks.Add(currentPara);
                currentPara = new RtfParagraph();
            }
            else
            {
                // empty paragraph — keep as line break
                doc.Blocks.Add(currentPara);
                currentPara = new RtfParagraph();
            }
        }

        void CommitCell()
        {
            FlushRun();
            if (currentCell != null)
            {
                if (currentPara.Runs.Count > 0 && (currentCell.Paragraphs.Count == 0 || !currentCell.Paragraphs.Contains(currentPara)))
                    currentCell.Paragraphs.Add(currentPara);
                currentRow?.Cells.Add(currentCell);
                currentCell = null;
                currentPara = new RtfParagraph { InTable = true };
            }
        }

        void CommitRow()
        {
            CommitCell();
            if (currentRow != null && currentRow.Cells.Count > 0)
            {
                if (currentTable == null) currentTable = new RtfTable();
                currentTable.Rows.Add(currentRow);
                currentRow = null;
            }
            cellBoundaries.Clear();
        }

        void CommitTable()
        {
            CommitRow();
            if (currentTable != null && currentTable.Rows.Count > 0)
            {
                doc.Blocks.Add(currentTable);
                currentTable = null;
            }
            currentState.InTable = false;
        }

        var depth = 1; // we already consumed the outer {
        var skipGroup = false;
        var skipDepth = 0;

        while (_pos < _rtf.Length && depth > 0)
        {
            var ch = _rtf[_pos];

            if (ch == '{')
            {
                _pos++;
                if (skipGroup)
                {
                    skipDepth++;
                    continue;
                }
                FlushRun(); // flush pending text before entering new formatting context
                stateStack.Push(currentState.Clone());
                depth++;

                // Check for destination groups we want to skip or handle
                var savedPos = _pos;
                SkipWhitespace();
                if (_pos < _rtf.Length && _rtf[_pos] == '\\')
                {
                    var kw = PeekKeyword();
                    if (kw == "\\*")
                    {
                        // unknown optional destination — skip entire group
                        _pos += 2;
                        skipGroup = true;
                        skipDepth = 1;
                        continue;
                    }
                    else if (kw == "\\fonttbl")
                    {
                        _pos += 8;
                        ParseFontTable(); // consumes up to and including closing }
                        if (stateStack.Count > 0) currentState = stateStack.Pop();
                        depth--;
                        continue;
                    }
                    else if (kw == "\\colortbl")
                    {
                        _pos += 9;
                        ParseColorTable(); // consumes up to and including closing }
                        if (stateStack.Count > 0) currentState = stateStack.Pop();
                        depth--;
                        continue;
                    }
                    else if (kw == "\\info")
                    {
                        _pos += 5;
                        ParseInfo(); // consumes up to and including closing }
                        if (stateStack.Count > 0) currentState = stateStack.Pop();
                        depth--;
                        continue;
                    }
                    else if (kw == "\\pict")
                    {
                        ReadControlWord(); // consume "\pict"
                        var image = ParsePict(); // consumes up to and including closing }
                        if (image != null) doc.Images.Add(image);
                        if (stateStack.Count > 0) currentState = stateStack.Pop();
                        depth--;
                        continue;
                    }
                    else if (kw == "\\stylesheet" || kw == "\\listtable" || kw == "\\listoverridetable"
                          || kw == "\\rsidtbl" || kw == "\\mmathPr" || kw == "\\themedata"
                          || kw == "\\colorschememapping" || kw == "\\latentstyles" || kw == "\\generator")
                    {
                        // skip these destination groups entirely (consumes including })
                        SkipGroup();
                        if (stateStack.Count > 0) currentState = stateStack.Pop();
                        depth--;
                        continue;
                    }
                    _pos = savedPos;
                }
                else
                {
                    _pos = savedPos;
                }
                continue;
            }

            if (ch == '}')
            {
                _pos++;
                if (skipGroup)
                {
                    skipDepth--;
                    if (skipDepth == 0)
                    {
                        skipGroup = false;
                        // The { that opened this optional group pushed state and incremented depth
                        if (stateStack.Count > 0) currentState = stateStack.Pop();
                        depth--;
                    }
                    continue;
                }
                FlushRun(); // flush pending text before restoring previous formatting context
                if (stateStack.Count > 0) currentState = stateStack.Pop();
                depth--;
                continue;
            }

            if (skipGroup) { _pos++; continue; }

            if (ch == '\\')
            {
                var (word, param, hasParam) = ReadControlWord();
                ApplyControlWord(word, param, hasParam, currentState,
                    ref currentRun, ref currentPara, ref currentRow, ref currentCell, ref currentTable,
                    cellBoundaries, CommitParagraph, CommitCell, CommitRow, CommitTable, FlushRun, doc);
                continue;
            }

            // Plain text
            if (ch == '\r' || ch == '\n') { _pos++; continue; }
            currentRun.Append(ch);
            _pos++;
        }

        // Flush remaining content
        FlushRun();
        if (currentPara.Runs.Count > 0)
            doc.Blocks.Add(currentPara);
        if (currentTable != null && currentTable.Rows.Count > 0)
            doc.Blocks.Add(currentTable);
    }
    #endregion

    #region 控制字处理
    private void ApplyControlWord(String word, Int32 param, Boolean hasParam,
        FormatState state,
        ref StringBuilder currentRun,
        ref RtfParagraph currentPara,
        ref RtfTableRow? currentRow,
        ref RtfTableCell? currentCell,
        ref RtfTable? currentTable,
        List<Int32> cellBoundaries,
        Action commitPara,
        Action commitCell,
        Action commitRow,
        Action commitTable,
        Action flushRun,
        RtfDocument doc)
    {
        switch (word)
        {
            // ── 文档元数据 ──
            case "\\rtf": break;
            case "\\ansi": case "\\mac": case "\\pc": case "\\pca": break;
            case "\\deff": break;
            case "\\deflang": break;

            // ── 字符格式 ──
            case "\\b":    state.Bold = !hasParam || param != 0; break;
            case "\\i":    state.Italic = !hasParam || param != 0; break;
            case "\\ul":   state.Underline = !hasParam || param != 0; break;
            case "\\ulnone": state.Underline = false; break;
            case "\\strike": case "\\striked": state.Strikethrough = !hasParam || param != 0; break;
            case "\\plain":
                state.Bold = false; state.Italic = false; state.Underline = false;
                state.Strikethrough = false; state.ForeColorIndex = 0;
                state.BackColorIndex = 0; state.FontSize = 24;
                break;
            case "\\f":    if (hasParam) state.FontIndex = param; break;
            case "\\fs":   if (hasParam) state.FontSize = param; break;
            case "\\cf":   if (hasParam) state.ForeColorIndex = param; break;
            case "\\highlight": if (hasParam) state.BackColorIndex = param; break;
            case "\\v":    state.Hidden = !hasParam || param != 0; break;

            // ── 段落格式 ──
            case "\\pard":
                state.Alignment = RtfAlignment.Left;
                state.LeftIndent = 0; state.RightIndent = 0; state.FirstLineIndent = 0;
                state.SpaceBefore = 0; state.SpaceAfter = 0; state.LineSpacing = 0;
                break;
            case "\\ql": state.Alignment = RtfAlignment.Left; break;
            case "\\qc": state.Alignment = RtfAlignment.Center; break;
            case "\\qr": state.Alignment = RtfAlignment.Right; break;
            case "\\qj": state.Alignment = RtfAlignment.Justify; break;
            case "\\li": if (hasParam) state.LeftIndent = param; break;
            case "\\ri": if (hasParam) state.RightIndent = param; break;
            case "\\fi": if (hasParam) state.FirstLineIndent = param; break;
            case "\\sb": if (hasParam) state.SpaceBefore = param; break;
            case "\\sa": if (hasParam) state.SpaceAfter = param; break;
            case "\\sl": if (hasParam) state.LineSpacing = param; break;

            // ── 段落/换行 ──
            case "\\par":
                commitPara();
                break;
            case "\\line":
                flushRun();
                currentPara.Runs.Add(new RtfRun { IsLineBreak = true });
                break;
            case "\\tab":
                currentRun.Append('\t');
                break;

            // ── 表格 ──
            case "\\trowd":
                state.InTable = true;
                if (currentRow == null) currentRow = new RtfTableRow();
                cellBoundaries.Clear();
                break;
            case "\\cellx":
                if (hasParam) cellBoundaries.Add(param);
                if (currentCell == null) currentCell = new RtfTableCell
                    { RightBoundary = hasParam ? param : 0 };
                break;
            case "\\cell":
                commitCell();
                if (currentRow != null && cellBoundaries.Count > currentRow.Cells.Count)
                {
                    var boundary = cellBoundaries[currentRow.Cells.Count];
                    currentCell = new RtfTableCell { RightBoundary = boundary };
                }
                else
                {
                    currentCell = new RtfTableCell();
                }
                break;
            case "\\row":
                commitRow();
                break;
            case "\\intbl":
                state.InTable = true;
                if (currentRow == null) currentRow = new RtfTableRow();
                if (currentCell == null) currentCell = new RtfTableCell();
                break;

            // ── Unicode 字符 ──
            case "\\u":
                if (hasParam)
                {
                    var ucp = param < 0 ? param + 65536 : param;
                    currentRun.Append((Char)ucp);
                    // consume skip char after \uN
                    SkipWhitespace();
                    if (_pos < _rtf.Length && _rtf[_pos] == '\'')
                    {
                        _pos++;
                        if (_pos + 1 < _rtf.Length)
                        {
                            _pos += 2; // skip hex escape
                        }
                    }
                    else if (_pos < _rtf.Length && _rtf[_pos] == '?')
                    {
                        _pos++;
                    }
                }
                break;

            // ── ANSI 十六进制字符 \'XX ──
            case "\\'":
                if (hasParam) currentRun.Append(DecodeAnsi((Byte)(param & 0xFF)));
                break;

            // ── 特殊字符 ──
            case "\\~": currentRun.Append('\u00A0'); break; // non-breaking space
            case "\\-": currentRun.Append('\u00AD'); break; // optional hyphen
            case "\\|": break;                              // formula character
            case "\\:": break;                              // index mark
            case "\\*": break;                              // skip destinations (handled in group parse)
            case "\\lquote": currentRun.Append('\u2018'); break;
            case "\\rquote": currentRun.Append('\u2019'); break;
            case "\\ldblquote": currentRun.Append('\u201C'); break;
            case "\\rdblquote": currentRun.Append('\u201D'); break;
            case "\\emdash": currentRun.Append('\u2014'); break;
            case "\\endash": currentRun.Append('\u2013'); break;
            case "\\bullet": currentRun.Append('\u2022'); break;

            // Everything else: ignore unknown control words
            default: break;
        }
    }
    #endregion

    #region 辅助：解析控制字
    /// <summary>读取当前位置的控制字（包括\\前缀），返回 (word, param, hasParam)</summary>
    private (String word, Int32 param, Boolean hasParam) ReadControlWord()
    {
        if (_pos >= _rtf.Length || _rtf[_pos] != '\\')
            return ("", 0, false);

        _pos++; // skip backslash

        if (_pos >= _rtf.Length) return ("\\", 0, false);

        var ch = _rtf[_pos];

        // Control symbol (single punctuation char)
        if (!Char.IsLetter(ch))
        {
            // special: \'XX ANSI hex char
            if (ch == '\'')
            {
                _pos++;
                if (_pos + 1 < _rtf.Length)
                {
                    var hex = _rtf.Substring(_pos, 2);
                    _pos += 2;
                    if (Int32.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var codePoint))
                    {
                        // Return as a special marker — caller appends the char
                        return ("\\\'", codePoint, true);
                    }
                }
                return ("\\\'", 0, false);
            }

            var sym = "\\" + ch;
            _pos++;
            return (sym, 0, false);
        }

        // Control word: letters
        var start = _pos - 1; // include backslash
        while (_pos < _rtf.Length && Char.IsLetter(_rtf[_pos])) _pos++;
        var kw = _rtf.Substring(start, _pos - start);

        // Optional numeric parameter
        var neg = false;
        if (_pos < _rtf.Length && _rtf[_pos] == '-')
        {
            neg = true;
            _pos++;
        }
        if (_pos < _rtf.Length && Char.IsDigit(_rtf[_pos]))
        {
            var numStart = _pos;
            while (_pos < _rtf.Length && Char.IsDigit(_rtf[_pos])) _pos++;
            var numStr = _rtf.Substring(numStart, _pos - numStart);
            var num = Int32.Parse(numStr);
            if (neg) num = -num;
            // Consume optional trailing space delimiter
            if (_pos < _rtf.Length && _rtf[_pos] == ' ') _pos++;
            return (kw, num, true);
        }
        // Consume optional trailing space delimiter
        if (_pos < _rtf.Length && _rtf[_pos] == ' ') _pos++;
        return (kw, 0, false);
    }

    /// <summary>查看当前位置的控制字（不消耗）</summary>
    private String PeekKeyword()
    {
        var savedPos = _pos;
        if (_pos >= _rtf.Length || _rtf[_pos] != '\\') return "";
        _pos++;
        if (_pos >= _rtf.Length) { _pos = savedPos; return ""; }
        if (!Char.IsLetter(_rtf[_pos])) { _pos = savedPos; return ""; }
        var start = _pos - 1;
        while (_pos < _rtf.Length && Char.IsLetter(_rtf[_pos])) _pos++;
        var kw = _rtf.Substring(start, _pos - start);
        _pos = savedPos;
        return kw;
    }
    #endregion

    #region 辅助：字体表/颜色表/Info
    private void ParseFontTable()
    {
        var depth = 1;
        var fontName = new StringBuilder();
        var inFontGroup = false;

        while (_pos < _rtf.Length && depth > 0)
        {
            var ch = _rtf[_pos];
            if (ch == '{') { depth++; inFontGroup = true; fontName.Clear(); _pos++; continue; }
            if (ch == '}')
            {
                depth--;
                if (inFontGroup && depth >= 1)
                {
                    var name = fontName.ToString().Trim().TrimEnd(';');
                    _fontTable.Add(name);
                    fontName.Clear();
                    inFontGroup = false;
                }
                _pos++;
                continue;
            }
            if (ch == '\\')
            {
                var (word, _, _) = ReadControlWord();
                // skip froman/fswiss etc; only capture font name text
                continue;
            }
            if (inFontGroup) fontName.Append(ch);
            _pos++;
        }
    }

    private void ParseColorTable()
    {
        // RTF 颜色表格式：;R G B; or ;\red255\green0\blue0;
        // 第一个 ; 通常是"自动颜色"占位（无 \red\green\blue），索引 0 = auto
        var r = 0; var g = 0; var b = 0;
        var hasValues = false;
        var depth = 1;
        while (_pos < _rtf.Length && depth > 0)
        {
            var ch = _rtf[_pos];
            if (ch == '{') { depth++; _pos++; continue; }
            if (ch == '}')
            {
                depth--;
                _pos++;
                if (depth == 0) break;
                continue;
            }
            if (ch == ';')
            {
                _colorTable.Add(hasValues ? (r << 16) | (g << 8) | b : -1);
                r = 0; g = 0; b = 0; hasValues = false;
                _pos++;
                continue;
            }
            if (ch == '\\')
            {
                var (word, param, hasParam) = ReadControlWord();
                if (word == "\\red" && hasParam) { r = param; hasValues = true; }
                else if (word == "\\green" && hasParam) { g = param; hasValues = true; }
                else if (word == "\\blue" && hasParam) { b = param; hasValues = true; }
                continue;
            }
            _pos++;
        }
    }

    private void ParseInfo()
    {
        var depth = 1;
        while (_pos < _rtf.Length && depth > 0)
        {
            var ch = _rtf[_pos];
            if (ch == '{') { depth++; _pos++; continue; }
            if (ch == '}') { depth--; _pos++; continue; }
            if (ch == '\\')
            {
                var (word, _, _) = ReadControlWord();
                if (word == "\\title") _title = ReadInfoValue();
                else if (word == "\\author") _author = ReadInfoValue();
                else if (word == "\\subject") _subject = ReadInfoValue();
                continue;
            }
            _pos++;
        }
    }

    private String ReadInfoValue()
    {
        var sb = new StringBuilder();
        while (_pos < _rtf.Length && _rtf[_pos] != '}' && _rtf[_pos] != '\\')
        {
            sb.Append(_rtf[_pos]);
            _pos++;
        }
        return sb.ToString().Trim();
    }

    /// <summary>解析 \pict 组，返回图片对象；消耗到包括结尾 } 在内</summary>
    private RtfImage? ParsePict()
    {
        var format = "wmf";
        var width = 0;
        var height = 0;
        var hexData = new StringBuilder();
        var depth = 1;

        while (_pos < _rtf.Length && depth > 0)
        {
            var ch = _rtf[_pos];
            if (ch == '{') { depth++; _pos++; continue; }
            if (ch == '}')
            {
                depth--;
                _pos++;
                break;
            }
            if (ch == '\\')
            {
                var (word, param, hasParam) = ReadControlWord();
                switch (word)
                {
                    case "\\pngblip": format = "png"; break;
                    case "\\jpegblip": format = "jpg"; break;
                    case "\\emfblip": format = "emf"; break;
                    case "\\wmetafile": format = "wmf"; break;
                    case "\\picw": if (hasParam) width = param; break;
                    case "\\pich": if (hasParam) height = param; break;
                }
                continue;
            }
            // 十六进制数据字符
            if ((ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F'))
            {
                hexData.Append(ch);
                _pos++;
                continue;
            }
            _pos++;
        }

        if (hexData.Length < 2) return null;

        var hex = hexData.ToString();
        var byteCount = hex.Length / 2;
        var bytes = new Byte[byteCount];
        for (var i = 0; i < byteCount; i++)
            bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);

        return new RtfImage { Data = bytes, Format = format, Width = width, Height = height };
    }
    #endregion

    #region 辅助：跳过组/空白
    /// <summary>跳过当前组（消耗包括结尾 } 在内的所有内容）</summary>
    private void SkipGroup()
    {
        var depth = 1;
        while (_pos < _rtf.Length && depth > 0)
        {
            var ch = _rtf[_pos++];
            if (ch == '{') depth++;
            else if (ch == '}') depth--;
        }
        // _pos is now past the closing }
    }

    private void SkipWhitespace()
    {
        while (_pos < _rtf.Length && (_rtf[_pos] == ' ' || _rtf[_pos] == '\t'
               || _rtf[_pos] == '\r' || _rtf[_pos] == '\n'))
            _pos++;
    }
    #endregion

    #region 辅助：构建 Run
    private RtfRun BuildRun(String text, FormatState state, Boolean isLineBreak)
    {
        var fontName = (state.FontIndex >= 0 && state.FontIndex < _fontTable.Count)
            ? _fontTable[state.FontIndex]
            : null;
        var foreColor = (state.ForeColorIndex > 0 && state.ForeColorIndex < _colorTable.Count)
            ? _colorTable[state.ForeColorIndex]
            : -1;
        var backColor = (state.BackColorIndex > 0 && state.BackColorIndex < _colorTable.Count)
            ? _colorTable[state.BackColorIndex]
            : -1;
        return new RtfRun
        {
            Text = text,
            FontName = fontName,
            FontSize = state.FontSize,
            Bold = state.Bold,
            Italic = state.Italic,
            Underline = state.Underline,
            Strikethrough = state.Strikethrough,
            ForeColor = foreColor,
            BackColor = backColor,
            IsLineBreak = isLineBreak,
        };
    }

    private static void ApplyParaFormat(RtfParagraph para, FormatState state)
    {
        para.Alignment = state.Alignment;
        para.LeftIndent = state.LeftIndent;
        para.RightIndent = state.RightIndent;
        para.FirstLineIndent = state.FirstLineIndent;
        para.SpaceBefore = state.SpaceBefore;
        para.SpaceAfter = state.SpaceAfter;
        para.LineSpacing = state.LineSpacing;
        para.InTable = state.InTable;
    }

    /// <summary>将 Windows-1252 字节转换为 Unicode 字符</summary>
    /// <param name="b">ANSI 字节值</param>
    /// <returns>对应的 Unicode 字符</returns>
    private static Char DecodeAnsi(Byte b)
    {
        if (b < 0x80) return (Char)b;
        if (b >= 0xA0) return (Char)b;
        // 0x80-0x9F Windows-1252 私有映射
        return (Char)s_cp1252[b - 0x80];
    }

    private static readonly UInt16[] s_cp1252 =
    [
        0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
        0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
        0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
        0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
    ];
    #endregion
}
