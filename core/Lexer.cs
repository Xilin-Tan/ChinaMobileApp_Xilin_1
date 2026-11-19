using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ReportTool.Core
{
    public enum TokType { Number, String, Ident, CellRef, RangeRef, Op, LParen, RParen, Comma, Bang, End }

    public class Token
    {
        public TokType Type;
        public string Text;
        public double Num;
        public int Row, Col, Row2, Col2;
        public bool RowAbs, ColAbs, Row2Abs, Col2Abs;
        public string SheetName;
        public int Pos;
        public override string ToString() { return Type + ":" + Text; }
    }

    /// <summary>
    /// Excel 公式词法分析。支持数字、字符串、单元格引用（$A$1 / A1）、
    /// 整列引用（A:A）、区间（A1:B2 / $D$5:$D$12）、运算符与函数名。
    /// 关键点：引用解析必须正确推进游标，否则 "$D$5:$D$12" 会被错误拆解。
    /// </summary>
    public class Lexer
    {
        readonly string _s;
        int _i;

        public Lexer(string s) { _s = s ?? ""; _i = 0; }

        public List<Token> Tokenize()
        {
            var list = new List<Token>();
            while (true)
            {
                SkipWs();
                if (_i >= _s.Length) { list.Add(new Token { Type = TokType.End, Pos = _i }); break; }
                char c = _s[_i];
                int start = _i;

                if (c == '"')
                {
                    _i++;
                    var sb = new StringBuilder();
                    while (_i < _s.Length)
                    {
                        if (_s[_i] == '"')
                        {
                            if (_i + 1 < _s.Length && _s[_i + 1] == '"') { sb.Append('"'); _i += 2; continue; }
                            _i++; break;
                        }
                        sb.Append(_s[_i]); _i++;
                    }
                    list.Add(new Token { Type = TokType.String, Text = sb.ToString(), Pos = start });
                    continue;
                }

                if (c == '(') { _i++; list.Add(new Token { Type = TokType.LParen, Text = "(", Pos = start }); continue; }
                if (c == ')') { _i++; list.Add(new Token { Type = TokType.RParen, Text = ")", Pos = start }); continue; }
                if (c == ',') { _i++; list.Add(new Token { Type = TokType.Comma, Text = ",", Pos = start }); continue; }
                if (c == '!') { _i++; list.Add(new Token { Type = TokType.Bang, Text = "!", Pos = start }); continue; }
                if (c == '%') { _i++; list.Add(new Token { Type = TokType.Op, Text = "%", Pos = start }); continue; }

                if (c == '+' || c == '-' || c == '*' || c == '/' || c == '&' || c == '^'
                    || c == '=' || c == '<' || c == '>')
                {
                    string op = c.ToString();
                    _i++;
                    if ((c == '<' || c == '>') && _i < _s.Length && _s[_i] == '=') { op += "="; _i++; }
                    if (c == '<' && _i < _s.Length && _s[_i] == '>') { op = "<>"; _i++; }
                    list.Add(new Token { Type = TokType.Op, Text = op, Pos = start });
                    continue;
                }

                if (char.IsDigit(c) || (c == '.' && _i + 1 < _s.Length && char.IsDigit(_s[_i + 1])))
                {
                    while (_i < _s.Length && (char.IsDigit(_s[_i]) || _s[_i] == '.')) _i++;
                    if (_i < _s.Length && (_s[_i] == 'e' || _s[_i] == 'E'))
                    {
                        int save = _i;
                        _i++;
                        if (_i < _s.Length && (_s[_i] == '+' || _s[_i] == '-')) _i++;
                        if (_i < _s.Length && char.IsDigit(_s[_i])) { while (_i < _s.Length && char.IsDigit(_s[_i])) _i++; }
                        else _i = save;
                    }
                    string numText = _s.Substring(start, _i - start);
                    double d = double.Parse(numText, NumberStyles.Float, CultureInfo.InvariantCulture);
                    list.Add(new Token { Type = TokType.Number, Text = numText, Num = d, Pos = start });
                    continue;
                }

                if (char.IsLetter(c) || c == '_' || c == '$')
                {
                    var tok = ReadNameOrRef();
                    if (tok != null) { list.Add(tok); continue; }
                }

                _i++;   // 未知字符，跳过，绝不原地打转
            }
            return list;
        }

        static bool IsColLetters(string s)
        {
            if (s.Length < 1 || s.Length > 3) return false;
            foreach (char ch in s) if (ch < 'A' || ch > 'Z') return false;
            return true;
        }

        Token ReadNameOrRef()
        {
            int start = _i;
            int cursor = _i;

            bool colAbs = false;
            if (cursor < _s.Length && _s[cursor] == '$') { colAbs = true; cursor++; }

            int lettersStart = cursor;
            while (cursor < _s.Length && char.IsLetter(_s[cursor])) cursor++;
            string letters = _s.Substring(lettersStart, cursor - lettersStart);

            if (letters.Length == 0)
            {
                _i = start;
                return null;
            }

            // 函数名：字母后（可跳空白）紧跟 "(" —— 一律按函数名处理。
            // 必须放在列名判断之前：MAX/MIN/SUM 这类 1~3 字母的函数名同时也是合法列名。
            int lookahead = cursor;
            while (lookahead < _s.Length && char.IsWhiteSpace(_s[lookahead])) lookahead++;
            if (lookahead < _s.Length && _s[lookahead] == '(')
            {
                _i = lookahead;
                return new Token { Type = TokType.Ident, Text = letters.ToUpperInvariant(), Pos = start };
            }

            if (!IsColLetters(letters.ToUpperInvariant()))
            {
                // 不是引用形态（如 TRUE / FALSE / 表名）
                _i = cursor;
                return new Token { Type = TokType.Ident, Text = letters.ToUpperInvariant(), Pos = start };
            }

            int col = XlsxReader.ColIndex(letters);
            bool rowAbs = false;
            int rowCursor = cursor;
            if (rowCursor < _s.Length && _s[rowCursor] == '$') { rowAbs = true; rowCursor++; }
            int digitsStart = rowCursor;
            while (rowCursor < _s.Length && char.IsDigit(_s[rowCursor])) rowCursor++;

            Token first;
            if (rowCursor > digitsStart)
            {
                int row = int.Parse(_s.Substring(digitsStart, rowCursor - digitsStart), CultureInfo.InvariantCulture);
                first = new Token
                {
                    Type = TokType.CellRef,
                    Text = _s.Substring(start, rowCursor - start),
                    Row = row, Col = col,
                    RowAbs = rowAbs, ColAbs = colAbs,
                    Pos = start
                };
                _i = rowCursor;
            }
            else
            {
                // 整列引用 A:A（A 后面直接是冒号）
                first = new Token
                {
                    Type = TokType.RangeRef,
                    Text = _s.Substring(start, cursor - start),
                    Row = 0, Col = col, Row2 = 0, Col2 = col,
                    ColAbs = colAbs, Col2Abs = colAbs,
                    Pos = start
                };
                _i = cursor;
            }

            TryRangeSuffix(first);
            return first;
        }

        /// <summary>若当前游标在 ':' 上，尝试把 first 扩展成区间引用。</summary>
        void TryRangeSuffix(Token first)
        {
            if (_i >= _s.Length || _s[_i] != ':') return;
            int save = _i;
            int cursor = _i + 1;

            bool col2Abs = false;
            if (cursor < _s.Length && _s[cursor] == '$') { col2Abs = true; cursor++; }

            int lettersStart = cursor;
            while (cursor < _s.Length && char.IsLetter(_s[cursor])) cursor++;
            string letters = _s.Substring(lettersStart, cursor - lettersStart);
            if (!IsColLetters(letters.ToUpperInvariant())) { _i = save; return; }

            int col2 = XlsxReader.ColIndex(letters);

            bool row2Abs = false;
            int rowCursor = cursor;
            if (rowCursor < _s.Length && _s[rowCursor] == '$') { row2Abs = true; rowCursor++; }
            int digitsStart = rowCursor;
            while (rowCursor < _s.Length && char.IsDigit(_s[rowCursor])) rowCursor++;

            int row2 = 0;
            if (rowCursor > digitsStart)
                row2 = int.Parse(_s.Substring(digitsStart, rowCursor - digitsStart), CultureInfo.InvariantCulture);

            first.Type = TokType.RangeRef;
            first.Row2 = row2;
            first.Col2 = col2;
            first.Row2Abs = row2Abs;
            first.Col2Abs = col2Abs;
            _i = rowCursor;
        }

        void SkipWs()
        {
            while (_i < _s.Length && char.IsWhiteSpace(_s[_i])) _i++;
        }
    }
}