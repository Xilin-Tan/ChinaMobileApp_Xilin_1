using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ReportTool.Core
{
    public class EvalContext
    {
        public Workbook Workbook;
        public Sheet Sheet;
        public int Row;
        public int Col;
        public int Depth;
    }

    /// <summary>
    /// Excel 公式求值器。支持 SUMIFS / SUM / ROUND / RANK / MAX / MIN / ABS / IF 等，
    /// 支持跨表引用、整列引用、绝对/相对引用（引用在解析时已确定，不做相对位移）。
    /// </summary>
    public class Evaluator
    {
        const int MaxDepth = 64;

        public CellValue Evaluate(string formula, EvalContext ctx)
        {
            if (ctx.Depth > MaxDepth) return CellValue.Error("#CIRC!");
            var tokens = new Lexer(formula).Tokenize();
            var p = new Parser(tokens);
            var node = p.ParseExpression();
            if (node == null) return CellValue.Error("#PARSE!");
            return EvalNode(node, ctx);
        }

        // 预编译：公式文本只解析一次，多轮迭代时省掉重复词法/语法分析。
        public class Compiled
        {
            internal Node Root;
            public string Error;
        }

        public Compiled Compile(string formula)
        {
            var c = new Compiled();
            try
            {
                var tokens = new Lexer(formula).Tokenize();
                c.Root = new Parser(tokens).ParseExpression();
                if (c.Root == null) c.Error = "#PARSE!";
            }
            catch (Exception ex)
            {
                c.Error = "#PARSE:" + ex.GetType().Name;
            }
            return c;
        }

        public CellValue Evaluate(Compiled compiled, EvalContext ctx)
        {
            if (ctx.Depth > MaxDepth) return CellValue.Error("#CIRC!");
            if (compiled == null || compiled.Root == null) return CellValue.Error(compiled != null ? compiled.Error : "#PARSE!");
            return EvalNode(compiled.Root, ctx);
        }

        // ---------- AST ----------
        internal abstract class Node { }
        class NumNode : Node { public double Value; }
        class StrNode : Node { public string Value; }
        class RefNode : Node { public Token Tok; }
        class UnaryNode : Node { public string Op; public Node Operand; }
        class BinNode : Node { public string Op; public Node Left, Right; }
        class FuncNode : Node { public string Name; public List<Node> Args = new List<Node>(); }

        class Parser
        {
            readonly List<Token> _t;
            int _i;
            public Parser(List<Token> t) { _t = t; }

            Token Cur { get { return _i < _t.Count ? _t[_i] : new Token { Type = TokType.End }; } }
            Token Peek(int n) { int k = _i + n; return k < _t.Count ? _t[k] : new Token { Type = TokType.End }; }
            void Advance() { if (_i < _t.Count) _i++; }

            public Node ParseExpression() { return ParseCompare(); }

            Node ParseCompare()
            {
                var left = ParseConcat();
                while (Cur.Type == TokType.Op && (Cur.Text == "=" || Cur.Text == "<>" || Cur.Text == "<" || Cur.Text == ">" || Cur.Text == "<=" || Cur.Text == ">="))
                {
                    string op = Cur.Text; Advance();
                    var right = ParseConcat();
                    left = new BinNode { Op = op, Left = left, Right = right };
                }
                return left;
            }

            Node ParseConcat()
            {
                var left = ParseAdd();
                while (Cur.Type == TokType.Op && Cur.Text == "&")
                {
                    Advance();
                    var right = ParseAdd();
                    left = new BinNode { Op = "&", Left = left, Right = right };
                }
                return left;
            }

            Node ParseAdd()
            {
                var left = ParseMul();
                while (Cur.Type == TokType.Op && (Cur.Text == "+" || Cur.Text == "-"))
                {
                    string op = Cur.Text; Advance();
                    var right = ParseMul();
                    left = new BinNode { Op = op, Left = left, Right = right };
                }
                return left;
            }

            Node ParseMul()
            {
                var left = ParseUnary();
                while (Cur.Type == TokType.Op && (Cur.Text == "*" || Cur.Text == "/" || Cur.Text == "^"))
                {
                    string op = Cur.Text; Advance();
                    var right = ParseUnary();
                    left = new BinNode { Op = op, Left = left, Right = right };
                }
                return left;
            }

            Node ParseUnary()
            {
                if (Cur.Type == TokType.Op && (Cur.Text == "-" || Cur.Text == "+"))
                {
                    string op = Cur.Text; Advance();
                    return new UnaryNode { Op = op, Operand = ParseUnary() };
                }
                if (Cur.Type == TokType.Op && Cur.Text == "%")
                {
                    Advance();
                    return new NumNode { Value = 0 };
                }
                return ParseAtom();
            }

            Node ParseAtom()
            {
                var t = Cur;
                switch (t.Type)
                {
                    case TokType.Number:
                        Advance();
                        return new NumNode { Value = t.Num };
                    case TokType.String:
                        Advance();
                        return new StrNode { Value = t.Text };
                    case TokType.CellRef:
                    case TokType.RangeRef:
                        Advance();
                        return new RefNode { Tok = t };
                    case TokType.Ident:
                        {
                            // 可能是 SheetName!A1 —— 但 SheetName 已被当作 Ident
                            string name = t.Text;
                            Advance();
                            if (Cur.Type == TokType.Bang)
                            {
                                Advance();
                                var refTok = Cur;
                                if (refTok.Type == TokType.CellRef || refTok.Type == TokType.RangeRef)
                                {
                                    Advance();
                                    refTok.SheetName = name;
                                    return new RefNode { Tok = refTok };
                                }
                                return new NumNode { Value = 0 };
                            }
                            if (Cur.Type == TokType.LParen)
                            {
                                Advance();
                                var fn = new FuncNode { Name = name.ToUpperInvariant() };
                                if (Cur.Type != TokType.RParen)
                                {
                                    while (true)
                                    {
                                        var arg = ParseExpression();
                                        if (arg != null) fn.Args.Add(arg);
                                        if (Cur.Type == TokType.Comma) { Advance(); continue; }
                                        break;
                                    }
                                }
                                if (Cur.Type == TokType.RParen) Advance();
                                return fn;
                            }
                            // 单独标识符（TRUE/FALSE）
                            if (name == "TRUE") return new NumNode { Value = 1 };
                            if (name == "FALSE") return new NumNode { Value = 0 };
                            return new NumNode { Value = 0 };
                        }
                    case TokType.LParen:
                        {
                            Advance();
                            var inner = ParseExpression();
                            if (Cur.Type == TokType.RParen) Advance();
                            return inner;
                        }
                }
                Advance();
                return new NumNode { Value = 0 };
            }
        }

        // ---------- 求值 ----------
        CellValue EvalNode(Node n, EvalContext ctx)
        {
            var num = n as NumNode;
            if (num != null) return CellValue.FromNumber(num.Value);
            var str = n as StrNode;
            if (str != null) return CellValue.FromText(str.Value);

            var un = n as UnaryNode;
            if (un != null)
            {
                var v = EvalNode(un.Operand, ctx);
                if (v.Kind == ValueKind.Error) return v;
                double d = v.ToNumber();
                return CellValue.FromNumber(un.Op == "-" ? -d : d);
            }

            var bn = n as BinNode;
            if (bn != null) return EvalBinary(bn, ctx);

            var rf = n as RefNode;
            if (rf != null)
            {
                return ReadRef(rf.Tok, ctx);
            }

            var fn = n as FuncNode;
            if (fn != null) return EvalFunc(fn, ctx);

            return CellValue.Blank;
        }

        CellValue EvalBinary(BinNode bn, EvalContext ctx)
        {
            if (bn.Op == "&")
            {
                var l = EvalNode(bn.Left, ctx);
                var r = EvalNode(bn.Right, ctx);
                return CellValue.FromText(ToDisplay(l) + ToDisplay(r));
            }

            if (bn.Op == "=" || bn.Op == "<>" || bn.Op == "<" || bn.Op == ">" || bn.Op == "<=" || bn.Op == ">=")
            {
                var l = EvalNode(bn.Left, ctx);
                var r = EvalNode(bn.Right, ctx);
                bool cmp;
                if (l.IsNumeric && r.IsNumeric)
                {
                    double a = l.ToNumber(), b = r.ToNumber();
                    cmp = CompareOp(bn.Op, a.CompareTo(b));
                }
                else
                {
                    string a = ToDisplay(l), b = ToDisplay(r);
                    cmp = CompareOp(bn.Op, string.Compare(a, b, StringComparison.OrdinalIgnoreCase));
                }
                return CellValue.FromBool(cmp);
            }

            var lv = EvalNode(bn.Left, ctx);
            if (lv.Kind == ValueKind.Error) return lv;
            var rv = EvalNode(bn.Right, ctx);
            if (rv.Kind == ValueKind.Error) return rv;

            double x = lv.ToNumber(), y = rv.ToNumber();
            switch (bn.Op)
            {
                case "+": return CellValue.FromNumber(x + y);
                case "-": return CellValue.FromNumber(x - y);
                case "*": return CellValue.FromNumber(x * y);
                case "/": return y == 0 ? CellValue.Error("#DIV/0!") : CellValue.FromNumber(x / y);
                case "^": return CellValue.FromNumber(Math.Pow(x, y));
            }
            return CellValue.Error("#OP!");
        }

        static bool CompareOp(string op, int c)
        {
            switch (op)
            {
                case "=": return c == 0;
                case "<>": return c != 0;
                case "<": return c < 0;
                case ">": return c > 0;
                case "<=": return c <= 0;
                case ">=": return c >= 0;
            }
            return false;
        }

        static string ToDisplay(CellValue v)
        {
            switch (v.Kind)
            {
                case ValueKind.Number: return CellValue.FormatNumber(v.Num);
                case ValueKind.Bool: return v.Bool ? "TRUE" : "FALSE";
                case ValueKind.Text: return v.Text ?? "";
                default: return "";
            }
        }

        Sheet ResolveSheet(string name, EvalContext ctx)
        {
            if (string.IsNullOrEmpty(name)) return ctx.Sheet;
            Sheet s;
            if (ctx.Workbook != null && ctx.Workbook.Sheets.TryGetValue(name, out s)) return s;
            return ctx.Sheet;
        }

        CellValue ReadRef(Token tok, EvalContext ctx)
        {
            var sheet = ResolveSheet(tok.SheetName, ctx);
            if (sheet == null) return CellValue.Blank;
            return sheet.Get(tok.Row, tok.Col);
        }

        // 把 ref 展开成 (row,col) 集合
        List<CellValue> CollectValues(Token tok, EvalContext ctx, out Sheet sheet)
        {
            sheet = ResolveSheet(tok.SheetName, ctx);
            var list = new List<CellValue>();
            if (sheet == null) return list;

            int r1 = tok.Row, c1 = tok.Col, r2 = tok.Row2, c2 = tok.Col2;
            if (tok.Type == TokType.CellRef) { r2 = r1; c2 = c1; }
            if (r1 == 0) r1 = 1;
            if (r2 == 0) r2 = sheet.MaxRow;
            if (r1 > r2) { int t = r1; r1 = r2; r2 = t; }
            if (c1 > c2) { int t = c1; c1 = c2; c2 = t; }

            long total = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (total > 5000000) { r2 = Math.Min(r2, r1 + 1000000 / Math.Max(1, c2 - c1 + 1)); }

            for (int c = c1; c <= c2; c++)
                for (int r = r1; r <= r2; r++)
                    list.Add(sheet.Get(r, c));
            return list;
        }

        double SumValues(List<CellValue> vals)
        {
            double s = 0;
            foreach (var v in vals) if (v.Kind == ValueKind.Number) s += v.Num;
            return s;
        }

        CellValue EvalFunc(FuncNode fn, EvalContext ctx)
        {
            switch (fn.Name)
            {
                case "SUM":
                    {
                        double total = 0;
                        foreach (var a in fn.Args)
                        {
                            var rf = a as RefNode;
                            if (rf != null) { Sheet sh; total += SumValues(CollectValues(rf.Tok, ctx, out sh)); }
                            else { var v = EvalNode(a, ctx); if (v.Kind == ValueKind.Error) return v; if (v.IsNumeric) total += v.ToNumber(); }
                        }
                        return CellValue.FromNumber(total);
                    }

                case "ROUND":
                    {
                        if (fn.Args.Count == 0) return CellValue.FromNumber(0);
                        var v = EvalNode(fn.Args[0], ctx);
                        if (v.Kind == ValueKind.Error) return v;
                        double x = v.ToNumber();
                        int digits = 0;
                        if (fn.Args.Count > 1) digits = (int)EvalNode(fn.Args[1], ctx).ToNumber();
                        if (digits > 15) digits = 15;
                        if (digits < -15) digits = -15;
                        return CellValue.FromNumber(Math.Round(x, digits, MidpointRounding.AwayFromZero));
                    }

                case "INT":
                    return CellValue.FromNumber(Math.Floor(EvalNode(fn.Args[0], ctx).ToNumber()));

                case "ABS":
                    return CellValue.FromNumber(Math.Abs(EvalNode(fn.Args[0], ctx).ToNumber()));

                case "MAX":
                case "MIN":
                    {
                        double best = fn.Name == "MAX" ? double.NegativeInfinity : double.PositiveInfinity;
                        bool any = false;
                        foreach (var a in fn.Args)
                        {
                            var rf = a as RefNode;
                            if (rf != null)
                            {
                                Sheet sh;
                                foreach (var v in CollectValues(rf.Tok, ctx, out sh))
                                {
                                    if (v.Kind != ValueKind.Number) continue;
                                    any = true;
                                    if (fn.Name == "MAX") { if (v.Num > best) best = v.Num; }
                                    else { if (v.Num < best) best = v.Num; }
                                }
                            }
                            else
                            {
                                var v = EvalNode(a, ctx);
                                if (v.IsNumeric)
                                {
                                    any = true;
                                    double d = v.ToNumber();
                                    if (fn.Name == "MAX") { if (d > best) best = d; }
                                    else { if (d < best) best = d; }
                                }
                            }
                        }
                        return CellValue.FromNumber(any ? best : 0);
                    }

                case "RANK":
                case "RANK.EQ":
                    {
                        if (fn.Args.Count < 2) return CellValue.Error("#ARG!");
                        double x = EvalNode(fn.Args[0], ctx).ToNumber();
                        var rangeRef = fn.Args[1] as RefNode;
                        if (rangeRef == null) return CellValue.Error("#ARG!");
                        int order = 0;
                        if (fn.Args.Count > 2) order = (int)EvalNode(fn.Args[2], ctx).ToNumber();

                        Sheet sh;
                        var vals = CollectValues(rangeRef.Tok, ctx, out sh);
                        int better = 0;
                        bool found = false;
                        foreach (var v in vals)
                        {
                            if (v.Kind != ValueKind.Number) continue;
                            if (Math.Abs(v.Num - x) < 1e-9) { found = true; continue; }
                            if (order == 0) { if (v.Num > x) better++; }
                            else { if (v.Num < x) better++; }
                        }
                        if (!found) return CellValue.Error("#N/A");
                        return CellValue.FromNumber(better + 1);
                    }

                case "SUMIFS":
                    return EvalSumIfs(fn, ctx);

                case "COUNTIFS":
                    {
                        int cnt = 0;
                        int pairs = fn.Args.Count / 2;
                        Sheet baseSheet = null;
                        int rows = 0;
                        for (int i = 0; i < pairs; i++)
                        {
                            var rf = fn.Args[i * 2] as RefNode;
                            if (rf == null) return CellValue.Error("#ARG!");
                            Sheet sh;
                            var vals = CollectValues(rf.Tok, ctx, out sh);
                            if (baseSheet == null) { baseSheet = sh; rows = vals.Count; }
                        }
                        if (baseSheet == null) return CellValue.FromNumber(0);
                        for (int r = 0; r < rows; r++)
                        {
                            bool ok = true;
                            for (int i = 0; i < pairs; i++)
                            {
                                var rf = (RefNode)fn.Args[i * 2];
                                Sheet sh; var vals = CollectValues(rf.Tok, ctx, out sh);
                                if (r >= vals.Count) { ok = false; break; }
                                if (!MatchCriteria(vals[r], fn.Args[i * 2 + 1], ctx)) { ok = false; break; }
                            }
                            if (ok) cnt++;
                        }
                        return CellValue.FromNumber(cnt);
                    }

                case "IF":
                    {
                        if (fn.Args.Count < 2) return CellValue.Error("#ARG!");
                        var cond = EvalNode(fn.Args[0], ctx);
                        bool truth = cond.Kind == ValueKind.Bool ? cond.Bool : cond.ToNumber() != 0;
                        if (truth) return EvalNode(fn.Args[1], ctx);
                        return fn.Args.Count > 2 ? EvalNode(fn.Args[2], ctx) : CellValue.FromBool(false);
                    }

                case "IFERROR":
                    {
                        var v = EvalNode(fn.Args[0], ctx);
                        if (v.Kind == ValueKind.Error) return fn.Args.Count > 1 ? EvalNode(fn.Args[1], ctx) : CellValue.FromNumber(0);
                        return v;
                    }

                case "COUNTA":
                    {
                        int cnt = 0;
                        foreach (var a in fn.Args)
                        {
                            var rf = a as RefNode;
                            if (rf != null) { Sheet sh; foreach (var v in CollectValues(rf.Tok, ctx, out sh)) if (v.Kind != ValueKind.Blank) cnt++; }
                            else if (EvalNode(a, ctx).Kind != ValueKind.Blank) cnt++;
                        }
                        return CellValue.FromNumber(cnt);
                    }
            }

            // 未知函数：返回 0 但标记，便于排查
            return CellValue.Error("#FUNC:" + fn.Name);
        }

        CellValue EvalSumIfs(FuncNode fn, EvalContext ctx)
        {
            if (fn.Args.Count < 3) return CellValue.Error("#ARG!");

            var plan = BuildPlan(fn, ctx);
            if (plan != null) return RunPlan(plan);

            var sumRef = fn.Args[0] as RefNode;
            if (sumRef == null) return CellValue.Error("#ARG!");

            Sheet sumSheet;
            var sumVals = CollectValues(sumRef.Tok, ctx, out sumSheet);

            int pairs = (fn.Args.Count - 1) / 2;
            var critVals = new List<List<CellValue>>();
            var critNodes = new List<Node>();
            for (int i = 0; i < pairs; i++)
            {
                var rf = fn.Args[1 + i * 2] as RefNode;
                if (rf == null) return CellValue.Error("#ARG!");
                Sheet sh;
                critVals.Add(CollectValues(rf.Tok, ctx, out sh));
                critNodes.Add(fn.Args[2 + i * 2]);
            }

            double total = 0;
            for (int r = 0; r < sumVals.Count; r++)
            {
                bool ok = true;
                for (int i = 0; i < pairs; i++)
                {
                    if (r >= critVals[i].Count) { ok = false; break; }
                    if (!MatchCriteria(critVals[i][r], critNodes[i], ctx)) { ok = false; break; }
                }
                if (!ok) continue;
                var sv = sumVals[r];
                if (sv.Kind == ValueKind.Number) total += sv.Num;
            }
            return CellValue.FromNumber(total);
        }
        // 整列引用且目标表无公式时，直接拿缓存数组（按下标=行号），避免每轮重复建 List。
        CellValue[] RowArray(Token tok, EvalContext ctx, out int maxRow)
        {
            maxRow = 0;
            if (tok == null || tok.Type != TokType.RangeRef) return null;
            if (tok.Row != 0 || tok.Row2 != 0 || tok.Col != tok.Col2) return null;
            var sheet = ResolveSheet(tok.SheetName, ctx);
            if (sheet == null) return null;
            var arr = sheet.FastColumn(tok.Col);
            if (arr == null) return null;
            maxRow = sheet.MaxRow;
            return arr;
        }

        // Fast path parameters: used when every argument is a plain full-column reference
        // against a formula-free sheet, so array index == row number.
        class SumIfsPlan
        {
            public CellValue[] Sum;
            public CellValue[][] Crit;
            public Crit[] Rules;
            public int MaxRow;
        }

        SumIfsPlan BuildPlan(FuncNode fn, EvalContext ctx)
        {
            var sumRef = fn.Args[0] as RefNode;
            if (sumRef == null) return null;
            int maxRow, ignored;
            var sumArr = RowArray(sumRef.Tok, ctx, out maxRow);
            if (sumArr == null) return null;

            int pairs = (fn.Args.Count - 1) / 2;
            var plan = new SumIfsPlan
            {
                Sum = sumArr,
                MaxRow = maxRow,
                Crit = new CellValue[pairs][],
                Rules = new Crit[pairs]
            };
            for (int i = 0; i < pairs; i++)
            {
                var rf = fn.Args[1 + i * 2] as RefNode;
                if (rf == null) return null;
                var arr = RowArray(rf.Tok, ctx, out ignored);
                if (arr == null) return null;
                plan.Crit[i] = arr;

                var critVal = EvalNode(fn.Args[2 + i * 2], ctx);
                if (critVal.Kind == ValueKind.Error) return null;
                plan.Rules[i] = ParseCrit(ToDisplay(critVal));
            }
            return plan;
        }

        static CellValue RunPlan(SumIfsPlan plan)
        {
            int pairs = plan.Crit.Length;
            int last = Math.Min(plan.MaxRow, plan.Sum.Length - 1);
            double total = 0;
            for (int r = 1; r <= last; r++)
            {
                var sv = plan.Sum[r];
                if (sv.Kind != ValueKind.Number) continue;

                bool ok = true;
                for (int i = 0; i < pairs; i++)
                {
                    var arr = plan.Crit[i];
                    if (r >= arr.Length) { ok = false; break; }
                    if (!MatchParsed(arr[r], plan.Rules[i])) { ok = false; break; }
                }
                if (ok) total += sv.Num;
            }
            return CellValue.FromNumber(total);
        }

        bool MatchCriteria(CellValue cellVal, Node critNode, EvalContext ctx)
        {
            var critVal = EvalNode(critNode, ctx);
            if (critVal.Kind == ValueKind.Error) return false;

            string critText = ToDisplay(critVal);
            return MatchCriteriaText(cellVal, critText);
        }

        bool MatchCriteriaText(CellValue cellVal, string critText)
        {
            return MatchParsed(cellVal, ParseCrit(critText));
        }

        struct Crit
        {
            public string Op;
            public string Rest;
            public double Num;
            public bool IsNum;
            public bool HasWild;
        }

        static Crit ParseCrit(string text)
        {
            var c = new Crit();
            c.Rest = text ?? "";
            if (text == null) return c;
            if (text.Length >= 2 && (text.StartsWith(">=") || text.StartsWith("<=") || text.StartsWith("<>")))
            {
                c.Op = text.Substring(0, 2);
                c.Rest = text.Substring(2);
            }
            else if (text.Length >= 1 && (text[0] == '>' || text[0] == '<' || text[0] == '='))
            {
                c.Op = text.Substring(0, 1);
                c.Rest = text.Substring(1);
            }
            double d;
            c.IsNum = double.TryParse(c.Rest, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
            c.Num = d;
            c.HasWild = c.Rest.IndexOf('*') >= 0 || c.Rest.IndexOf('?') >= 0;
            return c;
        }

        static bool MatchParsed(CellValue cellVal, Crit c)
        {
            string critText = c.Rest;
            if (c.Op != null)
            {
                if (c.IsNum && cellVal.Kind == ValueKind.Number)
                    return CompareOp(c.Op, cellVal.Num.CompareTo(c.Num));
                if (c.IsNum && cellVal.Kind == ValueKind.Blank)
                    return CompareOp(c.Op, 0.0.CompareTo(c.Num));
                if (c.Op == "<>" && !c.IsNum)
                    return !string.Equals(ToDisplay(cellVal), critText, StringComparison.OrdinalIgnoreCase);
                if (c.Op == "=") return string.Equals(ToDisplay(cellVal), critText, StringComparison.OrdinalIgnoreCase);
                return false;
            }
            if (c.HasWild) return WildcardMatch(ToDisplay(cellVal), critText);
            if (cellVal.Kind == ValueKind.Blank && critText.Length == 0) return true;
            return string.Equals(ToDisplay(cellVal), critText, StringComparison.OrdinalIgnoreCase);
        }

        // 旧实现保留作对照（不再被调用）
        bool MatchCriteriaTextLegacy(CellValue cellVal, string critText)
        {
            // 形如 ">=100" / "<>abc" 的比较条件
            if (critText.Length >= 2)
            {
                string op = null;
                string rest = null;
                if (critText.StartsWith(">=") || critText.StartsWith("<=") || critText.StartsWith("<>"))
                { op = critText.Substring(0, 2); rest = critText.Substring(2); }
                else if (critText[0] == '>' || critText[0] == '<' || critText[0] == '=')
                { op = critText.Substring(0, 1); rest = critText.Substring(1); }

                if (op != null)
                {
                    double numRest;
                    bool isNumRest = double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture, out numRest);
                    if (isNumRest && cellVal.Kind == ValueKind.Number)
                        return CompareOp(op, cellVal.Num.CompareTo(numRest));
                    if (isNumRest && cellVal.Kind == ValueKind.Blank)
                        return CompareOp(op, 0.0.CompareTo(numRest));
                    if (op == "<>" && !isNumRest)
                        return !string.Equals(ToDisplay(cellVal), rest, StringComparison.OrdinalIgnoreCase);
                    if (op == "=") return string.Equals(ToDisplay(cellVal), rest, StringComparison.OrdinalIgnoreCase);
                    return false;
                }
            }

            // 通配符
            if (critText.IndexOf('*') >= 0 || critText.IndexOf('?') >= 0)
                return WildcardMatch(ToDisplay(cellVal), critText);

            if (cellVal.Kind == ValueKind.Blank && critText.Length == 0) return true;
            return string.Equals(ToDisplay(cellVal), critText, StringComparison.OrdinalIgnoreCase);
        }

        static bool WildcardMatch(string text, string pattern)
        {
            int ti = 0, pi = 0, star = -1, match = 0;
            while (ti < text.Length)
            {
                if (pi < pattern.Length && (pattern[pi] == '?' || char.ToUpperInvariant(pattern[pi]) == char.ToUpperInvariant(text[ti])))
                { ti++; pi++; }
                else if (pi < pattern.Length && pattern[pi] == '*')
                { star = pi++; match = ti; }
                else if (star >= 0)
                { pi = star + 1; ti = ++match; }
                else return false;
            }
            while (pi < pattern.Length && pattern[pi] == '*') pi++;
            return pi == pattern.Length;
        }
    }
}
