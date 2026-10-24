// 模糊匹配：决定"打的这几个字命中哪几个候选、命中在哪几位、排第几档"。
//
// 匹配语义的规则，逐条列出：
//   1) 优先【整块连续】——query 作为子串出现时直接用它（视觉上就是"我打的
//      这几个字母连在一起"）。arr -> dataArr 命中 Arr，而不是散着的 a..r..r。
//   2) 否则逐字符贪心，优先落在词边界上（首字母缩略型：ds -> dataSheet）。
//   3) 贪心必须带【反向可行性上界】，否则会漏判：只看"这一档更优"、不管后面
//      还凑不凑得齐，query 明明是子序列也会被判成不匹配。
//      真实案例：getSettingTitle3 里输入 getse3 —— 第 3 个字符 t 贪心跳到词边界
//      T(10)（档 1 优于接着用的 t(2) 档 2），再往后就没有 s 了；可
//      [0,1,2,3,4,15] 本来就成立。
//   4) 命中位置是【逐字符】的一组下标，不是一段区间 —— UI 要拿它逐字符上色。
//      只给 (pos, span) 一段的话，dataArr 打 da 会把 da 整段标红、
//      打 ds 只能标一段，看起来跟 IDEA 完全不是一个东西。
//
// 刻意不依赖任何 VBA 知识：这里只做"字符串匹配"，类别判定在 Engine。

using System;
using System.Collections.Generic;

namespace VbeComplete
{
    /// <summary>一次匹配的结果。Positions 是命中字符的下标（升序）。</summary>
    internal struct FuzzyHit
    {
        public int[] Positions;
        public int Kind;      // 4 完全相等 / 3 前缀 / 2 词首缩略 / 1 连续块 / 0 分散
        public int First;     // 首个命中下标
        public int Span;      // 首尾跨度（-1 表示只有一个字符）
        public int Run;       // 最长连续段长度（判断"是不是真在打这个名字"用）
    }

    internal static class Fuzzy
    {
        /// <summary>
        /// 名字里的"词边界"下标集合。
        /// 边界 = 首字符 / 下划线之后 / 小写→大写的切换处（camelCase 驼峰）。
        /// 例如 dataSheet -> {0, 4}，DataArr -> {0, 4}。
        ///
        /// 必须用原始 name，不能先小写 —— 大小写切换信息一丢，就再也认不出
        ///   camelCase 的词首了，`ds` 想命中 `dataSheet` 的 D 和 S 就无从谈起。
        /// </summary>
        public static HashSet<int> WordBoundaries(string name)
        {
            var b = new HashSet<int>();
            if (string.IsNullOrEmpty(name)) return b;
            b.Add(0);
            char prev = name[0];
            for (int i = 1; i < name.Length; i++)
            {
                char ch = name[i];
                if (prev == '_' || (char.IsLower(prev) && char.IsUpper(ch))) b.Add(i);
                prev = ch;
            }
            return b;
        }

        /// <summary>
        /// 小写化，但**保证长度不变**。
        /// fuzzy 返回的下标要拿去索引原始名字（UI 高亮、Kind/Span 都按它算），
        /// 所以 low 必须与 name 逐字符对齐。C# 的 char.ToLowerInvariant 是逐字符的、
        /// 一定不变长；这里显式写出来是为了让"长度必须相等"这条约束有个落点
        /// （有些语言做大小写折叠会变长，才需要特殊处理）。
        /// </summary>
        public static string LowerKeepLen(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var buf = new char[s.Length];
            for (int i = 0; i < s.Length; i++) buf[i] = char.ToLowerInvariant(s[i]);
            return new string(buf);
        }

        /// <summary>
        /// 无法整块连续命中时，逐字符挑位置：优先词边界，其次紧接上一个命中。
        /// 档位越小越优先（同档取更靠前的）：
        ///   0 既接在上一个命中之后、又落在词边界
        ///   1 落在词边界（首字母缩略型：ds -> dataSheet）
        ///   2 紧接上一个命中（尽量连成块）
        ///   3 其它
        /// 反向 maxpos 保证"只要 query 真的是子序列就一定能凑出来"。
        /// </summary>
        public static int[] GreedyPositions(string low, string query, HashSet<int> bounds)
        {
            if (string.IsNullOrEmpty(low) || string.IsNullOrEmpty(query)) return null;
            int n = low.Length, k = query.Length;
            if (k > n) return null;

            // 反向可行性上界：maxpos[t] = query[t] 最靠右的合法位置
            var maxpos = new int[k];
            int limit = n - 1;
            for (int t = k - 1; t >= 0; t--)
            {
                if (limit < 0) return null;
                int j = low.LastIndexOf(query[t], limit);
                if (j < 0) return null;
                maxpos[t] = j;
                limit = j - 1;
            }

            var pos = new int[k];
            int i = 0;
            for (int t = 0; t < k; t++)
            {
                int hi = maxpos[t];
                int bestJ = -1, bestKey = int.MaxValue;
                for (int j = i; j <= hi; j++)
                {
                    if (low[j] != query[t]) continue;
                    bool isB = bounds.Contains(j);
                    int key = (j == i && isB) ? 0 : isB ? 1 : (j == i) ? 2 : 3;
                    if (bestJ < 0 || key < bestKey || (key == bestKey && j < bestJ))
                    {
                        bestKey = key; bestJ = j;
                    }
                }
                if (bestJ < 0) return null;
                pos[t] = bestJ;
                i = bestJ + 1;
            }
            return pos;
        }

        /// <summary>命中下标里最长的那段连续长度。</summary>
        public static int LongestRun(int[] positions)
        {
            if (positions == null || positions.Length == 0) return 0;
            int best = 1, cur = 1;
            for (int i = 1; i < positions.Length; i++)
            {
                cur = positions[i] == positions[i - 1] + 1 ? cur + 1 : 1;
                if (cur > best) best = cur;
            }
            return best;
        }

        /// <summary>
        /// 模糊匹配。query 为空表示"全都要"（手动 Ctrl+Space），此时 Kind=0、
        /// Positions 为空、First=0。
        ///
        /// Kind 越大越像用户要的：
        ///   4 完全相等 > 3 前缀 > 2 词首缩略 > 1 连续块 > 0 分散
        /// </summary>
        public static bool Match(string name, string query, out FuzzyHit hit)
        {
            hit = default(FuzzyHit);
            hit.Span = 0;
            if (string.IsNullOrEmpty(name)) return false;
            if (string.IsNullOrEmpty(query))
            {
                hit.Positions = Array.Empty<int>();
                hit.Kind = 0; hit.First = 0; hit.Span = 0; hit.Run = 0;
                return true;
            }

            string low = LowerKeepLen(name);
            string q = LowerKeepLen(query);
            if (q.Length > low.Length) return false;

            var bounds = WordBoundaries(name);

            bool contiguous = false;
            int[] positions;
            int at = low.IndexOf(q, StringComparison.Ordinal);
            if (at >= 0)
            {
                // 整块连续：取"最靠前、且尽量落在词边界"的那一次出现
                int best = -1, bestRank = int.MaxValue;
                int scan = at;
                while (scan >= 0)
                {
                    int rank = scan == 0 ? 0 : (bounds.Contains(scan) ? 1 : 2);
                    if (rank < bestRank || (rank == bestRank && scan < best))
                    {
                        bestRank = rank; best = scan;
                    }
                    scan = low.IndexOf(q, scan + 1, StringComparison.Ordinal);
                }
                positions = new int[q.Length];
                for (int i = 0; i < q.Length; i++) positions[i] = best + i;
                contiguous = true;
            }
            else
            {
                positions = GreedyPositions(low, q, bounds);
                if (positions == null) return false;
            }

            int kind;
            if (string.Equals(low, q, StringComparison.Ordinal)) kind = 4;
            else if (low.StartsWith(q, StringComparison.Ordinal)) kind = 3;
            else
            {
                bool allBound = true;
                foreach (int p in positions) if (!bounds.Contains(p)) { allBound = false; break; }
                kind = allBound ? 2 : (contiguous ? 1 : 0);
            }

            hit.Positions = positions;
            hit.Kind = kind;
            hit.First = positions[0];
            hit.Span = positions[positions.Length - 1] - positions[0] + 1;
            hit.Run = LongestRun(positions);
            return true;
        }
    }
}
