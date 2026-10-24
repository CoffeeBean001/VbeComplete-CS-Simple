// 代码片段库：SQLite 两列表（title 主键 + code），外加标题模糊匹配。
//
// 这是本工具自己的库：`snippets-cs.db`，只有这一个文件。
//   库里只存片段本身（title + code），字段布局由本工具自己定义。
//
// 表结构：
//     CREATE TABLE IF NOT EXISTS snippets (title TEXT PRIMARY KEY, code TEXT NOT NULL)
//   刻意不加 id：标题天然唯一、足以当主键，重命名走 UPDATE ... WHERE title=?。
//
// 本文件不碰 COM、不碰 WinForms —— 全是纯数据操作，能被 --snipetest 单独验证。

using System;
using System.Collections.Generic;
using System.IO;

namespace VbeComplete
{
    /// <summary>一条片段 + 当前搜索词在标题上的命中下标（高亮用）。</summary>
    internal sealed class SnippetHit
    {
        public string Title;
        public string Code;
        public List<int> Hits = new List<int>();
    }

    internal sealed class SnippetStore : IDisposable
    {
        private const string Schema =
            "CREATE TABLE IF NOT EXISTS snippets (title TEXT PRIMARY KEY, code TEXT NOT NULL)";

        private readonly SqliteDb _db = new SqliteDb();
        private string _lastError = "";

        public string LastError() => string.IsNullOrEmpty(_lastError) ? _db.LastError() : _lastError;

        /// <summary>
        /// 库文件路径：`SNIPPET_DB`（settings.ini / 环境变量）优先，否则程序目录下的
        /// `snippets-cs.db`。带 `-cs` 是为了不与其他工具的库文件撞名。
        /// </summary>
        public static string DbPath()
        {
            string p = Settings.Raw("SNIPPET_DB");
            if (!string.IsNullOrEmpty(p)) return p;
            return DefaultPath();
        }

        /// <summary>库里认的默认文件名（带 `-cs`，避免与其他工具的库文件撞名）。</summary>
        private static string DefaultPath() => System.IO.Path.Combine(Settings.Home, "snippets-cs.db");

        private bool Open()
        {
            if (_db.IsOpen) return true;
            _lastError = "";
            if (!_db.Open(DbPath()))
            {
                _lastError = _db.LastError();
                Log.Write("{0}", _lastError);
                return false;
            }
            if (!_db.Exec(Schema))
            {
                _lastError = "建表失败：" + _db.LastError();
                Log.Write("{0}", _lastError);
            }
            return true;
        }

        // ------------------------------------------------------------------
        // 读
        // ------------------------------------------------------------------

        /// <summary>全部片段 [(标题, 代码)]，按标题排序（不区分大小写）。</summary>
        public List<KeyValuePair<string, string>> All()
        {
            var rows = new List<KeyValuePair<string, string>>();
            if (!Open()) return rows;
            // 每次全量读都重新评估：清掉上一次的失败痕迹，本次若真出错会再写上。
            _lastError = "";
            _db.ResetError();
            const string sql = "SELECT title, code FROM snippets ORDER BY title COLLATE NOCASE";
            foreach (var r in _db.Query(sql, 0))
                rows.Add(new KeyValuePair<string, string>(
                    r.Length > 0 ? r[0] : "", r.Length > 1 ? r[1] : ""));
            if (!string.IsNullOrEmpty(_db.LastError())) _lastError = _db.LastError();
            return rows;
        }

        /// <summary>取一条的代码；不存在返回 null。</summary>
        public string Get(string title)
        {
            if (string.IsNullOrEmpty(title)) return null;
            if (!Open()) return null;
            var rows = _db.Query("SELECT code FROM snippets WHERE title=?", 1, title);
            return rows.Count == 0 ? null : (rows[0].Length > 0 ? rows[0][0] : "");
        }

        public bool Exists(string title) => Get(title) != null;

        public int Count()
        {
            if (!Open()) return 0;
            return (int)_db.Scalar("SELECT COUNT(*) FROM snippets");
        }

        // ------------------------------------------------------------------
        // 写
        // ------------------------------------------------------------------

        /// <summary>新增。标题已存在时不覆盖（返回 false，由界面提示改走"保存修改"）。</summary>
        public bool Add(string title, string code)
        {
            title = (title ?? "").Trim();
            if (title.Length == 0) return false;
            if (!Open()) return false;
            // INSERT OR IGNORE：标题重复时 sqlite3_changes 为 0 -> 返回 false
            return _db.Write("INSERT OR IGNORE INTO snippets (title, code) VALUES (?, ?)",
                title, code ?? "");
        }

        /// <summary>改代码（标题不变）。返回是否真的改到了行。</summary>
        public bool Update(string title, string code)
        {
            title = (title ?? "").Trim();
            if (title.Length == 0) return false;
            if (!Open()) return false;
            return _db.Write("UPDATE snippets SET code=? WHERE title=?", code ?? "", title);
        }

        /// <summary>改标题。新标题已被占用时不动任何数据、返回 false。</summary>
        public bool Rename(string oldTitle, string newTitle)
        {
            oldTitle = (oldTitle ?? "").Trim();
            newTitle = (newTitle ?? "").Trim();
            if (oldTitle.Length == 0 || newTitle.Length == 0) return false;
            if (oldTitle == newTitle) return true;
            if (!Open()) return false;
            if (Exists(newTitle)) return false;
            return _db.Write("UPDATE snippets SET title=? WHERE title=?", newTitle, oldTitle);
        }

        /// <summary>删除一条。返回是否真的删掉了。</summary>
        public bool Delete(string title)
        {
            title = (title ?? "").Trim();
            if (title.Length == 0) return false;
            if (!Open()) return false;
            return _db.Write("DELETE FROM snippets WHERE title=?", title);
        }

        public void Dispose() => _db.Dispose();
    }

    // ----------------------------------------------------------------------
    // 标题模糊匹配（纯函数；三档命中，强度递减）：
    //   整标题相等 > 前缀 > 任意位置连续子串 > 子序列（`gc` 命中 `GetCode`）。
    // 多个词用空格分开，每个词都要命中（AND）。
    //
    // ⚠️ 大小写归一必须保持长度：`"İ".ToLower()` 会变成 2 个字符，一旦拿它当下标
    //    去切原串就整体错位。所以逐字符转小写、长度不变才采用。
    //
    // ⚠️ 中文标题：逐字符比较天然支持，`char.IsUpper/IsLower` 对汉字都是 false，
    //    于是"词边界加 3 分"只在英文驼峰上生效 —— 正是想要的。
    // ----------------------------------------------------------------------
    internal static class SnippetMatch
    {
        private const string BoundaryChars = " _-.[](){}+*/\\,;:=<>\"'";

        private static string LowerKeepLen(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            var sb = new System.Text.StringBuilder(text.Length);
            foreach (char ch in text)
            {
                string low = ch.ToString().ToLowerInvariant();
                sb.Append(low.Length == 1 ? low[0] : ch);
            }
            return sb.ToString();
        }

        private static bool IsBoundary(string title, int idx)
        {
            if (idx <= 0) return true;
            char prev = title[idx - 1], cur = title[idx];
            if (BoundaryChars.IndexOf(prev) >= 0) return true;
            if (char.IsUpper(cur) && char.IsLower(prev)) return true;
            if (char.IsDigit(cur) && !char.IsDigit(prev)) return true;
            return false;
        }

        /// <summary>单个词在标题上的命中情况；没命中返回 null。</summary>
        private static int? TermScore(string termLow, string title, string titleLow,
            List<int> hitsOut)
        {
            int n = termLow.Length;
            if (n <= 0) return null;

            int pos = titleLow.IndexOf(termLow, StringComparison.Ordinal);
            if (pos >= 0)
            {
                for (int i = pos; i < pos + n; i++) hitsOut.Add(i);
                if (titleLow == termLow) return 100 + n * 4;
                if (pos == 0) return 80 + n * 4;
                return 60 + n * 4;
            }

            // 子序列：贪心从左往右扫
            var hits = new List<int>();
            int qi = 0;
            for (int idx = 0; idx < titleLow.Length && qi < n; idx++)
            {
                if (titleLow[idx] != termLow[qi]) continue;
                hits.Add(idx);
                qi++;
            }
            if (qi < n) return null;

            int score = 20 + n * 2;
            foreach (int h in hits) if (IsBoundary(title, h)) score += 3;
            for (int i = 0; i + 1 < hits.Count; i++) if (hits[i + 1] == hits[i] + 1) score += 1;
            hitsOut.AddRange(hits);
            return score;
        }

        /// <summary>query 能否模糊命中 title。命中返回分数，否则 null。</summary>
        public static int? Score(string query, string title, out List<int> hits)
        {
            hits = new List<int>();
            string q = (query ?? "").Trim();
            if (q.Length == 0) return 0;
            title = title ?? "";
            string titleLow = LowerKeepLen(title);
            int total = 0;
            var all = new List<int>();
            foreach (var term in q.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var got = new List<int>();
                int? s = TermScore(LowerKeepLen(term), title, titleLow, got);
                if (s == null) { hits.Clear(); return null; }
                total += s.Value;
                all.AddRange(got);
            }
            all.Sort();
            var uniq = new List<int>();
            foreach (int h in all) if (uniq.Count == 0 || uniq[uniq.Count - 1] != h) uniq.Add(h);
            hits = uniq;
            return total;
        }

        /// <summary>
        /// 按标题模糊过滤。空 query 原样返回全部（保持库里的顺序）；
        /// 否则按"分数高 -> 标题短 -> 字母序"排。
        /// </summary>
        public static List<SnippetHit> Filter(List<KeyValuePair<string, string>> rows, string query)
        {
            var output = new List<SnippetHit>();
            var src = rows ?? new List<KeyValuePair<string, string>>();
            bool empty = string.IsNullOrWhiteSpace(query);

            if (empty)
            {
                foreach (var r in src)
                    output.Add(new SnippetHit { Title = r.Key, Code = r.Value });
                return output;
            }

            var scored = new List<Tuple<int, int, string, SnippetHit>>();
            foreach (var r in src)
            {
                List<int> hits;
                int? s = Score(query, r.Key, out hits);
                if (s == null) continue;
                scored.Add(Tuple.Create(s.Value, (r.Key ?? "").Length,
                    LowerKeepLen(r.Key),
                    new SnippetHit { Title = r.Key, Code = r.Value, Hits = hits }));
            }
            scored.Sort((a, b) =>
            {
                int c = b.Item1.CompareTo(a.Item1); if (c != 0) return c;
                c = a.Item2.CompareTo(b.Item2); if (c != 0) return c;
                return string.CompareOrdinal(a.Item3, b.Item3);
            });
            foreach (var t in scored) output.Add(t.Item4);
            return output;
        }
    }
}
