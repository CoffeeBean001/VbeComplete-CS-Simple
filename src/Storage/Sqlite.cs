// 极薄的 SQLite 封装：只 P/Invoke 系统自带的 winsqlite3.dll，不引任何 NuGet。
//
// 为什么不装 Microsoft.Data.Sqlite：那会把一坨原生 sqlite 二进制塞进产物里，
//   而这个程序从第一天起的承诺就是"一个 exe 双击就能跑"。Windows 10 起
//   System32 下自带 winsqlite3.dll，导出的就是标准 SQLite C API。
//
// 这一层只管"把 SQL 跑掉、把列读回来"，不认识任何业务表。
//   snippets-cs.db 建在它上面。
//
// 写出去的必须是**真·SQLite 文件**：用户会用 DB Browser 之类去编辑词库，
//   自造的 JSON 格式在这个需求面前没有意义。

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace VbeComplete
{
    internal sealed class SqliteDb : IDisposable
    {
        private const int SQLITE_OK = 0;
        private const int SQLITE_ROW = 100;
        private const int SQLITE_DONE = 101;
        private static readonly IntPtr SQLITE_TRANSIENT = new IntPtr(-1);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl,
            CharSet = CharSet.Unicode)]
        private static extern int sqlite3_open16(string filename, out IntPtr db);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl,
            CharSet = CharSet.Unicode)]
        private static extern int sqlite3_prepare16_v2(IntPtr db, string sql, int nBytes,
            out IntPtr stmt, IntPtr pzTail);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl,
            CharSet = CharSet.Unicode)]
        private static extern int sqlite3_bind_text16(IntPtr stmt, int index, string value,
            int nBytes, IntPtr destructor);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr sqlite3_column_text16(IntPtr stmt, int col);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_column_count(IntPtr stmt);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_step(IntPtr stmt);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_finalize(IntPtr stmt);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_close(IntPtr db);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr sqlite3_errmsg16(IntPtr db);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_changes(IntPtr db);

        private IntPtr _db = IntPtr.Zero;
        private string _path;
        private string _lastError = "";

        public string LastError() => _lastError;
        public string Path => _path ?? "";

        /// <summary>数据库句柄是否已拿到。</summary>
        public bool IsOpen => _db != IntPtr.Zero;

        /// <summary>打开（不存在就建）。目录不存在会先建目录。失败返回 false，原因在 LastError。</summary>
        public bool Open(string path)
        {
            if (_db != IntPtr.Zero) return true;
            _path = path;
            try
            {
                string dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path));
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                IntPtr db;
                int rc = sqlite3_open16(path, out db);
                if (rc != SQLITE_OK || db == IntPtr.Zero)
                {
                    _lastError = "打开数据库失败 rc=" + rc + " 路径=" + path;
                    Log.Write("{0}", _lastError);
                    if (db != IntPtr.Zero) sqlite3_close(db);
                    return false;
                }
                _db = db;
                // 用户的词库 / 片段库是**允许用 DB Browser 等工具同时打开**的
                //   （文件头就写着"用户会用 DB Browser 去编辑词库"）。没有 busy
                //   超时的话，对端占着写锁时我们这一下会立刻拿到 SQLITE_BUSY，
                //   表现为"保存片段没反应"。2 秒足够对端写完一行。
                Exec("PRAGMA busy_timeout = 2000");
                ResetError();
                return true;
            }
            catch (Exception e)
            {
                // DllNotFound（老系统没有 winsqlite3.dll）也走这里，绝不抛给用户
                _lastError = "数据库不可用：" + e.Message;
                Log.Write("{0}", _lastError);
                return false;
            }
        }

        private string ErrMsg()
        {
            if (_db == IntPtr.Zero) return _lastError;
            try
            {
                IntPtr p = sqlite3_errmsg16(_db);
                return p == IntPtr.Zero ? _lastError : Marshal.PtrToStringUni(p);
            }
            catch { return _lastError; }
        }

        /// <summary>执行一条不带参数、也不读结果的 SQL（建表 / 建索引 / 触发器）。</summary>
        public bool Exec(string sql)
        {
            if (_db == IntPtr.Zero) return false;
            IntPtr stmt;
            try
            {
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero)
                {
                    _lastError = "SQL 准备失败：" + ErrMsg();
                    return false;
                }
                try
                {
                    rc = sqlite3_step(stmt);
                    bool ok = rc == SQLITE_DONE || rc == SQLITE_ROW;
                    if (!ok) _lastError = "SQL 执行失败：" + ErrMsg();
                    return ok;
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e) { _lastError = e.Message; return false; }
        }

        /// <summary>执行一条带 N 个文本参数的写语句，返回是否真的改到了行。</summary>
        public bool Write(string sql, params object[] args)
        {
            if (_db == IntPtr.Zero) return false;
            IntPtr stmt;
            try
            {
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero)
                {
                    _lastError = "SQL 准备失败：" + ErrMsg();
                    return false;
                }
                try
                {
                    for (int i = 0; i < args.Length; i++) BindOne(stmt, i + 1, args[i]);
                    rc = sqlite3_step(stmt);
                    if (rc != SQLITE_DONE && rc != SQLITE_ROW)
                    {
                        _lastError = "SQL 执行失败：" + ErrMsg();
                        return false;
                    }
                    return sqlite3_changes(_db) > 0;
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e) { _lastError = e.Message; return false; }
        }

        /// <summary>
        /// 参数只支持 string / int / bool / null。
        /// int 走 bind_text16（转成十进制文本）是**故意**的：SQLite 是动态类型，
        ///   表里那些 0/1 列本来就被当数字存，用文本绑进去写入时会被列的亲和性
        ///   （INTEGER）自动转回去。少一套 bind_int 就少一处原生签名写错的机会。
        /// null 绑成空串（本项目的表里没有可空列，用不上真正的 SQL NULL）。
        /// </summary>
        private void BindOne(IntPtr stmt, int index, object v)
        {
            if (v == null) { sqlite3_bind_text16(stmt, index, "", 0, SQLITE_TRANSIENT); return; }
            string s;
            if (v is bool) s = (bool)v ? "1" : "0";
            else if (v is int) s = ((int)v).ToString();
            else s = Convert.ToString(v) ?? "";
            sqlite3_bind_text16(stmt, index, s, s.Length * 2, SQLITE_TRANSIENT);
        }

        /// <summary>
        /// 跑一条查询，把每一行读成 string[]。最多 maxRows 行（0 = 不限）。
        /// SQL 里的 ? 按 args 顺序绑定。
        /// </summary>
        public List<string[]> Query(string sql, int maxRows, params object[] args)
        {
            var rows = new List<string[]>();
            if (_db == IntPtr.Zero) return rows;
            IntPtr stmt;
            try
            {
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero)
                {
                    _lastError = "SQL 准备失败：" + ErrMsg();
                    return rows;
                }
                try
                {
                    for (int i = 0; i < args.Length; i++) BindOne(stmt, i + 1, args[i]);
                    int ncol = ColumnCount(stmt);
                    // 不能写成 `while (sqlite3_step(stmt) == SQLITE_ROW)`，
                    //   于是 SQLITE_BUSY / SQLITE_ERROR / SQLITE_CORRUPT 全被当成
                    //   "没有更多行了"，_lastError 保持空 —— 读操作静默返回空表
                    //   或半张表，上层（片段库 / 词库）还以为一切正常。
                    //   典型场景：用户开着 DB Browser 编辑词库，它持有写锁。
                    int rc2 = SQLITE_DONE;
                    while ((rc2 = sqlite3_step(stmt)) == SQLITE_ROW)
                    {
                        var row = new string[ncol];
                        for (int c = 0; c < ncol; c++) row[c] = ColText(stmt, c);
                        rows.Add(row);
                        if (maxRows > 0 && rows.Count >= maxRows) break;
                    }
                    if (rc2 != SQLITE_ROW && rc2 != SQLITE_DONE)
                        _lastError = "读失败（rc=" + rc2 + "）：" + ErrMsg();
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e) { _lastError = e.Message; }
            return rows;
        }

        /// <summary>单值查询（COUNT(*) 之类）；出错或没行返回 def。</summary>
        public long Scalar(string sql, long def = 0, params object[] args)
        {
            var rows = Query(sql, 1, args);
            if (rows.Count == 0 || rows[0].Length == 0) return def;
            long n;
            return long.TryParse(rows[0][0], out n) ? n : def;
        }

        public void ResetError() => _lastError = "";

        private static int ColumnCount(IntPtr stmt)
        {
            try { return sqlite3_column_count(stmt); } catch { return 0; }
        }

        private static string ColText(IntPtr stmt, int col)
        {
            IntPtr p = sqlite3_column_text16(stmt, col);
            return p == IntPtr.Zero ? "" : (Marshal.PtrToStringUni(p) ?? "");
        }

        public void Dispose()
        {
            if (_db != IntPtr.Zero)
            {
                try { sqlite3_close(_db); } catch { }
                _db = IntPtr.Zero;
            }
        }
    }
}
