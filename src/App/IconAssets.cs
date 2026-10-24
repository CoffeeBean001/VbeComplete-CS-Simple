// 程序图标：从 exe 内嵌的 app.ico 里取，按需要的尺寸挑最合适的那一帧。
//
// 为什么内嵌而不是读外部文件：
//   1) 单文件发布时 exe 旁边没有 app.ico，读文件会变成"托盘一片空白"；
//   2) 托盘图标、任务栏图标、窗口图标必须是同一张图，用户才认得出；
//   3) .ico 里已经放了 16/20/24/32/40/48/64/128/256 九帧，`new Icon(stream, size)`
//      会自动挑最接近的一帧，不靠系统缩放（缩放出来的 16px 会糊）。
//
// `app.ico` 由 `tools/make-icon.py` 生成（深蓝圆角方块 + 白色 V），
//   改图标就改那个脚本再跑一遍，别手工塞位图。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace VbeComplete
{
    internal static class IconAssets
    {
        private const string ResourceName = "VbeComplete.app.ico";

        // 按尺寸缓存，且**永不替换已交出去的实例**：托盘 / 窗口 / 片段窗都直接
        //   拿这里的引用（它们不 Dispose 图标），替换或释放旧值会让正在显示的图标
        //   变成"已释放对象"—— 真机上表现为托盘一片空白或 NotifyIcon 抛异常。
        //   尺寸最多九种，一个进程里最多九个小对象，不值得为它冒风险。
        private static readonly Dictionary<int, Icon> _cache = new Dictionary<int, Icon>();
        private static readonly object _lock = new object();

        /// <summary>
        /// 取指定像素尺寸的程序图标。**返回的是共享实例，调用方不得 Dispose**
        /// （它归本类所有，进程活多久就活多久）；失败时退回系统默认图标（绝不抛）。
        /// </summary>
        public static Icon Get(int size)
        {
            if (size <= 0) size = 32;
            lock (_lock)
            {
                Icon hit;
                if (_cache.TryGetValue(size, out hit)) return hit;

                try
                {
                    var asm = typeof(IconAssets).Assembly;
                    using (Stream s = asm.GetManifestResourceStream(ResourceName))
                    {
                        if (s != null)
                        {
                            var ic = new Icon(s, new Size(size, size));
                            _cache[size] = ic;
                            return ic;
                        }
                    }
                    Log.Write("图标资源没找到（{0}），退回系统图标", ResourceName);
                }
                catch (Exception e) { Log.Write("图标加载失败：{0}", e.Message); }

                // 必须 Clone：SystemIcons.Application 是**进程共享**对象，交出去
                //   被别人 Dispose 掉，本进程里其它取用处就都拿到已释放图标。
                var fallback = (Icon)SystemIcons.Application.Clone();
                _cache[size] = fallback;
                return fallback;
            }
        }

        /// <summary>托盘用的那一份（32px 足够，高 DPI 下系统会自己挑大图标）。</summary>
        public static Icon Tray() => Get(32);

        /// <summary>窗口标题栏 / 任务栏用（32px，系统按需要缩到标题栏的高度）。</summary>
        public static Icon Window() => Get(32);
    }
}
