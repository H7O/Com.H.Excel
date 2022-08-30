using System;
using System.IO;

namespace Com.H.Excel
{
    internal static class IOExtensions
    {
        internal static string EnsureParentDirectory(this string path)
        {
            if (string.IsNullOrEmpty(path)) throw new ArgumentNullException(nameof(path));
            if (path.IndexOfAny(Path.GetInvalidPathChars()) != -1)
                throw new ArgumentException($"{nameof(path)} contains invalid characters.");
            var parentFolder = Directory.GetParent(path)?.FullName;
            if (parentFolder == null) throw new ArgumentException($"Can't find parent folder of '{path}'");
            if (Directory.Exists(parentFolder))
                return path;
            Directory.CreateDirectory(parentFolder);
            return path;

        }
    }
}
