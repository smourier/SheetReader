using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using SheetReader.Wpf.Test.Utilities;

namespace SheetReader.Wpf.Test
{
    public class Settings : Serializable<Settings>
    {
        public const string FileName = "settings.json";

        public static Settings Current { get; }
        public static string ConfigurationFilePath { get; }

        static Settings()
        {
            // data is stored in user's Documents
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), typeof(Settings).Namespace!);

            ConfigurationFilePath = Path.Combine(path, FileName);

            // build settings
            Current = Deserialize(ConfigurationFilePath)!;
        }

        public virtual void SerializeToConfiguration() => Serialize(ConfigurationFilePath);
        public static void BackupConfiguration() => Backup(ConfigurationFilePath);

        [DefaultValue(null)]
        [Browsable(false)]
        public virtual IList<RecentFile>? RecentFilesPaths { get => GetPropertyValue((IList<RecentFile>?)null); set => SetPropertyValue(value); }

        private Dictionary<string, RecentFile> GetRecentFiles()
        {
            var dic = new Dictionary<string, RecentFile>(StringComparer.Ordinal);
            var recents = RecentFilesPaths;
            if (recents != null)
            {
                foreach (var recent in recents)
                {
                    if (recent?.FilePath == null)
                        continue;

                    if (!IsValidPath(recent.FilePath))
                        continue;

                    dic[recent.FilePath] = new RecentFile() { FilePath = recent.FilePath, LastAccessTime = recent.LastAccessTime, LoadOptions = recent.LoadOptions };
                }
            }
            return dic;
        }

        private void SaveRecentFiles(Dictionary<string, RecentFile> dic)
        {
            var list = dic.Select(kv => new RecentFile { FilePath = kv.Key, LastAccessTime = kv.Value.LastAccessTime, LoadOptions = kv.Value.LoadOptions }).OrderByDescending(r => r.LastAccessTime).ToList();
            if (list.Count == 0)
            {
                RecentFilesPaths = null;
            }
            else
            {
                RecentFilesPaths = list;
            }
            SerializeToConfiguration();
        }

        private static bool IsValidPath(string filePath)
        {
            if (!Uri.TryCreate(filePath, UriKind.Absolute, out var uri))
                return IOUtilities.PathIsFile(filePath);

            if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                return true;

            return IOUtilities.PathIsFile(filePath);
        }

        public void CleanRecentFiles() => SaveRecentFiles(GetRecentFiles());
        public void ClearRecentFiles()
        {
            RecentFilesPaths = null;
            SerializeToConfiguration();
        }

        public void AddRecentFile(string filePath, LoadOptions options)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            if (!IsValidPath(filePath))
                return;

            var dic = GetRecentFiles();
            dic[filePath] = new RecentFile() { FilePath = filePath, LastAccessTime = DateTime.UtcNow, LoadOptions = options };
            SaveRecentFiles(dic);
        }
    }
}
