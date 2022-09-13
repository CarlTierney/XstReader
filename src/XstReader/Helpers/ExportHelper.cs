﻿// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using RtfPipe.Tokens;
using XstReader.Exporter;

namespace XstReader.App.Helpers
{
    public static class ExportHelper
    {
        private static SaveFileDialog SaveFileDialog = new();
        private static FolderBrowserDialog FolderBrowserDialog = new() { ShowNewFolderButton = true };

        public static bool AskDirectoryPath(ref string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
                FolderBrowserDialog.SelectedPath = path;

            if (FolderBrowserDialog.ShowDialog() != DialogResult.OK)
                return false;
            path = FolderBrowserDialog.SelectedPath;
            return true;
        }

        public static bool AskFileName(ref string fileName)
        {
            if (!string.IsNullOrWhiteSpace(fileName))
                SaveFileDialog.FileName = fileName;

            if (SaveFileDialog.ShowDialog() != DialogResult.OK)
                return false;
            fileName = SaveFileDialog.FileName;
            return true;
        }

        public static bool ExportAttachmentsToDirectory(XstFolder? folder, string path, bool includeSubfolders, XstAttachmentExportOptions options)
        {
            if (folder == null)
                return false;

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            bool ret = true;
            foreach (var message in folder.Messages.OrderBy(m => m.LastModificationTime))
            {
                if (!ExportAttachmentsToDirectory(message, path, options))
                    ret = false;
            }

            if (includeSubfolders)
            {
                foreach (var subFolder in folder.Folders)
                {
                    var subfolderPathBase = Path.Combine(path, subFolder.DisplayName.ReplaceInvalidFileNameChars("_"));
                    var subfolderPath = subfolderPathBase;
                    int i = 1;
                    while (Directory.Exists(subfolderPath))
                        subfolderPath = $"{subfolderPathBase}({i++})";
                    if (!ExportAttachmentsToDirectory(subFolder, subfolderPath, includeSubfolders, options))
                        ret = false;
                }
            }
            return ret;
        }

        public static bool ExportAttachmentsToDirectory(XstMessage? message, string path, XstAttachmentExportOptions options)
        {
            if (message == null)
                return false;

            //if (!Directory.Exists(path))
            //    Directory.CreateDirectory(path);

            bool ret = true;

            var subfolderPath = path;
            if (options.EachMessageInFolder)
            {
                var subfolderPathBase = Path.Combine(path, message.DisplayName.ReplaceInvalidFileNameChars("_"));
                subfolderPath = subfolderPathBase;
                int i = 1;
                while (Directory.Exists(subfolderPath))
                    subfolderPath = $"{subfolderPathBase}({i++})";
            }
            if (!ExportAttachmentsToDirectory(message.Attachments, subfolderPath, options))
                ret = false;

            return ret;
        }
        public static bool ExportAttachmentsToDirectory(IEnumerable<XstAttachment>? attachments, string path, XstAttachmentExportOptions options)
        {
            var attachmentsToExport = attachments?.Where(a => a.IsFile && (!a.IsHidden || options.IncludeHidden)).OrderBy(a => a.LastModificationTime);
            if (!(attachmentsToExport?.Any() ?? false))
                return false;
            bool ret = true;

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            foreach (var attachment in attachmentsToExport)
            {
                string extension = Path.GetExtension(attachment.FileNameForSaving);
                string fileNameBase = Path.Combine(path, Path.GetFileNameWithoutExtension(attachment.FileNameForSaving));
                string fileName = fileNameBase + extension;
                int i = 1;
                while (File.Exists(fileName))
                    fileName = $"{fileNameBase}({i++}){extension}";
                try { attachment.SaveToFile(fileName); }
                catch { ret = false; }
            }
            return ret;
        }


        public static bool ExportFolderToHtmlFiles(XstFolder? folder, bool includeSubfolders)
        {
            if (folder == null)
                return false;

            if (FolderBrowserDialog.ShowDialog() != DialogResult.OK)
                return false;

            return ExportFolderToHtmlFiles(folder, FolderBrowserDialog.SelectedPath, includeSubfolders);
        }
        public static bool ExportFolderToHtmlFiles(XstFolder? folder, string path, bool includeSubfolders)
        {
            if (folder == null)
                return false;

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            bool ret = true;
            foreach (var message in folder.Messages.OrderBy(m => m.Date ?? DateTime.MinValue))
            {
                string fileNameBase = Path.Combine(path, message.GetFilenameForExport());
                string fileName = fileNameBase + ".html";
                int i = 1;
                while (File.Exists(fileName))
                    fileName = $"{fileNameBase}({i++}).html";
                if (!ExportMessageToHtmlFile(message, fileName))
                    ret = false;
            }
            if (includeSubfolders)
            {
                foreach (var subFolder in folder.Folders)
                {
                    var subfolderPathBase = Path.Combine(path, subFolder.DisplayName.ReplaceInvalidFileNameChars("_"));
                    var subfolderPath = subfolderPathBase;
                    int i = 1;
                    while (Directory.Exists(subfolderPath))
                        subfolderPath = $"{subfolderPathBase}({i++})";
                    if (!ExportFolderToHtmlFiles(subFolder, subfolderPath, includeSubfolders))
                        ret = false;
                }
            }
            return ret;
        }

        public static bool ExportMessageToHtmlFile(XstMessage? message, bool openFile)
        {
            if (message == null)
                return false;

            SaveFileDialog.FileName = message.GetFilenameForExport() + ".html";


            if (SaveFileDialog.ShowDialog() == DialogResult.OK &&
                ExportMessageToHtmlFile(message, SaveFileDialog.FileName))
            {
                if (openFile)
                    SystemHelper.OpenWith(SaveFileDialog.FileName);
                return true;
            }
            return false;
        }

        public static bool ExportMessageToHtmlFile(XstMessage? message, string fileName)
        {
            if (message == null)
                return false;

            try
            {
                File.WriteAllText(fileName, message.RenderAsHtml(false));
                if (message.Date.HasValue)
                    File.SetCreationTime(fileName, message.Date.Value);
            }
            catch { return false; }
            return true;
        }
    }
}
