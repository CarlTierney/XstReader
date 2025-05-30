﻿// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using Krypton.Docking;
using Krypton.Toolkit;
using System.Data;
using XstReader.App.Controls;
using XstReader.App.Helpers;

namespace XstReader.App
{
    public partial class MainForm : KryptonForm
    {
        private XstFolderTreeControl FolderTreeControl { get; } = new XstFolderTreeControl() { Name = "Folders Tree" };
        private XstMessageListControl MessageListControl { get; } = new XstMessageListControl() { Name = "Message List" };
        private XstMessageViewControl MessageViewControl { get; } = new XstMessageViewControl() { Name = "Message View" };
        private XstPropertiesControl PropertiesControl { get; } = new XstPropertiesControl() { Name = "Properties" };
        private XstPropertiesInfoControl InfoControl { get; } = new XstPropertiesInfoControl() { Name = "Information" };


        private XstFile? _XstFile = null;
        private XstFile? XstFile
        {
            get => _XstFile;
            set => SetXstFile(value);
        }
        private void SetXstFile(XstFile? value)
        {
            _XstFile = value;
            FolderTreeControl.SetDataSource(value);
            CloseFileToolStripMenuItem.Enabled = value != null;
        }

        private XstElement? _CurrentXstElement = null;
        private XstElement? CurrentXstElement
        {
            get => _CurrentXstElement;
            set => SetCurrentXstElement(value);
        }
        private void SetCurrentXstElement(XstElement? value)
        {
            if (_CurrentXstElement == value)
                return;

            
            if (value is XstFolder folder)
            {
                MessageListControl.SetDataSource(folder?.Messages?.OrderByDescending(m => m.Date));//, MessageFilter.GetSelectedFilter());
                MessageListControl.SetError(folder?.HasErrorInMessages ?? false, folder?.ErrorInMessages ?? "");
            }
            else if (value is XstMessage message)
            {
                MessageViewControl.SetDataSource(message);
            }

            MessageToolStripMenuItem.Enabled = value != null && value is XstMessage;
            InfoControl.SetDataSource(value);
            PropertiesControl.SetDataSource(value);

            UpdateMenu();

            _CurrentXstElement = value;
        }

        private XstFile? GetCurrentXstFile() => FolderTreeControl.GetDataSource();
        private XstFolder? GetCurrentXstFolder() => FolderTreeControl.GetSelectedItem();
        private XstMessage? GetCurrentXstMessage() => MessageViewControl.GetDataSource();

        private void UpdateMenu()
        {
            FileExportFoldersToolStripMenuItem.Enabled =
                FileExportAttachmentsToolStripMenuItem.Enabled =
                GetCurrentXstFile() != null;

            FolderToolStripMenuItem.Enabled = GetCurrentXstFolder() != null;
            MessageToolStripMenuItem.Enabled = GetCurrentXstMessage() != null;
            MessageExportAttachmentsToolStripMenuItem.Enabled = GetCurrentXstMessage()?.Attachments?.Any(a => a.IsFile) ?? false;
        }

        public MainForm()
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            OpenToolStripMenuItem.Click += OpenXstFile;
            CloseFileToolStripMenuItem.Click += (s, e) => CloseXstFile();

            FolderTreeControl.SelectedItemChanged += (s, e) => CurrentXstElement = e.Element;
            FolderTreeControl.GotFocus += (s, e) => CurrentXstElement = FolderTreeControl.GetSelectedItem();

            MessageListControl.SelectedItemChanged += (s, e) => CurrentXstElement = e.Element;
            MessageListControl.GotFocus += (s, e) => CurrentXstElement = MessageListControl.GetSelectedItem();

            MessageViewControl.SelectedItemChanged += (s, e) => CurrentXstElement = e.Element;
            MessageViewControl.GotFocus += (s, e) => CurrentXstElement = MessageViewControl.GetSelectedItem();

            AboutToolStripMenuItem.Click += (s, e) => { using var f = new AboutForm(); f.ShowDialog(); };

            FileExportFoldersToolStripMenuItem.Click += (s, e) => ExportHelper.ExportMessages(GetCurrentXstFile());
            FileExportAttachmentsToolStripMenuItem.Click += (s, e) => ExportHelper.ExportAttachments(GetCurrentXstFile());

            FolderExportMessagesToolStripMenuItem.Click += (s, e) => ExportHelper.ExportMessages(GetCurrentXstFolder());
            FolderExportAttachmentsToolStripMenuItem.Click += (s, e) => ExportHelper.ExportAttachments(GetCurrentXstFolder());

            MessagePrintToolStripMenuItem.Click += (s, e) => MessageViewControl.Print();
            MessageExportToolStripMenuItem.Click += (s, e) => ExportHelper.ExportMessages(GetCurrentXstMessage());
            MessageExportAttachmentsToolStripMenuItem.Click += (s, e) => ExportHelper.ExportAttachments(GetCurrentXstMessage());

            SettingsToolStripMenuItem.Click += (s, e) => new ExportOptionsForm().ShowDialog();

            LayoutDefaultToolStripMenuItem.Click += (s, e) =>
            {
                try
                {
                    var fileName = Path.GetTempFileName() + ".xml";
                    File.WriteAllBytes(fileName, Properties.Resources.layout_default);
                    KryptonDockingManager.LoadConfigFromFile(fileName);
                    File.Delete(fileName);
                }
                catch { }
            };
            LayoutClassic3PanelToolStripMenuItem.Click += (s, e) =>
            {
                try
                {
                    var fileName = Path.GetTempFileName() + ".xml";
                    File.WriteAllBytes(fileName, Properties.Resources.layout_3panels);
                    KryptonDockingManager.LoadConfigFromFile(fileName);
                    File.Delete(fileName);
                }
                catch { }
            };

            Reset();
            UpdateMenu();
        }


        private void OpenXstFile(object? sender, EventArgs e)
        {
            if (OpenXstFileDialog.ShowDialog(this) == DialogResult.OK)
                LoadXstFile(OpenXstFileDialog.FileName);
        }

        private void Reset()
        {
            if (XstFile != null)
            {
                
                XstFile.Dispose();
                XstFile = null;

                FolderTreeControl.ClearContents();
                MessageListControl.ClearContents();
                MessageViewControl.ClearContents();
                //RecipientListControl.ClearContents();
                //AttachmentListControl.ClearContents();
                MessageViewControl.ClearContents();
               
            }
        }
        private void LoadXstFile(string filename)
        {
            CloseXstFile();
            XstFile = new XstFile(filename);
        }

        private void CloseXstFile()
        {
            Reset();
        }

        protected override void OnClosed(EventArgs e)
        {
            try { KryptonDockingManager.SaveConfigToFile(Path.Combine(Application.StartupPath, "Layout.xml")); }
            catch { }

            try { XstReaderOptions.SaveToFile(Path.Combine(Application.StartupPath, "Options.xml"), XstReaderEnvironment.Options); }
            catch { }

            Reset();
            base.OnClosed(e);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            LoadForm();
        }

        private void LoadForm()
        {
            KryptonMessagePanel.BeginInit();
            KryptonMessagePanel.Controls.Add(MessageViewControl);
            MessageViewControl.Dock = DockStyle.Fill;

            // Setup docking functionality
            KryptonDockingManager.ManageControl(KryptonMainPanel);

            KryptonDockingManager.AddXstDockSpaceInTabs(DockingEdge.Top, MessageListControl);
            KryptonDockingManager.AddXstDockSpaceInTabs(DockingEdge.Right, InfoControl, PropertiesControl);
            KryptonDockingManager.AddXstDockSpaceInTabs(DockingEdge.Left, FolderTreeControl);

            try { KryptonDockingManager.LoadConfigFromFile(Path.Combine(Application.StartupPath, "Layout.xml")); }
            catch { LayoutDefaultToolStripMenuItem.PerformClick(); }
            KryptonMessagePanel.EndInit();

            try { XstReaderEnvironment.Options = XstReaderOptions.LoadFromFile(Path.Combine(Application.StartupPath, "Options.xml")); }
            catch { }
        }
    }
}
