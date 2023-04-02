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
using System.Diagnostics;
using XstReader.App.Common;

namespace XstReader.App.Controls
{
    public partial class XstMessageContentViewControl : UserControl,
                                                        IXstDataSourcedControl<XstMessage>,
                                                        IXstElementDoubleClickable<XstElement>
    {
        private XstRecipientListControl RecipientListControl { get; } = new XstRecipientListControl() { Name = "Recipients List" };
        private XstAttachmentListControl AttachmentListControl { get; } = new XstAttachmentListControl() { Name = "Attachments List" };


        public XstMessageContentViewControl()
        {
            InitializeComponent();
            Initialize();
        }
        private void Initialize()
        {
            if (DesignMode) return;

            RecipientListControl.SelectedItemChanged += (s, e) => RaiseSelectedItemChanged(e.Element);
            RecipientListControl.GotFocus += (s, e) => RaiseSelectedItemChanged();
            RecipientListControl.DoubleClickItem += (s, e) => RaiseDoubleClickItem(e.Element);

            AttachmentListControl.SelectedItemChanged += (s, e) => RaiseSelectedItemChanged(e.Element);
            AttachmentListControl.GotFocus += (s, e) => RaiseSelectedItemChanged();
            AttachmentListControl.DoubleClickItem += (s, e) => RaiseDoubleClickItem(e.Element);

            ExportToolStripButton.Enabled = false;
            ExportToolStripButton.Click += (s, e) => Helpers.ExportHelper.ExportMessages(_DataSource);
            PrintToolStripButton.Enabled = false;
            PrintToolStripButton.Click += (s, e) => Print();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            OnLoad();
        }

        private void OnLoad()
        {
            KryptonDockingManager.ManageControl(MainKryptonPanel);
            KryptonDockingManager.AddXstDockSpaceInStack(DockingEdge.Top, RecipientListControl, AttachmentListControl);
        }

        public event EventHandler<XstElementEventArgs>? SelectedItemChanged;

        private void RaiseSelectedItemChanged() => RaiseSelectedItemChanged(GetSelectedItem());
        private void RaiseSelectedItemChanged(XstElement? element) => SelectedItemChanged?.Invoke(this, new XstElementEventArgs(element));

        public event EventHandler<XstElementEventArgs>? DoubleClickItem;
        private void RaiseDoubleClickItem(XstElement? element) => DoubleClickItem?.Invoke(this, new XstElementEventArgs(element));

        private XstMessage? _DataSource;
        public XstMessage? GetDataSource()
            => _DataSource;

        public void SetDataSource(XstMessage? dataSource)
        {
            CleanTempFile();
            _DataSource = dataSource;

            ExportToolStripButton.Enabled = _DataSource != null;
            PrintToolStripButton.Enabled = _DataSource != null;

            RecipientListControl.SetDataSource(dataSource?.Recipients.Items);
            AttachmentListControl.SetDataSource(dataSource?.Attachments);
            AttachmentListControl.SetError(dataSource?.HasErrorInAttachments ?? false, dataSource?.ErrorInAttachments ?? "");

            //var htmlText = _DataSource?.GetHtmlVisualization() ?? "";
            var htmlText = _DataSource.RenderAsHtml();
            try
            {
                if (htmlText.Length < 1500000)
                    KryptonWebBrowser.DocumentText = htmlText;
                else
                    KryptonWebBrowser.Url = new Uri(SetTempFileContent(htmlText));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error showing message");
            }

            TransportHeadersTextBox.Text = _DataSource?.Properties[ElementProperties.PropertyCanonicalName.PidTagTransportMessageHeaders]?.Value ?? string.Empty;
        }

        public void Print()
        {
            if (_DataSource == null)
                return;
            KryptonWebBrowser.ShowPrintPreviewDialog();
        }

        public XstElement? GetSelectedItem()
        {
            if (RecipientListControl.Focused)
                return RecipientListControl.GetSelectedItem();
            if (AttachmentListControl.Focused)
                return AttachmentListControl.GetSelectedItem();

            return _DataSource;
        }

        public void SetSelectedItem(XstElement? item)
        { }

        private string? _TempFileName = null;
        private void CleanTempFile()
        {
            if (_TempFileName != null && File.Exists(_TempFileName))
                try { File.Delete(_TempFileName); }
                catch { }
            _TempFileName = null;
        }
        private string SetTempFileContent(string text)
        {
            _TempFileName = Path.GetTempFileName() + ".html";
            File.WriteAllText(_TempFileName, text);

            return _TempFileName;
        }
        public void ClearContents()
        {
            RecipientListControl.ClearContents();
            AttachmentListControl.ClearContents();

            GetDataSource()?.ClearContents();
            SetDataSource(null);
        }
    }
}
