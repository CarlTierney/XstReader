﻿// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using XstReader.App.Common;

namespace XstReader.App.Controls
{
    public partial class XstMessageViewControl : UserControl,
                                                 IXstDataSourcedControl<XstMessage>,
                                                 IXstElementSelectable<XstElement>
    {

        private XstMessageContentViewControl MessageContentControl { get; } = new XstMessageContentViewControl();

        public XstMessageViewControl()
        {
            InitializeComponent();
            Initialize();
        }
        private void Initialize()
        {
            if (DesignMode) return;

            MessageContentControl.DoubleClickItem += (s, e) => AddTab(e?.Element);
            MessageContentControl.SelectedItemChanged += (s, e) => RaiseSelectedItemChanged(e.Element);

            MainKryptonNavigator.CloseAction += (s, e) =>
            {
                if (MainKryptonNavigator.SelectedIndex == 0)
                    e.Action = Krypton.Navigator.CloseButtonAction.None;
            };
            MainKryptonNavigator.SelectedPageChanged += (s, e) =>
            {
                if (MainKryptonNavigator.SelectedPage?.Tag is XstElement element)
                    RaiseSelectedItemChanged(element);
            };
        }

        public event EventHandler<XstElementEventArgs>? SelectedItemChanged;
        private void RaiseSelectedItemChanged(XstElement? element) => SelectedItemChanged?.Invoke(this, new XstElementEventArgs(element));

        private XstMessage? _DataSource;
        public XstMessage? GetDataSource()
            => _DataSource;

        public void SetDataSource(XstMessage? dataSource)
        {
            if (dataSource != null && _DataSource != null && dataSource.GetHashCode() == _DataSource.GetHashCode())
                return;

            _DataSource = dataSource;

            if (dataSource != null)
            {
                if (MainKryptonNavigator.Pages.Count == 0)
                {
                    var page = new Krypton.Navigator.KryptonPage();
                    MainKryptonNavigator.Pages.Add(page);
                    page.Controls.Add(MessageContentControl);
                    MessageContentControl.Dock = DockStyle.Fill;
                }
                MainKryptonNavigator.Pages[0].Text = $"Message: {dataSource.Subject}";
                MainKryptonNavigator.Pages[0].Tag = dataSource;
                MainKryptonNavigator.SelectedIndex = 0;
                while (MainKryptonNavigator.Pages.Count > 1)
                    MainKryptonNavigator.Pages.RemoveAt(1);
            }
            MessageContentControl.SetDataSource(dataSource);
        }

        public XstElement? GetSelectedItem()
        {
            return _DataSource;
        }

        public void SetSelectedItem(XstElement? item)
        { }

        public void ClearContents()
        {
            MainKryptonNavigator.Pages.Clear();
            MessageContentControl.ClearContents();

            GetDataSource();
            SetDataSource(null);
        }

        private void AddTab(XstElement? element)
        {
            if (element == null)
                return;
            if (element is XstAttachment attach && !attach.CanBeOpenedInApp())
                return;


            var page = MainKryptonNavigator.Pages.FirstOrDefault(p => p.Tag == element);
            if (page == null)
            {
                page = new Krypton.Navigator.KryptonPage
                {
                    Text = $"{element.ElementType}: {element.DisplayName}",
                    Tag = element,
                };
                MainKryptonNavigator.Pages.Add(page);

                if (element is XstAttachment attachment)
                {
                    if (attachment.IsEmail)
                    {
                        var viewer = new XstMessageContentViewControl();
                        page.Controls.Add(viewer);
                        viewer.Dock = DockStyle.Fill;
                        viewer.SetDataSource(attachment.AttachedEmailMessage);
                        viewer.SelectedItemChanged += (s, e) => RaiseSelectedItemChanged(e.Element);
                        viewer.DoubleClickItem += (s, e) => AddTab(e.Element);
                    }
                    else
                    {
                        var viewer = new XstAttachmentViewControl();
                        page.Controls.Add(viewer);
                        viewer.Dock = DockStyle.Fill;
                        viewer.SetDataSource(attachment);
                        viewer.SelectedItemChanged += (s, e) => RaiseSelectedItemChanged(e.Element);
                    }
                }
            }
            MainKryptonNavigator.SelectedPage = page;
        }

        public void Print()
            => MessageContentControl.Print();
    }
}
