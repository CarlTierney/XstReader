// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.

using SearchTextBox;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;

namespace XstReader
{
    /// <summary>
    /// XstReader is a viewer for xst (.ost and .pst) files
    /// This file contains the interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private View view = new View();
        private XstFile xstFile = null;
        private List<string> tempFileNames = new List<string>();
        private int searchIndex = -1;

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = view;

            // For testing purposes, use these flags to control the display of print headers
            //view.DisplayPrintHeaders = true;
            //view.DisplayEmailType = true;

            // Supply the Search control with the list of sections
            searchTextBox.SectionsList = new List<string> { "Subject", "From/To", "Date", "Cc", "Bcc" };
            searchTextBox.SectionsInitiallySelected = new List<bool> { true, true, true, false, false };

            if (Properties.Settings.Default.Top != 0.0)
            {
                this.Top = Properties.Settings.Default.Top;
                this.Left = Properties.Settings.Default.Left;
                this.Height = Properties.Settings.Default.Height;
                this.Width = Properties.Settings.Default.Width;
            }
        }

        public void OpenFile(string fileName)
        {
            if (!System.IO.File.Exists(fileName))
                return;

            view.Clear();
            ShowStatus("Loading...");
            Mouse.OverrideCursor = Cursors.Wait;

            // Load on a background thread so we can keep the UI in sync
            Task.Factory.StartNew((Action)(() =>
            {
                try
                {
                    xstFile = new XstFile(fileName);
                    var root = xstFile.RootFolder;

                    // We may be called on a background thread, so we need to dispatch this to the UI thread
                    Application.Current.Dispatcher.Invoke(new Action(() => view.UpdateFolderViews(root)));
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error reading xst file");
                }
            }))
            // When loading completes, update the UI using the UI thread 
            .ContinueWith((task) =>
            {
                Application.Current.Dispatcher.Invoke(new Action(() =>
                {
                    ShowStatus(null);
                    Mouse.OverrideCursor = null;
                    Title = "Xst Reader - " + System.IO.Path.GetFileName(fileName);
                }));
            });
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            // Ask for a .ost or .pst file to open
            string fileName = GetXstFileName();

            if (fileName != null)
                OpenFile(fileName);
        }

        private void exportAllProperties_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ExportEmailProperties(view.SelectedFolder.MessageViews);
        }

        private void exportAllEmails_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ExportEmails(view.SelectedFolder.MessageViews);
        }

        private void treeFolders_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            try
            {
                
                FolderView fv = (FolderView)e.NewValue;
                view.SelectedFolder = fv;

                if (fv != null)
                {
                    view.SetMessage(null);
                    ShowMessage(null);
                    fv.MessageViews.Clear();
                    ShowStatus("Reading messages...");
                    Mouse.OverrideCursor = Cursors.Wait;

                    // Read messages on a background thread so we can keep the UI in sync
                    Task.Factory.StartNew(() =>
                    {
                        try
                        {
                            fv.Folder.GetMessages();
                            // We may be called on a background thread, so we need to dispatch this to the UI thread
                            Application.Current.Dispatcher.Invoke(new Action(() => fv.UpdateMessageViews()));
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.ToString(), "Error reading messages");
                        }
                    })
                    // When loading completes, update the UI using the UI thread 
                    .ContinueWith((task) =>
                    {
                        Application.Current.Dispatcher.Invoke(new Action(() =>
                        {
                            ShowStatus(null);
                            Mouse.OverrideCursor = null;
                            view.SelectedFolderChanged(fv.Folder);
                        }));
                    });

                    // If there is no sort in effect, sort by date in descending order
                    SortMessages("Date", ListSortDirection.Descending, ifNoneAlready: true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Unexpected error reading messages");
            }
        }

        private void listMessages_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            searchIndex = listMessages.SelectedIndex;
            searchTextBox.ShowSearch = true;
            MessageView mv = (MessageView)listMessages.SelectedItem;

            if (mv != null)
            {
                try
                {
                    //mv.Message.ReadMessageDetails();
                    ShowMessage(mv);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error reading message details");
                }
            }
            view.SetMessage(mv);
        }

        private void listMessagesColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            // Sort the messages by the clicked on column
            GridViewColumnHeader column = (sender as GridViewColumnHeader);
            SortMessages(column.Tag.ToString(), ListSortDirection.Ascending);

            searchIndex = listMessages.SelectedIndex;
            listMessages.ScrollIntoView(listMessages.SelectedItem);
        }

        private void listRecipients_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                view.SelectedRecipientChanged((XstRecipient)listRecipients.SelectedItem);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error reading recipient");
            }
        }

        private void listAttachments_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                view.SelectedAttachmentsChanged(listAttachments.SelectedItems.Cast<XstAttachment>());
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error reading attachment");
            }
        }

        private void exportEmail_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (listMessages.SelectedItems.Count > 1)
            {
                ExportEmails(listMessages.SelectedItems.Cast<MessageView>());
            }
            else
            {
                string fullFileName = GetEmailExportFileName(view.CurrentMessage.ExportFileName, view.CurrentMessage.MessageFormatter.ExportFileExtension);
                if (fullFileName != null)
                {
                    try
                    {
                        view.CurrentMessage.MessageFormatter.SaveMessage(fullFileName);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Error exporting email");
                    }
                }
            }
        }

        private void SortMessages(string property, ListSortDirection direction, bool ifNoneAlready = false)
        {
            var targetCol = ((GridView)listMessages.View).Columns.Select(c => (GridViewColumnHeader)c.Header)
                            .First(h => h.Tag.ToString() == property);

            // See if we have any sorts applied, in which case, we may be done
            var sorts = listMessages.Items.SortDescriptions;
            if (ifNoneAlready && sorts.Count > 0)
                return;

            // Look for a sort on the target columm
            var matches = sorts.Where(s => s.PropertyName == property);
            if (matches.Count() > 0)
            {
                // If there is one, just toggle it
                var sort = matches.First();
                direction = ((sort.Direction == ListSortDirection.Ascending) ?
                                ListSortDirection.Descending : ListSortDirection.Ascending);
                sorts.Remove(sort);
            }
            //else
            //{
            //    // If there is not one, see if we have the maximum number of sorts applied already
            //    // The algorithm works for any limit with no changes other than this test
            //    if (sorts.Count >= 2)
            //    {
            //        // If so, remove the oldest one
            //        var oldSort = sorts.Last();
            //        var oldCol = ((GridView)listMessages.View).Columns.Select(c => (GridViewColumnHeader)c.Header)
            //               .First(h => h.Tag.ToString() == oldSort.PropertyName);
            //        sorts.Remove(oldSort);

            //        // And the adorner that went with it
            //        var oldAdorners = AdornerLayer.GetAdornerLayer(oldCol);
            //        var oldAdorner = oldAdorners.GetAdorners(oldCol)?.Cast<SortAdorner>()?.FirstOrDefault(s => s != null);
            //        if (oldAdorner != null)
            //            oldAdorners.Remove(oldAdorner);
            //    }
            //}

            // Apply the requested sort as the dominant one, whatever it was before
            sorts.Insert(0, new SortDescription(property, direction));

            // Find any sort adorner applied to the target column
            var adorners = AdornerLayer.GetAdornerLayer(targetCol);
            var adorner = adorners.GetAdorners(targetCol)?.Cast<SortAdorner>()?.FirstOrDefault(s => s != null);
            // If there is one, remove it
            if (adorner != null)
                adorners.Remove(adorner);

            // Create and apply the requested adorner
            adorner = new SortAdorner(targetCol, direction);
            adorners.Add(adorner);
        }

        private void ExportEmails(IEnumerable<MessageView> messages)
        {
            string folderName = GetEmailsExportFolderName();

            if (folderName != null)
            {
                ShowStatus("Exporting emails...");
                Mouse.OverrideCursor = Cursors.Wait;

                // Export emails on a background thread so we can keep the UI in sync
                Task.Factory.StartNew<Tuple<int, int>>(() =>
                {
                    MessageView current = null;
                    int good = 0, bad = 0;
                    // If files already exist, we overwrite them.
                    // But if emails within this batch generate the same filename,
                    // use a numeric suffix to distinguish them
                    HashSet<string> usedNames = new HashSet<string>();
                    foreach (MessageView mv in messages)
                    {
                        try
                        {
                            current = mv;
                            string fileName = mv.ExportFileName;
                            for (int i = 1; ; i++)
                            {
                                if (!usedNames.Contains(fileName))
                                {
                                    usedNames.Add(fileName);
                                    break;
                                }
                                else
                                    fileName = $"{mv.ExportFileName} ({i})";
                            }
                            Application.Current.Dispatcher.Invoke(new Action(() =>
                            {
                                ShowStatus($"Exporting {mv.ExportFileName}");
                            }));
                            // Ensure that we have the message contents
                            var fullFileName = $"{Path.Combine(folderName, fileName)}.{mv.MessageFormatter.ExportFileExtension}";
                            mv.MessageFormatter.SaveMessage(fullFileName);
                            good++;
                        }
                        catch (System.Exception ex)
                        {
                            var result = MessageBox.Show($"Error '{ex.Message}' exporting email '{current.Subject}'",
                                                         "Error exporting emails", MessageBoxButton.OKCancel);
                            bad++;
                            if (result == MessageBoxResult.Cancel)
                                break;
                        }
                    }
                    return new Tuple<int, int>(good, bad);
                })
                // When exporting completes, update the UI using the UI thread 
                .ContinueWith((task) =>
                {
                    Application.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        ShowStatus(null);
                        Mouse.OverrideCursor = null;
                        txtStatus.Text = $"Completed with {task.Result.Item1} successes and {task.Result.Item2} failures";
                    }));
                });
            }
        }

        private void ExportEmailProperties(IEnumerable<MessageView> messages)
        {
            string fileName = GetPropertiesExportFileName(view.SelectedFolder.Name);

            if (fileName != null)
            {
                ShowStatus("Exporting properties...");
                Mouse.OverrideCursor = Cursors.Wait;

                // Export properties on a background thread so we can keep the UI in sync
                Task.Factory.StartNew(() =>
                {
                    try
                    {
                        messages.Select(v => v.Message).SavePropertiesToFile(fileName);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Error exporting properties");
                    }
                })
                // When exporting completes, update the UI using the UI thread 
                .ContinueWith((task) =>
                {
                    Application.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        ShowStatus(null);
                        Mouse.OverrideCursor = null;
                    }));
                });
            }
        }

        private void SaveAttachments(IEnumerable<XstAttachment> attachments)
        {
            string folderName = GetAttachmentsSaveFolderName();

            if (folderName != null)
            {
                try
                {
                    attachments.SaveToFolder(folderName, view.CurrentMessage.Date);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error '{ex.Message}' saving attachments to '{view.CurrentMessage.Subject}'", "Error saving attachments");
                }
            }
        }

        private void exportEmailProperties_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (listMessages.SelectedItems.Count > 1)
            {
                ExportEmailProperties(listMessages.SelectedItems.Cast<MessageView>());
            }
            else
            {
                string fileName = GetPropertiesExportFileName(view.CurrentMessage.ExportFileName);

                if (fileName != null)
                    view.CurrentMessage.Message.Properties.Items.NonBinary().SaveToFile(fileName);
            }
        }

        private void btnSaveAllAttachments_Click(object sender, RoutedEventArgs e)
        {
            SaveAttachments(view.CurrentMessage.Attachments);
        }

        private void btnCloseEmail_Click(object sender, RoutedEventArgs e)
        {
            view.PopMessage();
            ShowMessage(view.CurrentMessage);
        }

        private void rbContent_Click(object sender, RoutedEventArgs e)
        {
            view.ShowContent = true;
        }

        private void rbProperties_Click(object sender, RoutedEventArgs e)
        {
            view.ShowContent = false;
        }

        private void searchTextBox_OnSearch(object sender, RoutedEventArgs e)
        {
            try
            {
                var args = e as SearchEventArgs;
                bool subject = args.Sections.Contains("Subject");
                bool fromTo = args.Sections.Contains("From/To");
                bool date = args.Sections.Contains("Date");
                bool cc = args.Sections.Contains("Cc");
                bool bcc = args.Sections.Contains("Bcc");
                bool found = false;
                switch (args.SearchEventType)
                {
                    case SearchEventType.Search:
                        for (int i = 0; i < listMessages.Items.Count; i++)
                        {
                            found = PropertyHitTest(i, args.Keyword, subject, fromTo, date, cc, bcc);
                            if (found)
                                break;
                        }

                        if (!found)
                            searchIndex = -1;
                        break;
                    case SearchEventType.Next:
                        for (int i = searchIndex + 1; i < listMessages.Items.Count; i++)
                        {
                            found = PropertyHitTest(i, args.Keyword, subject, fromTo, date, cc, bcc);
                            if (found)
                            {
                                searchTextBox.ShowSearch = true;
                                break;
                            }
                        }
                        break;
                    case SearchEventType.Previous:
                        for (int i = searchIndex - 1; i >= 0; i--)
                        {
                            found = PropertyHitTest(i, args.Keyword, subject, fromTo, date, cc, bcc);
                            if (found)
                            {
                                searchTextBox.ShowSearch = true;
                                break;
                            }
                        }
                        break;
                }

                if (!found)
                    searchTextBox.IndicateSearchFailed(args.SearchEventType);
            }
            catch
            {
                // Unclear what we can do here, as we were invoked by an event from the search text box control
            }
        }

        private bool PropertyHitTest(int index, string text, bool subject, bool fromTo, bool date, bool cc, bool bcc)
        {
            MessageView mv = listMessages.Items[index] as MessageView;
            if ((subject && mv.Subject != null && mv.Subject.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0) ||
                (fromTo && mv.FromTo != null && mv.FromTo.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0) ||
                (date && mv.DisplayDate.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0) ||
                (cc && mv.CcDisplayList.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0) ||
                (bcc && mv.BccDisplayList.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0))
            {
                searchIndex = index;
                listMessages.UnselectAll();
                mv.IsSelected = true;
                listMessages.ScrollIntoView(mv);
                return true;
            }
            else
                return false;
        }

        private void ShowStatus(string status)
        {
            if (status != null)
            {
                view.IsBusy = true;
                txtStatus.Text = status;
            }
            else
            {
                view.IsBusy = false;
                txtStatus.Text = "";
            }
        }

        private void ShowMessage(MessageView mv)
        {
            try
            {
                //clear any existing status
                ShowStatus(null);

                if (mv != null)
                {
                    // Populate the view of the attachments
                    mv.SortAndSaveAttachments(mv.Message.Attachments.ToList());

                    // Can't bind HTML content, so push it into the control, if the message is HTML
                    if (mv.ShowHtml)
                    {
                        string body = mv.Message.Body.Text;

                        if (body != null)
                        {
                            body = mv.MessageFormatter.EmbedHtmlHeader(body);

                            wbMessage.NavigateToString(body);
                            if (mv.MayHaveInlineAttachment)
                            {
                                mv.SortAndSaveAttachments();  // Re-sort attachments in case any new in-line rendering discovered
                            }
                        }
                    }
                    // Can't bind RTF content, so push it into the control, if the message is RTF
                    else if (mv.ShowRtf)
                    {
                        //TODO: Rtf support
                        rtfMessage.SelectAll();
                        if (mv.Message.Body.Bytes != null)
                            using (var ms = new MemoryStream(mv.Message.Body.Bytes))
                                rtfMessage.Selection.Load(ms, DataFormats.Rtf);
                        var body = mv.MessageFormatter.GetBodyAsFlowDocument();

                        mv.MessageFormatter.EmbedRtfHeader(body);

                        rtfMessage.Document = body;
                    }
                    // Could bind text content, but use push so that we can optionally add headers
                    else if (mv.ShowText)
                    {
                        var body = mv.Body.Text;
                        body = mv.MessageFormatter.EmbedTextHeader(body);

                        txtMessage.Text = body;
                        scrollTextMessage.ScrollToHome();
                    }
                }
                else
                {
                    // Clear the displays, in case we were showing that type before
                    wbMessage.Navigate("about:blank");
                    rtfMessage.Document.Blocks.Clear();
                    txtMessage.Text = "";
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error reading message body");
            }
        }

        private void OpenEmailAttachment(XstAttachment a)
        {
            XstMessage m = a.AttachedEmailMessage;
            var mv = new MessageView(m);
            ShowMessage(mv);
            view.PushMessage(mv);
        }

        #region File and folder dialogs

        private string GetXstFileName()
        {
            // Ask for a .ost or .pst file to open
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog
            {
                Filter = "xst files (*.ost;*.pst)|*.ost;*.pst|All files (*.*)|*.*",
                FilterIndex = 1,
                InitialDirectory = Properties.Settings.Default.LastFolder
            };
            if (dialog.InitialDirectory == "")
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialog.RestoreDirectory = true;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.LastFolder = Path.GetDirectoryName(dialog.FileName);
                Properties.Settings.Default.Save();
                return dialog.FileName;
            }
            else
                return null;
        }

        private string GetAttachmentsSaveFolderName()
        {
            // Find out where to save the attachments
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Choose folder for saving attachments",
                RootFolder = Environment.SpecialFolder.MyComputer,
                SelectedPath = Properties.Settings.Default.LastAttachmentFolder
            };
            if (dialog.SelectedPath == "")
                dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialog.ShowNewFolderButton = true;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.LastAttachmentFolder = dialog.SelectedPath;
                Properties.Settings.Default.Save();
                return dialog.SelectedPath;
            }
            else
                return null;
        }

        private string GetSaveAttachmentFileName(string defaultFileName)
        {
            var dialog = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Specify file to save to",
                InitialDirectory = Properties.Settings.Default.LastAttachmentFolder
            };
            if (dialog.InitialDirectory == "")
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialog.Filter = "All Files (*.*)|*.*";
            dialog.FileName = defaultFileName;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.LastAttachmentFolder = Path.GetDirectoryName(dialog.FileName);
                Properties.Settings.Default.Save();
                return dialog.FileName;
            }
            else
                return null;
        }

        private string GetEmailsExportFolderName()
        {
            // Find out where to export the emails
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Choose folder to export emails into",
                RootFolder = Environment.SpecialFolder.MyComputer,
                SelectedPath = Properties.Settings.Default.LastExportFolder
            };
            if (dialog.SelectedPath == "")
                dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialog.ShowNewFolderButton = true;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.LastExportFolder = dialog.SelectedPath;
                Properties.Settings.Default.Save();
                return dialog.SelectedPath;
            }
            else
                return null;
        }

        private string GetEmailExportFileName(string defaultFileName, string extension)
        {
            var dialog = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Specify file to save to",
                InitialDirectory = Properties.Settings.Default.LastExportFolder
            };
            if (dialog.InitialDirectory == "")
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialog.Filter = String.Format("{0} Files (*.{0})|*.{0}|All Files (*.*)|*.*", extension);
            dialog.FileName = defaultFileName;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.LastExportFolder = Path.GetDirectoryName(dialog.FileName);
                Properties.Settings.Default.Save();
                return dialog.FileName;
            }
            else
                return null;
        }

        private string GetPropertiesExportFileName(string defaultName)
        {
            var dialog = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Specify properties export file",
                InitialDirectory = Properties.Settings.Default.LastExportFolder
            };
            if (dialog.InitialDirectory == "")
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialog.Filter = "csv files (*.csv)|*.csv";
            dialog.FileName = defaultName;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Properties.Settings.Default.LastExportFolder = Path.GetDirectoryName(dialog.FileName);
                Properties.Settings.Default.Save();
                return dialog.FileName;
            }
            else
                return null;
        }
        #endregion

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Clean up temporary files
            foreach (var fileFullName in tempFileNames)
            {
                // Wrap in try in case the file is still open
                try
                {
                    File.Delete(fileFullName);
                }
                catch { }
            }

            if (WindowState == WindowState.Maximized)
            {
                // Use the RestoreBounds as the current values will be 0, 0 and the size of the screen
                Properties.Settings.Default.Top = RestoreBounds.Top;
                Properties.Settings.Default.Left = RestoreBounds.Left;
                Properties.Settings.Default.Height = RestoreBounds.Height;
                Properties.Settings.Default.Width = RestoreBounds.Width;
            }
            else
            {
                Properties.Settings.Default.Top = this.Top;
                Properties.Settings.Default.Left = this.Left;
                Properties.Settings.Default.Height = this.Height;
                Properties.Settings.Default.Width = this.Width;
            }
            Properties.Settings.Default.Save();
        }

        private string SaveAttachmentToTemporaryFile(XstAttachment a)
        {
            if (a == null)
                return null;

            string fileFullName = Path.ChangeExtension(
                Path.GetTempPath() + Guid.NewGuid().ToString(), Path.GetExtension(a.FileNameForSaving)); ;

            try
            {
                (listAttachments.SelectedItem as XstAttachment)?.SaveToFile(fileFullName, null);
                tempFileNames.Add(fileFullName);
                return fileFullName;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error saving attachment");
                return null;
            }
        }

        private void attachmentEmailCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            var a = listAttachments.SelectedItem as XstAttachment;
            e.CanExecute = a != null && a.IsEmail;
        }

        //private void openEmail_Executed(object sender, ExecutedRoutedEventArgs e)
        //{
        //    var a = listAttachments.SelectedItem as Attachment;
        //    OpenEmailAttachment(a);
        //}

        private void attachmentCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            var a = listAttachments.SelectedItem as XstAttachment;
            e.CanExecute = a != null;
        }

        private void attachmentFileCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            var a = listAttachments.SelectedItem as XstAttachment;
            e.CanExecute = a != null && a.IsFile;
        }

        private void openAttachment_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var a = listAttachments.SelectedItem as XstAttachment;

            if (a.IsFile)
            {
                string fileFullname = SaveAttachmentToTemporaryFile(a);
                if (fileFullname == null)
                    return;

                using (Process.Start(fileFullname)) { }
            }
            else if (a.IsEmail)
                OpenEmailAttachment(a);

            e.Handled = true;
        }

        private void openAttachmentWith_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var a = listAttachments.SelectedItem as XstAttachment;
            string fileFullname = SaveAttachmentToTemporaryFile(a);
            if (fileFullname == null)
                return;

            if (Environment.OSVersion.Version.Major > 5)
            {
                IntPtr hwndParent = Process.GetCurrentProcess().MainWindowHandle;
                tagOPENASINFO oOAI = new tagOPENASINFO
                {
                    cszFile = fileFullname,
                    cszClass = String.Empty,
                    oaifInFlags = tagOPEN_AS_INFO_FLAGS.OAIF_ALLOW_REGISTRATION | tagOPEN_AS_INFO_FLAGS.OAIF_EXEC
                };
                SHOpenWithDialog(hwndParent, ref oOAI);
            }
            else
            {
                using (Process.Start("rundll32", "shell32.dll,OpenAs_RunDLL " + fileFullname)) { }
            }
            e.Handled = true;
        }

        private void saveAttachmentAs_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (listAttachments.SelectedItems.Count > 1)
            {
                SaveAttachments(listAttachments.SelectedItems.Cast<XstAttachment>());
            }
            else
            {
                var a = listAttachments.SelectedItem as XstAttachment;
                var fullFileName = GetSaveAttachmentFileName(a.LongFileName);

                if (fullFileName != null)
                {
                    try
                    {
                        a.SaveToFile(fullFileName, view.CurrentMessage.Date);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Error saving attachment");
                    }
                }
            }
            e.Handled = true;
        }

        // Plumbing to enable access to SHOpenWithDialog
        [DllImport("shell32.dll", EntryPoint = "SHOpenWithDialog", CharSet = CharSet.Unicode)]
        private static extern int SHOpenWithDialog(IntPtr hWndParent, ref tagOPENASINFO oOAI);
        private struct tagOPENASINFO
        {
            [MarshalAs(UnmanagedType.LPWStr)]
            public string cszFile;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string cszClass;

            [MarshalAs(UnmanagedType.I4)]
            public tagOPEN_AS_INFO_FLAGS oaifInFlags;
        }
        [Flags]
        private enum tagOPEN_AS_INFO_FLAGS
        {
            OAIF_ALLOW_REGISTRATION = 0x00000001,   // Show "Always" checkbox
            OAIF_REGISTER_EXT = 0x00000002,   // Perform registration when user hits OK
            OAIF_EXEC = 0x00000004,   // Exec file after registering
            OAIF_FORCE_REGISTRATION = 0x00000008,   // Force the checkbox to be registration
            OAIF_HIDE_REGISTRATION = 0x00000020,   // Vista+: Hide the "always use this file" checkbox
            OAIF_URL_PROTOCOL = 0x00000040,   // Vista+: cszFile is actually a URI scheme; show handlers for that scheme
            OAIF_FILE_IS_URI = 0x00000080    // Win8+: The location pointed to by the pcszFile parameter is given as a URI
        }

        private void btnInfo_Click(object sender, RoutedEventArgs e)
        {
            Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string Repository = "https://github.com/Dijji/XstReader";

            StringBuilder msg = new StringBuilder(100);
            msg.AppendLine("View Microsoft Outlook Mail files");
            msg.Append("Version: ");
            msg.AppendLine(version.ToString());
            msg.Append("Repository: ");
            msg.AppendLine(Repository);

            MessageBox.Show(msg.ToString(), "About XstReader");
        }
    }
}
