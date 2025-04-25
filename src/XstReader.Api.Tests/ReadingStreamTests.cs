using System;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using XstReader.Api;


namespace XstReader.Api.Tests
{
    [TestClass]
    public sealed class ReadingStreamTests
    {

        [TestMethod]
        public void OpenFileWithXstReader()
        {
            // Arrange
            string filePath = Path.Combine("D:\\ediscovery\\Test_Export\\01.20.2025-1204PM\\Exchange\\carl.tierney@vision2.com.pst");


            using var xstFile = new XstFile(filePath);


            // Assert
            Assert.IsNotNull(xstFile);
            Assert.IsNotNull(xstFile.RootFolder);
        }


        [TestMethod]
        public void OpenFileAndPassStreamToXstReader()
        {
            // Arrange
            string filePath = Path.Combine("D:\\ediscovery\\Test_Export\\01.20.2025-1204PM\\Exchange\\carl.tierney@vision2.com.pst");
            using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            // Act
            using var xstFile = new XstFile(fileStream);


            // Assert
            Assert.IsNotNull(xstFile);
            Assert.IsNotNull(xstFile.RootFolder);
        }

        [TestMethod]
        public void ReadInboxFromStream()
        {
            string filePath = Path.Combine("D:\\ediscovery\\Test_Export\\01.20.2025-1204PM\\Exchange\\carl.tierney@vision2.com.pst");
            using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            // Act
            using var xstFile = new XstFile(fileStream);


            // Assert
            Assert.IsNotNull(xstFile);
            Assert.IsNotNull(xstFile.RootFolder);


            FindInboxMessages(xstFile.RootFolder);
        }

        public bool FindInboxMessages(XstFolder folder)
        {
            if (folder.DisplayName == "Inbox")
            {
                foreach (var m in folder.Messages)
                {
                    Debug.WriteLine($"{m.InternetMessageId} - {m.Subject} - {m.Date} - has attachments:{m.HasAttachments}");
                    foreach (var a in m.Attachments)
                    {
                        Debug.WriteLine($"Attachment: {a.DisplayName} - {a.Size} bytes");
                    }
                }
                    
                return true; 
            }
            
            foreach(var f in folder.Folders)
            {
                if (FindInboxMessages(f))
                    break;
            }

            return false;
        }

    }
}
