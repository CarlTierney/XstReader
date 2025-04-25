using System;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using XstReader.Api;


namespace XstReader.Api.Tests
{
    [TestClass]
    public sealed class ReadingStreamTests
    {
        private string _pstFilepath = "<PSTFILEPATH HERE>";

        [TestMethod]
        public void OpenFileWithXstReader()
        {
            // Arrange
            


            using var xstFile = new XstFile(_pstFilepath);


            // Assert
            Assert.IsNotNull(xstFile);
            Assert.IsNotNull(xstFile.RootFolder);
        }


        [TestMethod]
        public void OpenFileAndPassStreamToXstReader()
        {
            // Arrange
           
            using FileStream fileStream = new FileStream(_pstFilepath, FileMode.Open, FileAccess.Read);
            // Act
            using var xstFile = new XstFile(fileStream);


            // Assert
            Assert.IsNotNull(xstFile);
            Assert.IsNotNull(xstFile.RootFolder);
        }

        [TestMethod]
        public void ReadInboxFromStream()
        {
            
            using FileStream fileStream = new FileStream(_pstFilepath, FileMode.Open, FileAccess.Read);
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
