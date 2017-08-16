using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml;

namespace Email_Attachment_Tester
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        bool RunTestFlag = false;
        static List<Email> QueuedEmails = new List<Email>();
        static List<Email> AttachedEmails = new List<Email>();
        static List<Email> ErroredEmails = new List<Email>();

        async void RunTest(int AttachQueueLimit)
        {
            if (!ValidateFileExtension(xmlFilePath.Text, ".xml")
                | !ValidateFileExtension(msgFilePath.Text, ".msg")
                | (SlashifyDirectoryPath(historyQueuePath.Text) == null))
            {
                MessageBox.Show("File paths not valid");
                RunTestFlag = false;
                return;
            }

            XmlDocument StandardMsgXml = GetXmlDocument(xmlFilePath.Text);
            if (StandardMsgXml == null)
            {
                MessageBox.Show("Unable to load XML file");
                RunTestFlag = false;
                return;
            }

            Guid SessionGuid = Guid.NewGuid();

            int EmailNumber = 0;

            while (RunTestFlag)
            {
                while (QueuedEmails.Count() < AttachQueueLimit)
                {
                    EmailNumber++;
                    Email Message = new Email();
                    Message.Number = EmailNumber;
                    Message.SessionGuid = SessionGuid;
                    Message.Name = "TestEmail" + Message.Number + "-" + SessionGuid;

                    if (!QueueEmail(Message, StandardMsgXml, SlashifyDirectoryPath(historyQueuePath.Text), msgFilePath.Text))
                    {
                        MessageBox.Show("Failing to queue messages.");
                        RunTestFlag = false;
                        return;
                    }
                }

                for (int i = 0; i < QueuedEmails.Count; i++)
                {
                    Email Message = QueuedEmails[i];
                    if (CheckIfEmailAttached(Message, SlashifyDirectoryPath(historyQueuePath.Text)))
                    {
                        AttachedEmails.Add(Message);
                        QueuedEmails.Remove(Message);
                        i--;
                        continue;
                    }
                    if (CheckIfEmailErrored(Message, SlashifyDirectoryPath(historyQueuePath.Text)))
                    {
                        //Message.Error = GetEmailError(); // Implement later, maybe
                        ErroredEmails.Add(Message);
                        QueuedEmails.Remove(Message);
                        i--;
                        continue;
                    }
                }
                successCount.Content = AttachedEmails.Count.ToString();
                failCount.Content = ErroredEmails.Count.ToString();

                await Task.Delay(1000);
            }
        }

        bool CheckIfEmailAttached(Email Message, string HistoryQueuePath)
        {
            try
            {
                using (StreamReader Reader = new StreamReader(HistoryQueuePath + "Logs\\successfullhistories.xml"))
                {
                    string Contents = Reader.ReadToEnd();
                    if (Contents.Contains(Message.Name))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        bool CheckIfEmailErrored(Email Message, string HistoryQueuePath)
        {
            try
            {
                using (StreamReader Reader = new StreamReader(HistoryQueuePath + "Logs\\nothandledmessages.xml"))
                {
                    string Contents = Reader.ReadToEnd();
                    if (Contents.Contains(Message.Name))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        bool QueueEmail(Email Message, XmlDocument StandardMsgXml, string HistoryQueuePath, string MsgFilePath)
        {
            string NewMsgFilePath = HistoryQueuePath + Message.Name + ".msg";
            string NewXmlFilePath = HistoryQueuePath + Message.Name + ".xml";

            string NewMsgFilePathEncoded = EncodeFilePath(NewMsgFilePath);

            XmlDocument NewXml = StandardMsgXml;
            NewXml = WriteXmlAttribute(NewXml, "outlookhistory", "messagefilelocation", NewMsgFilePathEncoded);

            Message.EncodedLocation = NewMsgFilePathEncoded;

            try
            {
                NewXml.Save(NewXmlFilePath);
            }
            catch
            {
                return false;
            }

            try
            {
                CopyFile(MsgFilePath, HistoryQueuePath, "TestEmail" + Message.Number + "-" + Message.SessionGuid + ".msg");
            }
            catch
            {
                return false;
            }

            QueuedEmails.Add(Message);

            return true;
        }

        string EncodeFilePath(string Path)
        {
            string NewString = Path.Replace(":", "_x003A_").Replace(@"\", "_x005C_").Replace(@" ", "_x0020_");

            return NewString;
        }

        bool ValidateFileExtension(string Path, string Extension)
        {
            if(!Path.EndsWith(Extension) | !File.Exists(Path))
                return false;
            return true;
        }

        string SlashifyDirectoryPath(string Path)
        {
            if (!Directory.Exists(Path))
                return null;

            if (!Path.EndsWith("\\"))
            {
                Path = Path + "\\";
            }

            return Path;
        }

        XmlDocument GetXmlDocument(string Path)
        {
            XmlDocument XmlDoc = new XmlDocument();
            try
            {
                XmlDoc.Load(Path);
            }
            catch
            {
                return null;
            }
            return XmlDoc;
        }

        string ReadXmlValue(XmlDocument XmlDoc, string XPath)
        {
            XmlNode Node = null;
            if (XmlDoc.SelectSingleNode(XPath) != null)
            {
                Node = XmlDoc.SelectSingleNode(XPath);
            }
            else return null;

            if (Node.InnerText != null)
            {
                return Node.InnerText;
            }
            else return null;
        }

        XmlDocument WriteXmlValue(XmlDocument XmlDoc, string XPath, string Value)
        {
            XmlNode Node = null;
            if (XmlDoc.SelectSingleNode(XPath) != null)
            {
                Node = XmlDoc.SelectSingleNode(XPath);
            }
            else return null;

            try
            {
                Node.InnerText = Value;
            }
            catch
            {
                return null;
            }
            return XmlDoc;
        }

        XmlDocument WriteXmlAttribute(XmlDocument XmlDoc, string XPath, string Attribute, string Value)
        {
            XmlNode Node = null;
            if (XmlDoc.SelectSingleNode(XPath) != null)
            {
                Node = XmlDoc.SelectSingleNode(XPath);
            }
            else return null;

            try
            {
                XmlAttribute NewAttribute = XmlDoc.CreateAttribute(Attribute);
                NewAttribute.Value = Value;
                Node.Attributes.SetNamedItem(NewAttribute);
            }
            catch
            {
                return null;
            }
            return XmlDoc;
        }

        bool CopyFile(string SourceFile, string NewPath, string NewName)
        {
            if (!File.Exists(SourceFile))
                return false;

            if (!Directory.Exists(NewPath))
                return false;

            NewPath = SlashifyDirectoryPath(NewPath);

            try
            {
                File.Copy(SourceFile, NewPath + NewName);
            }
            catch
            {
                return false;
            }
            return true;
        }

        private void FilePath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void FilePath_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                (sender as TextBox).Text = files[0];
            }
        }

        void StartTest_Click(object sender, RoutedEventArgs e)
        {
            QueuedEmails.Clear();
            AttachedEmails.Clear();
            ErroredEmails.Clear();

            RunTestFlag = true;
            RunTest(Convert.ToInt32(queueCount.Text));

            msgFilePath.IsEnabled = false;
            xmlFilePath.IsEnabled = false;
            historyQueuePath.IsEnabled = false;
            queueCount.IsEnabled = false;
            startButton.IsEnabled = false;
            stopButton.IsEnabled = true;
        }

        void StopTest_Click(object sender, RoutedEventArgs e)
        {
            RunTestFlag = false;

            msgFilePath.IsEnabled = true;
            xmlFilePath.IsEnabled = true;
            historyQueuePath.IsEnabled = true;
            queueCount.IsEnabled = true;
            startButton.IsEnabled = true;
            stopButton.IsEnabled = false;
        }

        private void TextInput_NumbersOnly(object sender, TextCompositionEventArgs e)
        {
            Regex AllowedText = new Regex("[^0-9]+");
            e.Handled = AllowedText.IsMatch(e.Text);
        }
    }

    class Email
    {
        public DateTime CreateTime = DateTime.Now;
        public string Error { get; set; }
        public int Number { get; set; }
        public string EncodedLocation { get; set; }
        public string Name { get; set; }
        public Guid SessionGuid { get; set; }
    }
}
