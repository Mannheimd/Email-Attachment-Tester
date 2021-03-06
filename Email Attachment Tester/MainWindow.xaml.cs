﻿using System;
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
        static List<EmailFile> EmailFilePool = new List<EmailFile>();
        int PendingAttach = 0;

        async void RunTest(int AttachQueueLimit)
        {
            string HistoryQueuePath = historyQueuePath.Text;

            foreach (EmailFile file in EmailFilePool)
            {
                if (!ValidateFileExtension(file.XmlFilePath, ".xml")
                | !ValidateFileExtension(file.MsgFilePath, ".msg"))
                {
                    MessageBox.Show("An email file path isn't valid: \n"
                        + file.XmlFilePath + "\n"
                        + file.MsgFilePath);
                    RunTestFlag = false;
                    return;
                }

                file.StandardMsgXml = GetXmlDocument(file.XmlFilePath);
                if (file.StandardMsgXml == null)
                {
                    MessageBox.Show("Unable to load an XML file: \n"
                        + file.XmlFilePath + "\n"
                        + file.MsgFilePath);
                    RunTestFlag = false;
                    return;
                }
            }

            if (SlashifyDirectoryPath(HistoryQueuePath) == null)
            {
                MessageBox.Show("HistoryQueue path not valid");
                RunTestFlag = false;
                return;
            }

            int AttachDelay = 0;
            if (attachDelay.Text != null && attachDelay.Text != "")
            {
                AttachDelay = Convert.ToInt32(attachDelay.Text) * 1000;
            }

            Guid SessionGuid = Guid.NewGuid();

            int EmailNumber = 0;
            PendingAttach = 0;

            Random EmailFileChooser = new Random();

            while (RunTestFlag && EmailFilePool.Count() != 0)
            {
                Parallel.For(0, AttachQueueLimit - (QueuedEmails.Count() + PendingAttach), async i =>
                {
                    EmailFile emailFile = EmailFilePool[EmailFileChooser.Next(EmailFilePool.Count)];

                    EmailNumber++;
                    PendingAttach++;
                    Email Message = new Email();
                    Message.Number = EmailNumber;
                    Message.SessionGuid = SessionGuid;
                    Message.Name = "TestEmail" + Message.Number + "-" + SessionGuid;

                    await QueueEmail(Message, emailFile.StandardMsgXml, SlashifyDirectoryPath(HistoryQueuePath), emailFile.MsgFilePath, AttachDelay);
                });

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

        async Task<bool> QueueEmail(Email Message, XmlDocument StandardMsgXml, string HistoryQueuePath, string MsgFilePath, int AttachDelay)
        {
            await Task.Delay(AttachDelay);

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
            catch (Exception error)
            {
                MessageBox.Show("Failed to attach an email properly.\n" + error.Message);
                return false;
            }

            try
            {
                CopyFile(MsgFilePath, HistoryQueuePath, "TestEmail" + Message.Number + "-" + Message.SessionGuid + ".msg");
            }
            catch (Exception error)
            {
                MessageBox.Show("Failed to attach an email properly.\n" + error.Message);
                return false;
            }

            QueuedEmails.Add(Message);
            PendingAttach--;

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

        void AddFileToPool_Click(object sender, RoutedEventArgs e)
        {
            EmailFile file = new EmailFile()
            {
                MsgFilePath = msgFilePath.Text,
                XmlFilePath = xmlFilePath.Text
            };

            EmailFilePool.Add(file);

            filesInPool.Content = EmailFilePool.Count();
        }

        void ResetFilePool_Click(object sender, RoutedEventArgs e)
        {
            EmailFilePool.Clear();

            filesInPool.Content = EmailFilePool.Count();
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

    class EmailFile
    {
        public string MsgFilePath { get; set; }
        public string XmlFilePath { get; set; }
        public XmlDocument StandardMsgXml { get; set; }
    }
}
