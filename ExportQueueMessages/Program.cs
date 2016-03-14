using System;
using System.Collections.Generic;
using System.IO;
using System.Messaging;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Message = System.Messaging.Message;

namespace ExportQueueMessages
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Queues on This Machine:");
            var queues = MessageQueue.GetPrivateQueuesByMachine(Environment.MachineName);
            foreach (var queue in queues)
            {
                Console.WriteLine(".\\" + queue.QueueName);
            }

            // TODO: make below only run when on a domain joined machine
            //queues = MessageQueue.GetPublicQueuesByMachine(Environment.MachineName);
            //foreach (var queue in queues)
            //{
            //    Console.WriteLine(".\\" + queue.QueueName);
            //}
            Console.WriteLine("FormatName:DIRECT=OS:.\\System$;Deadletter");
            Console.WriteLine("FormatName:DIRECT=OS:.\\System$;Deadxact");

            string queueName = string.Empty;
            while (string.IsNullOrEmpty(queueName))
            {
                Console.WriteLine("Enter queue name: ");
                queueName = Console.ReadLine();
                if (string.IsNullOrEmpty(queueName) == false)
                {
                    try
                    {
                        if (!string.Equals(queueName, "FormatName:DIRECT=OS:.\\System$;Deadletter", StringComparison.InvariantCultureIgnoreCase)
                            && !string.Equals(queueName, "FormatName:DIRECT=OS:.\\System$;Deadxact", StringComparison.InvariantCultureIgnoreCase)
                            && MessageQueue.Exists(queueName) == false)
                        {
                            // Valid queue path, but doesn't exist
                            Console.WriteLine("Queue {0} not found.", queueName);
                            queueName = string.Empty;
                        }
                    }
                    catch (Exception)
                    {
                        // Invalid queue path formatting
                        Console.WriteLine("Queue {0} not found.", queueName);
                        queueName = string.Empty;
                    }
                }
            }
            Console.WriteLine("Queue Found!");

            var outputFolderDialog = new System.Windows.Forms.FolderBrowserDialog()
            {
                Description = "Queue Output Folder", ShowNewFolderButton = true
            };
            var dialogResult = outputFolderDialog.ShowDialog();
            while (dialogResult != DialogResult.OK)
            {
                dialogResult = outputFolderDialog.ShowDialog();
            }
            var outputFolderPath = outputFolderDialog.SelectedPath;

            //Setup MSMQ using path from user...
            MessageQueue q = new MessageQueue(queueName);

            // Loop over all messages and write them to a file...
            q.MessageReadPropertyFilter.SetAll();
            MessageEnumerator msgEnum = q.GetMessageEnumerator2();
            int k = 0;
            Console.Write("Processing...");
            while (msgEnum.MoveNext())
            {
                if (k % 100 == 0)
                    Console.Write(".");
                Message msg = msgEnum.Current;
                byte[] data = new byte[msg.BodyStream.Length];
                msg.BodyStream.Read(data, 0, (int)msg.BodyStream.Length);
                string messageContent = ASCIIEncoding.ASCII.GetString(data);

                string fileName = outputFolderPath + "\\" + msg.ArrivedTime.ToString("yyyy-MM-dd hh.mm.ss tt") + "-" + msg.Id.Replace("\\", "-") + ".xml";
                string messageWithHeaders = $"<?xml version=\"1.0\"?><MsmqMessage><Id>{msg.Id}</Id><ArrivedTime>{msg.ArrivedTime}</ArrivedTime><Label>{msg.Label}</Label><Class>{msg.Acknowledgment}</Class><Headers>";
                var headers = DeserializeMessageHeaders(msg);
                foreach (var header in headers)
                {
                    messageWithHeaders += $"<Header><Name>{header.Key}</Name><Value><![CDATA[{header.Value}]]></Value></Header>";
                }
                messageWithHeaders += $"</Headers><Content><![CDATA[{messageContent}]]></Content></MsmqMessage>";
                XDocument formattedXml = XDocument.Parse(messageWithHeaders);
                System.IO.File.WriteAllText(fileName, formattedXml.ToString());

                k++;
            }
            Console.Write(Environment.NewLine);
            Console.WriteLine("All done! Hit any key to exit.");
            Console.ReadKey();
        }

        private static Dictionary<string, string> DeserializeMessageHeaders(Message m)
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            if (m.Extension.Length == 0)
                return dictionary;
            object obj;
            using (StringReader stringReader1 = new StringReader(Encoding.UTF8.GetString(m.Extension).TrimEnd(new char[1])))
            {
                StringReader stringReader2 = stringReader1;
                XmlReaderSettings settings = new XmlReaderSettings()
                {
                    CheckCharacters = false
                };
                XmlSerializer headerSerializer = new XmlSerializer(typeof(List<HeaderInfo>));
                using (XmlReader xmlReader = XmlReader.Create((TextReader)stringReader2, settings))
                    obj = headerSerializer.Deserialize(xmlReader);
            }
            foreach (HeaderInfo headerInfo in (List<HeaderInfo>)obj)
            {
                if (headerInfo.Key != null)
                    dictionary.Add(headerInfo.Key, headerInfo.Value);
            }
            return dictionary;
        }
    }

    [Serializable]
    public class HeaderInfo
    {
        /// <summary>
        /// The key used to lookup the value in the header collection.
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// The value stored under the key in the header collection. 
        /// </summary>
        public string Value { get; set; }
    }
}
