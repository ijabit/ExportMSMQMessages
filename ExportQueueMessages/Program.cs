using System;
using System.Messaging;
using System.Text;
using System.Windows.Forms;

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

            // Loop over all messages and write them to a file... (in this case XML)
            q.MessageReadPropertyFilter.SetAll();
            MessageEnumerator msgEnum = q.GetMessageEnumerator2();
            int k = 0;
            Console.Write("Processing...");
            while (msgEnum.MoveNext())
            {
                if (k % 100 == 0)
                    Console.Write(".");
                System.Messaging.Message msg = msgEnum.Current;
                byte[] data = new byte[msg.BodyStream.Length];
                msg.BodyStream.Read(data, 0, (int)msg.BodyStream.Length);
                string strMessage = ASCIIEncoding.ASCII.GetString(data);

                //msg.Formatter = new ActiveXMessageFormatter();
                string fileName = outputFolderPath + "\\" + msg.ArrivedTime.ToString("yyyy-MM-dd hh.mm.ss tt") + "-" + msg.Id.Replace("\\", "-") + ".xml";
                System.IO.File.WriteAllText(fileName, strMessage);
                k++;
            }
            Console.Write(Environment.NewLine);
            Console.WriteLine("All done! Hit any key to exit.");
            Console.ReadKey();
        }
    }
}
