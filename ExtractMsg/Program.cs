using System.Net.Mail;
using MsgReader.Outlook;
using RtfPipe.Tokens;

if (args.Length == 0)
{
    Console.WriteLine("Please provide path to .msg to extract.");
    Environment.Exit(-1);
}

List<object> messagesToProcess = new List<object>();
string fileName = args[0];
var msg = new MsgReader.Outlook.Storage.Message(fileName);
messagesToProcess.Add(msg);
while (messagesToProcess.Count > 0)
{
    // if it's a message, print its text.
    // check if there are any attachements, if so add to queue
    if (messagesToProcess[0] is Storage.Message)
    {
        Storage.Message message = (messagesToProcess[0] as Storage.Message)!;
        Console.WriteLine(message.BodyText.Replace("\\n", Environment.NewLine));
        messagesToProcess.RemoveAt(0);
        if (message.Attachments.Count > 0)
        {
            messagesToProcess.AddRange(message.Attachments);
        }
    }
    else
    {
        // attachment.
        Storage.Attachment a = messagesToProcess[0] as Storage.Attachment;
        File.WriteAllBytes(a.FileName, a.Data);
        Console.WriteLine("Wrote: " + a.FileName);
        messagesToProcess.RemoveAt(0);
    }
}
Console.WriteLine("Done.");

