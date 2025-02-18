using MsgReader.Outlook;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

bool showDiag = false; // show diagnostic messages
bool useFolder = false; // no seperate folder
string fileName = "";

if (args.Length == 0)
{
    PrintHelp("No file provided.");
    Environment.Exit(-1);
}
else
{
    ProcessArguments();
}

if (!File.Exists(fileName))
{
    PrintHelp("File not found.");
    Environment.Exit(-1);
}

List<object> messagesToProcess = new List<object>();
string basePath = "";

try
{
    // Open the Message
    using (var msg = new Storage.Message(fileName))
    {
        messagesToProcess.Add(msg);
        if (useFolder)
        {
            basePath = msg.Subject + Path.DirectorySeparatorChar;
            var dir = Directory.CreateDirectory(basePath);
            if (showDiag)
            {
                Console.WriteLine($"Created folder {dir.Name}");
            }
        }
        
        while (messagesToProcess.Count > 0)
        {
            // if it's a message, write its text.
            if (messagesToProcess[0] is Storage.Message)
            {
                Storage.Message message = (messagesToProcess[0] as Storage.Message)!;

                File.WriteAllText($"{basePath}{message.Subject}.txt", message.BodyText);
                messagesToProcess.RemoveAt(0);
                
                // msg might contain attachments, process them as well
                if (message.Attachments.Count > 0)
                {
                    messagesToProcess.AddRange(message.Attachments);
                }
            }
            else
            {
                // process attachment
                Storage.Attachment a = (messagesToProcess[0] as Storage.Attachment)!;
                if (a != null)
                {
                    var fileToWrite = $"{basePath}{a.FileName}";
                    if (File.Exists($"{basePath}{a.FileName}"))
                    {
                        // Sometimes nested files have the same name :(
                        fileToWrite = $"{basePath}copy of {a.FileName}";
                    }
                    
                    File.WriteAllBytes(fileToWrite, a.Data);
                    if(showDiag) Console.WriteLine($"Wrote {a.Data.Length} bytes to {a.FileName}");
                }
                messagesToProcess.RemoveAt(0);
            }
        }
    }
}
catch (Exception e)
{
    Console.WriteLine($"There was an unhandled error of type {e.GetType()}: {e.Message}");
}

if(showDiag) Console.WriteLine("Done.");

void ProcessArguments()
{
    for (int c = 0; c < args.Length; ++c)
    {
        if (args[c].StartsWith("-"))
        {
            switch (args[c])
            {
                case ("-d"):
                {
                    showDiag = true;
                    Console.WriteLine("Diagnostics on.");
                    break;
                }
                case "-f":
                {
                    useFolder = true;
                    break;
                }
                case "-h":
                {
                    PrintHelp();
                    Environment.Exit(0);
                    break;
                }
                case "-v":
                    PrintVersionInfo();
                    Environment.Exit(0);
                    break;
                default:
                {
                    PrintHelp($"Unknown parameter {args[c]}");
                    Environment.Exit(-1);
                    break;
                }
            }
        }
        else
        {
            fileName = args[c];
        }
    }
}

void PrintHelp(string error = "")
{
    if (!string.IsNullOrEmpty(error))
    {
        Console.WriteLine($"Error: {error}");
        Console.WriteLine();
    }
    Console.WriteLine("ExtractMsg will open up a .msg message.");
    Console.WriteLine("Body text will be written as text-file and (nested) attachments exported.");
    Console.WriteLine("(c) 2024,2025 Ruben Steins - MIT License");
    Console.WriteLine("");
    Console.WriteLine("Usage:");
    Console.WriteLine("ExtractMsg <filename> [-d] [-f] [-h] [-v]");
    Console.WriteLine();
    Console.WriteLine("filename:          The name of the .msg file to extract.");
    Console.WriteLine("-d     Diagnostic. Show some diagnostic messages.");
    Console.WriteLine("-f     Folder.     Create a new folder called 'filename' and extract into that.");
    Console.WriteLine("-h     Help.       Shows this text.");
    Console.WriteLine("-v     Version.    Print version info for this application.");
}

void PrintVersionInfo()
{
    Console.WriteLine("ExtractMsg - Version 1.1 - 20250218");
}
