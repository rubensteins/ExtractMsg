using MsgReader.Outlook;

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
    PrintHelp("Invalid file name.");
    Environment.Exit(-1);
}

List<object> messagesToProcess = new List<object>();
string basePath = "";

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
    Console.WriteLine("(c) 2024 Ruben Steins - MIT License");
    Console.WriteLine("");
    Console.WriteLine("Usage:");
    Console.WriteLine("ExtractMsg filename [-f] [-d] [-h]");
    Console.WriteLine();
    Console.WriteLine("filename:          The name of the .msg file to extract");
    Console.WriteLine("-f     Folder.     Create a new folder called 'filename' and extract into that.");
    Console.WriteLine("-d     Diagnostic. Show some diagnostic messages.");
    Console.WriteLine("-h     Help.       Shows this text.");
}