﻿using MsgReader.Outlook;

bool showDiag = false; // show diagnostic messages
bool useFolder = false; // no seperate folder

if (args.Length == 0)
{
    PrintHelp();
    Environment.Exit(-1);
}
else
{
    ProcessArguments();
}

List<object> messagesToProcess = new List<object>();
string fileName = args[0];

// Open the Message
using (var msg = new MsgReader.Outlook.Storage.Message(fileName))
{
    messagesToProcess.Add(msg);
    while (messagesToProcess.Count > 0)
    {
        // if it's a message, write its text.
        if (messagesToProcess[0] is Storage.Message)
        {
            Storage.Message message = (messagesToProcess[0] as Storage.Message)!;
            File.WriteAllText(message.Subject + ".txt", message.BodyText);
            //Console.WriteLine(message.BodyText.Replace("\\n", Environment.NewLine));
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
                File.WriteAllBytes(a.FileName, a.Data);
                if(showDiag) Console.WriteLine("Wrote: " + a.FileName);
            }
            messagesToProcess.RemoveAt(0);
        }
    }
}

Console.WriteLine("Done.");

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
                    break;
                }
                default:
                {
                    PrintHelp();
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

void PrintHelp()
{
    Console.WriteLine("ExtractMsg will open up a .msg message.");
    Console.WriteLine("(c) 2024 Ruben Steins - MIT License");
    Console.WriteLine("Body text will be written as text-file and (nested) attachments exported.");
    Console.WriteLine("");
    Console.WriteLine("Usage:");
    Console.WriteLine("ExtractMsg filename [-f] [-d] [-h]");
    Console.WriteLine();
    Console.WriteLine("filename:          The name of the .msg file to extract");
    Console.WriteLine("-f     Folder.     Create a new folder called 'filename' and extract into that.");
    Console.WriteLine("-d     Diagnostic. Show some diagnostic messages.");
    Console.WriteLine("-h     Help.       Shows this text.");
}