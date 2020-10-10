using Microsoft.Office.Interop.Access.Dao;
using System;
using System.IO;

namespace ConsoleApp
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length != 2)
                return;

            var dbPath = args[0];
            var propValue = args[1] == "1";

            if (!File.Exists(dbPath))
            {
                Console.WriteLine($"{dbPath} does not exist.");
                return;
            }

            ChangeAllowBypassKey(dbPath, propValue);

            Console.ReadKey();
        }

        private static void ChangeAllowBypassKey(string dbPath, bool propValue)
        {
            try
            {
                var dbe = new DBEngine();
                var db = dbe.OpenDatabase(dbPath);

                Property prop = db.Properties["AllowBypassKey"];
                prop.Value = propValue;

                if (propValue)
                {
                    Console.WriteLine("Property 'AllowBypassKey' is set to 'True'.");
                    Console.WriteLine("You can access the design (developer) mode by keep pressing SHIFT key while opening the file.");
                }
                else
                {
                    Console.WriteLine("Property 'AllowBypassKey' is set to 'False'.");
                    Console.WriteLine("You can no longer use the SHIFT key to enter the design mode.");
                }
            }
            catch (Exception e)
            {
                if (e.Message.Contains("SSMA"))
                {
                    var url = @"https://www.microsoft.com/en-us/download/details.aspx?id=10910";
                    Console.WriteLine($"The following error is thrown: {e.Message}.{Environment.NewLine}Please download and install the MS Access runtime (32 bit or 64 bit).{Environment.NewLine}URL: {url}");
                }
                else
                {
                    Console.WriteLine($"The following error is thrown: {e.Message}");
                }
            }
        }
    }
}