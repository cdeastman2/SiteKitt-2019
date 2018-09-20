using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using Microsoft.ConfigurationManagement.ManagementProvider;
using Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine;


namespace ConfigurationManagerSnippets
{
    class Program
    {
        static void Main(string[] args)
        {
            // Setup snippet class.

            string computer = "";
            string userName = "";
            string password = "";

            SnippetClass snippets = new SnippetClass();

            Console.WriteLine("Computer you want to connect to (Enter . for local): ");
            computer = Console.ReadLine();
            Console.WriteLine();

            if (computer == ".")
            {
                computer = System.Net.Dns.GetHostName();
                userName = "";
                password = "";
            }
            else
            {
                Console.WriteLine("Please enter the user name: ");
                userName = Console.ReadLine();

                Console.WriteLine("Please enter your password:");
                password = snippets.ReturnPassword();
            }

            // Make connection to provider.
            WqlConnectionManager WMIConnection = snippets.Connect(computer, userName, password);

            // Call snippet function and pass the provider connection object.
            snippets.SNIPPETMETHODNAME(WMIConnection);
        }
    }

    class SnippetClass
    {
        public WqlConnectionManager Connect(string serverName, string userName, string userPassword)
        {
            try
            {               
                SmsNamedValuesDictionary namedValues = new SmsNamedValuesDictionary();
                WqlConnectionManager connection = new WqlConnectionManager(namedValues);
                if (System.Net.Dns.GetHostName().ToUpper() == serverName.ToUpper())
                {
                    connection.Connect(serverName);
                }
                else
                {
                    connection.Connect(serverName, userName, userPassword);
                }
                return connection;
            }
            catch (SmsException ex)
            {
                Console.WriteLine("Failed to Connect. Error: " + ex.Message);
                return null;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Failed to authenticate. Error:" + ex.Message);
                return null;
            }
        }

        public void SNIPPETMETHODNAME(WqlConnectionManager connection)
        {
            // Insert snippet code here.
        }

        public string ReturnPassword()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    password += info.KeyChar;
                    info = Console.ReadKey(true);
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        password = password.Substring
                        (0, password.Length - 1);
                    }
                    info = Console.ReadKey(true);
                }
            }
            for (int i = 0; i < password.Length; i++)
                Console.Write("*");
            return password;
        }
    }
}