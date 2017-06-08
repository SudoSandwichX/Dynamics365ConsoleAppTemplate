using System;
using Microsoft.Xrm.Tooling.Connector;
using System.Security;
using System.IO;

/// <summary>
/// Simple console application for authenticating to Dynamics 365 (CRM).
/// </summary>
namespace DynamicsCrmConsoleApp
{
    class Program
    {
        #region Static Variables
        private static CrmServiceClient conn;
        private static bool isDebugMode = false;
        #endregion

        static void Main(string[] args)
        {
            #region Build Dynamics 365 Connection
            bool isConnected = false;

            // Allow retries on bad login
            while (!isConnected)
            {
                conn = GetCrmConnection();

                // Check if connection is good.
                if (!conn.IsReady)
                {
                    if (conn.LastCrmError != string.Empty)
                        Console.WriteLine($"Error: {conn.LastCrmError}");

                    if (conn.LastCrmException != null)
                        Console.WriteLine($"Exception: {conn.LastCrmException}");

                    Console.WriteLine("\nPress Esacpe to exit. Any other key to try again...");

                    if (Console.ReadKey().Key == ConsoleKey.Escape)
                        return;

                    Console.WriteLine("");
                    Console.Clear();
                }
                else
                {
                    isConnected = true;
                }
            }

            Console.WriteLine("Success: Connected to Dynamics 365\n");
            #endregion

            // Connected to Dyanmics 365 (CRM)
            Console.WriteLine("\nPress any key to exit.");
            Console.ReadKey();

        }

        #region Supporting Methods
        /// <summary>
        /// Get secured password from user
        /// </summary>
        /// <returns>password as SecureString</returns>
        private static SecureString GetPassword()
        {
            SecureString result = new SecureString();
            ConsoleKeyInfo key;
            bool isEnterKey = false;

            do
            {
                key = Console.ReadKey(true);

                // Ignore any key out of range.
                isEnterKey = key.Key == ConsoleKey.Enter;
                if (!isEnterKey)
                {
                    if (key.Key == ConsoleKey.Backspace && result.Length > 0)
                    {
                        Console.Write("\b \b");
                        if (result.Length > 0)
                            result.RemoveAt(result.Length - 1);
                    }
                    else if (key.Key != ConsoleKey.Backspace)
                    {
                        // Append the character to the password.
                        result.AppendChar(key.KeyChar);
                        Console.Write("*");
                    }
                }

                // Exit if Enter key is pressed.
            } while (!isEnterKey);
            return result;
        }

        /// <summary>
        /// Build connection to Dynamics 365 online
        /// </summary>
        /// <returns>CRM Online connection as CrmServiceClient</returns>
        private static CrmServiceClient GetCrmConnection()
        {
            string orgName = "";
            string crmUserId = "";
            SecureString crmPassword = new SecureString();
            string region = String.Empty;

            Console.Write("Dynamics 365 Organization Name: ");
            orgName = Console.ReadLine();

            Console.Write("Enter Username <JohnDoe | JohnDoe@test.onmicrosoft.com>: ");
            crmUserId = Console.ReadLine();

            if (!crmUserId.Contains("@"))
            {
                crmUserId = $"{crmUserId}@{orgName}.onmicrosoft.com";
            }

            Console.Write("Enter password: ");
            crmPassword = GetPassword();
            Console.WriteLine();

            // Get region input
            Console.WriteLine("\nIs Organization in North America? ");
            Console.WriteLine(
                "   1. Yes\n" +
                "   2. No\n" +
                "   *Values other than 1 will perform a search in all geographies. Not choosing \"Yes\" may slow down login times.\n");
            Console.Write("Enter value: ");

            if (Console.ReadKey().Key == ConsoleKey.D1)
            {
                region = "NorthAmerica";
            }

            string regionOutput = region != string.Empty ? region : "All";
            Console.WriteLine($" ... Region: {regionOutput}");

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();

            Console.WriteLine("\nConnecting...");
            // CRM Connection
            CrmServiceClient conn = new CrmServiceClient(crmUserId, crmPassword, region, orgName, isOffice365: true);

            return conn;
        }

        /// <summary>
        /// Simplified logging for debugging.
        /// Enabled when isDebugging == true
        /// </summary>
        /// <param name="logMessage"></param>
        private static async void Log(string logMessage)
        {
            if (isDebugMode)
            {
                string path = Directory.GetCurrentDirectory();

                // Write the text asynchronously to a new file named "WriteTextAsync.txt".
                using (StreamWriter outputFile = File.AppendText(path + @"\Log.txt"))
                {
                    await outputFile.WriteLineAsync(logMessage);
                }
            }
        }
        #endregion
    }
}
