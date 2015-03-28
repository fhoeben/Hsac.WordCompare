using System;
using System.IO;

namespace Hsac.WordCompare
{
    public class Program
    {
        public static void Main(string[] arguments)
        {
            if (arguments.Length < 2 || arguments[0] == "/?")
            {
                ShowHelp();
                Exit(0);
            }

            try
            {
                var exitCode = Compare(arguments);
                Exit(exitCode);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error comparing documents: " + e.Message);
                Exit(-1);
            }
        }

        private static int Compare(string[] arguments)
        {
            var docxCompare = new DocxCompare();
            int currentArg = 0;
            if (arguments[currentArg] == "/wordDiff")
            {
                docxCompare.UseWord = true;
                currentArg++;
            }
            var expectedFile = arguments[currentArg++];
            var actualFile = arguments[currentArg++];
            var compareResult = docxCompare.AreEqual(expectedFile, actualFile);
            int exitCode;
            if (compareResult == CompareResult.SameContent)
            {
                Console.WriteLine("Document content is identical");
                exitCode = 0;
            }
            else if (compareResult == CompareResult.EqualContent)
            {
                Console.WriteLine("Document content is not identical, but Word found no changes, replacing expected by actual");
                File.Copy(actualFile, expectedFile, true);
                exitCode = 0;
            }
            else
            {
                Console.WriteLine("Document content does not match");
                if (docxCompare.UseWord)
                {
                    Console.WriteLine("Differences between documents are stored as: " + docxCompare.DiffFile);
                }
                exitCode = 1;
            }
            return exitCode;
        }

        private static void ShowHelp()
        {
            Console.WriteLine("Compares docx document content.");
            Console.WriteLine("Usage: [<options>] <expectedDocx> <actualDocx>");
            Console.WriteLine("  options:");
            Console.WriteLine("     /wordDiff:  Consider documents equal if Word finds no revisions comparing documents");
            Console.WriteLine();
            Console.WriteLine("Exit codes:");
            Console.WriteLine("  -1: Error");
            Console.WriteLine("   0: Documents are equal");
            Console.WriteLine("   1: Differences found between documents");
        }

        private static void Exit(int exitCode)
        {
#if DEBUG
            Console.ReadLine();
#endif
            Environment.Exit(exitCode);
        }

    }
}
