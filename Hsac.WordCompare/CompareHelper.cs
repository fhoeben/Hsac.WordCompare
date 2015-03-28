using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;

namespace Hsac.WordCompare
{
    public class CompareHelper
    {
        public static bool CompareFiles(string fileName1, string fileName2)
        {
            bool result;
            if (GetFileSize(fileName1) != GetFileSize(fileName2))
            {
                result = false;
            }
            else
            {
                using (var file1 = new FileStream(fileName1, FileMode.Open))
                using (var file2 = new FileStream(fileName2, FileMode.Open))
                    result = CompareStreams(file1, file2);
            }
            return result;
        }

        public static bool CompareStreams(Stream stream1, Stream stream2)
        {
            const int bufferSize = 4096;
            var buffer1 = new byte[bufferSize];
            var buffer2 = new byte[bufferSize];
            while (true)
            {
                var count1 = stream1.Read(buffer1, 0, bufferSize);
                var count2 = stream2.Read(buffer2, 0, bufferSize);

                if (count1 != count2)
                    return false;

                if (count1 == 0)
                    return true;

                if (!MemCmp(buffer1, buffer2, count1))
                    return false;
            }
        }

        private static long GetFileSize(string fileName)
        {
            return new FileInfo(fileName).Length;
        }

        public static CompareResult CompareDocx(string expectedFileName, string outputFilename, bool useWordToCompareDocx, out string diffFilename)
        {
            CompareResult result;
            diffFilename = null;
            var zipCompare = CompareZips(expectedFileName, outputFilename);
            if (useWordToCompareDocx)
            {
                if (zipCompare.Any())
                {
                    // use word to get better info on what is wrong
                    result = CompareUsingWord(expectedFileName, outputFilename, out diffFilename);
                }
                else
                {
                    result = CompareResult.SameContent;
                }
            }
            else
            {
                var allDiffs = zipCompare.ToList();
                result = allDiffs.Any() ? CompareResult.Different : CompareResult.SameContent;
            }
            return result;
        }

        private static IEnumerable<KeyValuePair<string, string>> CompareZips(string expectedFileName, string outputFilename)
        {
            using (var expectedFile = File.OpenRead(expectedFileName))
            using (var expectedZip = new ZipArchive(expectedFile))
            using (var actualFile = File.OpenRead(outputFilename))
            using (var actualZip = new ZipArchive(actualFile))
            {
                ICollection<string> expEntries = expectedZip.Entries.Select(entry => entry.FullName).ToList();
                ICollection<string> actEntries = actualZip.Entries.Select(entry => entry.FullName).ToList();
                var missingFiles = expEntries.Where(x => !actEntries.Contains(x));
                foreach (var missingFile in missingFiles)
                {
                    yield return new KeyValuePair<string, string>(missingFile, "File present in expected, not in actual");
                }
                var extraFiles = actEntries.Where(x => !expEntries.Contains(x));
                foreach (var extraFile in extraFiles)
                {
                    yield return new KeyValuePair<string, string>(extraFile, "File present in actual, not in expected");
                }
                for (int i = 0; i < expEntries.Count; i++)
                {
                    var expEntry = expectedZip.Entries[i];
                    var actEntry = actualZip.Entries.FirstOrDefault(x => x.FullName == expEntry.FullName);
                    if (actEntry != null)
                    {
                        var expLength = expEntry.Length;
                        var actLength = actEntry.Length;
                        if (expLength != actLength)
                        {
                            yield return new KeyValuePair<string, string>(actEntry.FullName, "Different Length: " + expLength + " vs. " + actLength);
                            continue;
                        }
                        if (!CompareStreams(expEntry.Open(), actEntry.Open()))
                        {
                            yield return new KeyValuePair<string, string>(actEntry.FullName, "Content differs");
                        }
                    }
                }
            }
        }

        [DllImport("msvcrt.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int memcmp(byte[] b1, byte[] b2, long count);

        private static bool MemCmp(byte[] b1, byte[] b2, int count)
        {
            return memcmp(b1, b2, count) == 0;
        }

        public static bool CompareByteArray(byte[] b1, byte[] b2)
        {
            // Validate buffers are the same length.
            // This also ensures that the count does not exceed the length of either buffer.  
            return b1.Length == b2.Length && MemCmp(b1, b2, b1.Length);
        }

        private static CompareResult CompareUsingWord(string expectedFilename, string outputFilename, out string diffFilename)
        {
            CompareResult result = CompareResult.Different;
            diffFilename = null;
            Application wordApp = null;
            Document expDoc = null;

            try
            {
                wordApp = new Application();
                expDoc = OpenDocument(wordApp, expectedFilename);
                var actualPath = Path.GetFullPath(outputFilename);
                expDoc.Compare(actualPath, "Comparison", WdCompareTarget.wdCompareTargetCurrent, true, true, false, false, false);
                if (GetRevisionCount(expDoc) == 0)
                {
                    result = CompareResult.EqualContent;
                }
                else
                {
                    diffFilename = Path.ChangeExtension(actualPath, ".diff.docx");
                    expDoc.TrackRevisions = true;
                    expDoc.SaveAs(diffFilename);
                }
            }
            finally
            {
                CleanUpDoc(expDoc);
                CleanUpApp(wordApp);

#if DEBUG
                if (diffFilename != null)
                {
                    // Open the result
                    System.Diagnostics.Process.Start(diffFilename);
                }
#endif
            }
            return result;
        }

        private static int GetRevisionCount(Document expDoc)
        {
            int revisionSum = 0;
            revisionSum += expDoc.Revisions.Count;
            foreach (Section section in expDoc.Sections)
            {
                var index = section.Index;
                revisionSum += section.Range.Revisions.Count;
                foreach (HeaderFooter header in section.Headers)
                {
                    revisionSum += header.Range.Revisions.Count;
                }
                foreach (HeaderFooter footer in section.Footers)
                {
                    revisionSum += footer.Range.Revisions.Count;
                }
            }
            return revisionSum;
        }

        private static Document OpenDocument(Application application, string filename)
        {
            object objFile = Path.GetFullPath(filename);
            return application.Documents.Open(ref objFile);
        }

        private static void CleanUpApp(Application wordApp)
        {
            if (wordApp != null)
            {
                wordApp.Quit(false);
                Marshal.FinalReleaseComObject(wordApp);
            }
        }

        private static void CleanUpDoc(Document doc)
        {
            if (doc != null)
            {
                doc.Close(false);
                Marshal.FinalReleaseComObject(doc);
            }
        }
    }
}
