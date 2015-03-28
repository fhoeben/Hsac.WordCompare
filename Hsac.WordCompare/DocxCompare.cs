

namespace Hsac.WordCompare
{
    public class DocxCompare
    {
        public bool UseWord { get; set; }
        public string DiffFile { get; private set; }

        public CompareResult AreEqual(string expectedFile, string actualFile)
        {
            CompareResult result;
            if (CompareHelper.CompareFiles(expectedFile, actualFile))
            {
                result = CompareResult.SameContent;
            }
            else
            {
                string diffFilename;
                result = CompareHelper.CompareDocx(expectedFile, actualFile, UseWord, out diffFilename);
                DiffFile = diffFilename;
            }
            return result;
        }
    }
}
