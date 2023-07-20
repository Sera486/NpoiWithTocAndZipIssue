/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.XWPF.UserModel;
using System.IO;
using System.IO.Compression;

namespace SimpleDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            var doc = new XWPFDocument();

            doc.CreateTOC();

            using (var zipStream = new FileStream("test.zip", FileMode.Create))
            {
                using (var zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Create))
                {
                    var entry = zipArchive.CreateEntry($"simple-broken.docx", CompressionLevel.Optimal);
                    using (var entryStream = entry.Open())
                    {
                        doc.Write(entryStream);
                    }
                }
            }
        }
    }
}