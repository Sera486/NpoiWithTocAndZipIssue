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

            using (var fs = new FileStream("simple.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
        }
    }
}