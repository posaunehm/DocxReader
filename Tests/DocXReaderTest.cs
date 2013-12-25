using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using DocxReader;

namespace Tests
{
    [TestFixture]
    public class DocXReaderTest
    {
        [TestCase]
        public void ファイルをオープンしテキストを出力する()
        {
            var sut = new DocXReader(@"Data\Hello World.docx");

            sut.Text.Is("Hello World!");
        }

        [TestCase]
        public void ファイルをオープンしテキストを出力する_()
        {
            var sut = new DocXReader(@"Data\FooBar.docx");

            sut.Text.Is("FooBar");
        }
    }
}
