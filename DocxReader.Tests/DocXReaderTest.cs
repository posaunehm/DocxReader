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

        [TestCase]
        public void ファイルをオープンしプロパティを取得する()
        {
            var sut = new DocXReader(@"Data\FileProperties.docx");

            sut.Properties.Title.Is("Title");
            sut.Properties.SubTitle.Is("サブタイトル");
            sut.Properties.Creator.Is("posaunehm");
            sut.Properties.Keywords.Is("Keyword");
            sut.Properties.Description.Is("");
            sut.Properties.Category.Is("分類");
            sut.Properties.Status.Is("状態");
            sut.Properties.CreatedDate.Month.Is(12);
            sut.Properties.CreatedDate.Day.Is(25);
        }
    }
}
