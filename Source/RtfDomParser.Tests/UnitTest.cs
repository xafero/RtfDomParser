using System.IO;
using NUnit.Framework;

namespace RtfDomParser.Tests
{
    public class Tests
    {
        [Test]
        [TestCase("Text box", 2767)]
        [TestCase("Simple table", 1820)]
        [TestCase("NestTable", 4410)]
        [TestCase("Simple text", 1850)]
        [TestCase("htmlrtf3", 9365)]
        [TestCase("Formated text", 1753)]
        [TestCase("Images", 1873)]
        [TestCase("Blank", 1207)]
        [TestCase("Complex table", 4201)]
        [TestCase("htmlrtf2", 3639)]
        [TestCase("htmlrtf1", 14679)]
        [TestCase("Three layer nested table", 5789)]
        public void ShouldLoadFromFile(string fileName, int length)
        {
            var file = Path.Combine(TestContext.CurrentContext.TestDirectory, "Resources", fileName + ".rtf");

            var doc = new RTFDomDocument();
            doc.Load(file);
            var text = doc.ToDomString();
            Assert.AreEqual(length, text.Length);

            doc = new RTFDomDocument();
            doc.LoadRTFText(File.ReadAllText(file));
            var rtfText = doc.ToDomString();
            Assert.AreEqual(text, rtfText);
        }
    }
}