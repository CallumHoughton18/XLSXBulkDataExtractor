using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.WPFLogic.Models;

namespace XLSXBulkDataExtractor.WPFLogicTests
{
    public class DataRetrievalRequestTests
    {
        [Test]
        public void ColumnNameConversionValidLetterTest()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "Z";

            Assert.That(sut.ColumnNumber, Is.EqualTo(26));
        }

        [Test]
        public void ColumnNameConversionValidLowerCaseTest()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "c";

            Assert.That(sut.ColumnNumber, Is.EqualTo(3));
        }

        [Test]
        public void ColumnNameConversionValidLetterTest2()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "ZM";

            Assert.That(sut.ColumnNumber, Is.EqualTo(689));
        }

        [Test]
        public void ColumnNameConversionValidUpperAndLowerCaseTest()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "zM";

            Assert.That(sut.ColumnNumber, Is.EqualTo(689));
        }

        [Test]
        public void ColumnNameConversionValidLetterTest3()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "BB";

            Assert.That(sut.ColumnNumber, Is.EqualTo(54));
        }

        [Test]
        public void ColumnNameConversionValidLetterTest4()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "ZZ";

            Assert.That(sut.ColumnNumber, Is.EqualTo(702));
        }

        [Test]
        public void ColumnNameConversionValidLetterTest5()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "Z";

            Assert.That(sut.ColumnNumber, Is.EqualTo(26));
        }

        [Test]
        public void ColumnNameConversionInvalidStringTest()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "!TEST!FALSE!";

            Assert.That(sut.ColumnNumber, Is.EqualTo(1)); //ie hasn't changed
        }

        [Test]
        public void ColumnNameConversionValidNumberTest()
        {
            var sut = new DataRetrievalRequest();
            sut.ColumnName = "23";

            Assert.That(sut.ColumnNumber, Is.EqualTo(23));
        }
    }
}
