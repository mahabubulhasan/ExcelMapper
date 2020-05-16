using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace Uzzal.ExcelMapper.Test
{
    public class ExcelServiceTest
    {
        private readonly ITestOutputHelper output;
        private readonly string xlsFile;
        private readonly string xlsxFile;
        private readonly ExcelService reader;

        public ExcelServiceTest(ITestOutputHelper output)
        {
            this.output = output;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            xlsFile = Path.GetFullPath(@"resources\\Sample.xls");
            xlsxFile = Path.GetFullPath(@"resources\\Sample.xlsx");

            output.WriteLine(xlsFile);
            reader = new ExcelService(xlsFile);
        }

        [Fact]
        public void HasColumnTest()
        {
            Assert.True(reader.HasColumns(new string[] { "Region", "Area", "UserID" }));
            Assert.False(reader.HasColumns(new string[] { "RegionNo", "AreaNo", "UserIDNo" }));
        }

        [Fact]
        public void Sample()
        {
            Assert.True(true);
        }
    }
}
