using System;
using System.Collections.Generic;
using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace Uzzal.ExcelMapper.Test
{
    public class ExcelMapTest
    {
        private readonly ITestOutputHelper output;
        private readonly string xlsFile;
        private readonly string xlsxFile;

        public ExcelMapTest(ITestOutputHelper output)
        {
            this.output = output;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            xlsFile = @"resources\\Sample.xls";
            xlsxFile = @"resources\\Sample.xlsx";

            output.WriteLine(xlsFile);
        }

        [Fact]
        public void MapTest()
        {
            var reader = new ExcelService(xlsFile);
            var service = new ExcelMap<SampleModel>(reader);
            var map = service.Map();

            Assert.Collection(map, 
                x =>
                {
                    Assert.Contains("MYMENSINGH", x.Region);
                    Assert.Contains("MYMENSINGH", x.Area);
                    Assert.Contains("13", x.UserId);
                    Assert.Contains("MR.MD. SHAKHWAT HOSSAIN", x.Name);
                    Assert.Contains("RSM", x.Designation);
                },
                x =>
                {
                    Assert.Contains("MYMENSINGH", x.Region);
                    Assert.Contains("MYMENSINGH", x.Area);
                    Assert.Contains("19", x.UserId);
                    Assert.Contains("MR.REZAUL KARIM", x.Name);
                    Assert.Contains("DSM", x.Designation);
                },
                x =>
                {
                    Assert.Contains("MYMENSINGH", x.Region);
                    Assert.Contains("MYMENSINGH", x.Area);
                    Assert.Contains("930", x.UserId);
                    Assert.Contains("MR. ZINNATH ALI", x.Name);
                    Assert.Contains("PSO", x.Designation);
                }

            );
        }
    }

    class SampleModel
    {
        public string Region { get; set; }
        public string Area { get; set; }
        
        [ExcelColumn("UserID")]
        public string UserId { get; set; }
        public string Name { get; set; }

        [ExcelColumn("DESIGNATION")]
        public string Designation { get; set; }
    }
}
