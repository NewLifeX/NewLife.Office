using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>PDF 增强测试 — 注释类型、表单填充读取</summary>
public class PdfEnhancementTests
{
    [Fact(DisplayName = "PDF—Caret注释写入")]
    public void Annotation_Caret()
    {
        var path = Path.GetTempFileName() + ".pdf";
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Caret, PageIndex = 0,
                X = 100, Y = 700, Width = 10, Height = 15,
                Contents = "插入此处", Author = "审阅者"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);
            Assert.True(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—Polygon注释写入")]
    public void Annotation_Polygon()
    {
        var path = Path.GetTempFileName() + ".pdf";
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Polygon, PageIndex = 0,
                X = 100, Y = 700, Width = 50, Height = 50, Contents = "多边形区域"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);
            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /Polygon", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—Squiggly注释写入")]
    public void Annotation_Squiggly()
    {
        var path = Path.GetTempFileName() + ".pdf";
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Squiggly, PageIndex = 0,
                X = 100, Y = 700, Width = 80, Height = 5, Contents = "波浪下划线"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);
            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /Squiggly", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—表单创建+填充+读取往返")]
    public void FormField_FillAndRead()
    {
        var path = Path.GetTempFileName() + ".pdf";
        try
        {
            using var writer = new PdfWriter();
            writer.AddTextField("txtName", 100, 700, 200, 20, "张三");
            writer.AddCheckBox("chkAgree", 100, 650, 12, true);
            var options = new List<String> { "北京", "上海", "广州" };
            writer.AddComboBox("selCity", 100, 620, 200, 20, options, 1);
            writer.Save(path);

            using var reader = new PdfReader(path);
            var form = reader.ReadFormFields();
            Assert.NotNull(form);
            Assert.True(form!.Fields.Count >= 3);
            var nameField = form.Fields.FirstOrDefault(f => f.FullName == "txtName");
            Assert.NotNull(nameField);
            Assert.Equal("张三", nameField!.Value);
            var selField = form.Fields.FirstOrDefault(f => f.FullName == "selCity");
            Assert.NotNull(selField);
            Assert.Contains("上海", selField!.Options);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—SetFormFieldValue修改值")]
    public void FormField_SetValue()
    {
        var path = Path.GetTempFileName() + ".pdf";
        try
        {
            using var writer = new PdfWriter();
            writer.AddTextField("txtName", 100, 700, 200, 20, "初始值");
            writer.SetFormFieldValue("txtName", "修改后");
            writer.Save(path);

            using var reader = new PdfReader(path);
            var form = reader.ReadFormFields();
            Assert.NotNull(form);
            var nameField = form!.Fields.FirstOrDefault(f => f.FullName == "txtName");
            Assert.NotNull(nameField);
            Assert.Equal("修改后", nameField!.Value);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
