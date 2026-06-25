using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using NewLife;
using NewLife.Office;
using NewLife.Log;
using NewLife.Reflection;

namespace Test;

class Program
{
    static void Main(String[] args)
    {
        Runtime.CreateConfigOnMissing = false;
        XTrace.UseConsole();

        try
        {
            TestPptxMaster();
        }
        catch (Exception ex)
        {
            XTrace.WriteException(ex);
        }

        Console.WriteLine("OK!");
        Console.ReadKey();
    }

    /// <summary>演示 PPT 母版页功能：从企业模板加载母版/版式/主题后添加幻灯片</summary>
    static void TestPptxMaster()
    {
        // 1. 先创建一个带有多版式的模板 pptx（模拟企业品牌模板）
        var templatePath = "template_demo.pptx".GetFullPath();
        using (var templateWriter = new PptxWriter())
        {
            // 默认使用硬编码母版，无模板加载时就是简单空白版式
            templateWriter.AddSlide();
            templateWriter.AddTextBox(0, "模板占位 - 此页在加载时会被丢弃", 2, 5, 20, 2);
            templateWriter.Save(templatePath);
        }
        Console.WriteLine($"模板文件已生成: {templatePath}");
        Console.WriteLine();

        // 2. 从模板创建新演示文稿，复用其母版/版式/主题
        var outputPath = "output_master_demo.pptx".GetFullPath();
        using (var writer = new PptxWriter())
        {
            // 从模板加载母版基础设施
            writer.LoadMaster(templatePath);

            // 查看可用版式
            Console.WriteLine($"版式数量: {writer.GetLayoutCount()}");
            for (var i = 0; i < writer.GetLayoutCount(); i++)
            {
                Console.WriteLine($"  版式 {i}: {writer.GetLayoutName(i)}");
            }
            Console.WriteLine();

            // 使用版式 0 添加标题幻灯片
            var titleSlide = writer.AddSlide(0);
            writer.AddTextBox(0, "企业季度报告", 2, 3, 22, 3, fontSize: 36, bold: true);
            writer.AddTextBox(0, "2026 Q2", 2, 8, 10, 2, fontSize: 18);
            Console.WriteLine("已添加标题幻灯片");

            // 使用版式 0 添加内容幻灯片
            var contentSlide = writer.AddSlide(0);
            writer.AddTextBox(1, "业绩概览", 2, 1, 20, 2, fontSize: 28, bold: true);
            writer.AddTextBox(1, "营收增长 25%，利润率提升 5 个百分点", 2, 4, 20, 4, fontSize: 18);
            Console.WriteLine("已添加内容幻灯片");

            // 也可以编程式自定义幻灯片背景
            writer.SetBackground(1, "F2F2F2");

            writer.Save(outputPath);
        }
        Console.WriteLine();
        Console.WriteLine($"输出文件已生成: {outputPath}");

        // 3. 用读取器验证
        using var reader = new PptxReader(outputPath);
        Console.WriteLine();
        Console.WriteLine($"生成文件包含 {reader.GetSlideCount()} 张幻灯片:");
        for (var i = 0; i < reader.GetSlideCount(); i++)
        {
            Console.WriteLine($"--- 幻灯片 {i + 1} ---");
            var text = reader.GetSlideText(i);
            Console.WriteLine(text.Length > 200 ? text[..200] + "..." : text);
        }
    }
}