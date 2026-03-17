using System.Text;

namespace XUnitTest.Integration;

/// <summary>集成测试基类，提供共享配置</summary>
public abstract class IntegrationTestBase
{
    /// <summary>输出文件目录</summary>
    protected static readonly String OutputDir = "./files".GetFullPath();

    static IntegrationTestBase()
    {
        // 注册 CodePages 编码提供程序，PdfReader 需要 1252 编码
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Directory.CreateDirectory(OutputDir);
    }
}
