using System.Xml.Xsl;
using System.Xml;
using System.Security;

namespace ioh
{

    public static class MathConverter
    {
        // XSLT 文件相对路径（相对于可执行文件目录或项目输出目录）
        // 请按实际情况修改：如果你把 OMML2LaTeX.XSL 放到项目根目录下的 "XSLT" 文件夹内，
        // 那么相对路径可能是 "XSLT\\OMML2LaTeX.XSL"。
        private static readonly string Omml2LatexMllPath = Path.Combine(AppContext.BaseDirectory, "xsl_yarosh", "OMML2MML.XSL");
        private static readonly string Omml2LatexXsltPath = Path.Combine(AppContext.BaseDirectory, "xsl_yarosh", "mmltex.xsl");

        public static string ConvertOmmlToMML(string ommlXml)
        {
            if (string.IsNullOrWhiteSpace(ommlXml))
                throw new ArgumentException("传入的 OMML XML 不能为空。", nameof(ommlXml));

            if (!File.Exists(Omml2LatexMllPath))
                throw new FileNotFoundException("找不到 OMML → LaTeX 的 XSLT 文件，请确认路径是否正确： " + Omml2LatexMllPath);

            try
            {
                // 1. 准备一个 XmlReader 从 ommlXml 中读取 OMML 片段
                using (StringReader stringReader = new StringReader(ommlXml))
                using (XmlReader xmlReader = XmlReader.Create(stringReader))
                {
                    // 2. 准备一个 XslCompiledTransform，并载入 XSLT
                    XslCompiledTransform xslt = new XslCompiledTransform();

                    // XsltSettings(trustedXsltScript = true) 允许在 XSLT 中使用脚本
                    var xsltSettings = new XsltSettings(enableDocumentFunction: false, enableScript: false);
                    // 如果 XSLT 里自带 <xsl:import ...> 或 <xsl:include ...>，需要在第四个参数里指定 XmlResolver
                    // 这里设为 null，禁止任何外部 DTD/实体解析
                    xslt.Load(Omml2LatexMllPath, xsltSettings, new XmlUrlResolver());

                    // 3. 准备一个 StringWriter 来接收转换结果
                    using (StringWriter sw = new StringWriter())
                    using (XmlWriter xmlWriter = XmlWriter.Create(sw, xslt.OutputSettings))
                    {
                        // 4. 执行转换：输入是 xmlReader，输出写到 xmlWriter（最终就是 StringWriter）
                        xslt.Transform(xmlReader, xmlWriter);
                        // 5. 获取转换后的字符串（可能包含换行、缩进；我们去除前后多余空白）
                        string latexResult = sw.ToString().Trim();
                        return latexResult;
                    }
                }
            }
            catch (XsltException xe)
            {
                throw new InvalidOperationException("在执行 XSLT 转换时出错，请检查 OMML 和 XSLT 文件的兼容性。", xe);
            }
        }
        /// <summary>
        /// 将一个 OMML（OfficeMath）的 XML 字符串，转换成对应的 LaTeX 字符串。
        /// </summary>
        /// <param name="ommlXml">OMML 节点（例如："<m:oMath xmlns:m='...'>…</m:oMath>"）</param>
        /// <returns>转换后的 LaTeX 文本（不带行首/行尾空白）。</returns>
        public static string ConvertOmmlToLatex(string ommlXml)
        {
            if (string.IsNullOrWhiteSpace(ommlXml))
                throw new ArgumentException("传入的 OMML XML 不能为空。", nameof(ommlXml));

            if (!File.Exists(Omml2LatexXsltPath))
                throw new FileNotFoundException("找不到 OMML → LaTeX 的 XSLT 文件，请确认路径是否正确： " + Omml2LatexXsltPath);

            try
            {
                // 1. 准备一个 XmlReader 从 ommlXml 中读取 OMML 片段
                using (StringReader stringReader = new StringReader(ommlXml))
                using (XmlReader xmlReader = XmlReader.Create(stringReader))
                {
                    // 2. 准备一个 XslCompiledTransform，并载入 XSLT
                    XslCompiledTransform xslt = new XslCompiledTransform();

                    // XsltSettings(trustedXsltScript = true) 允许在 XSLT 中使用脚本
                    var xsltSettings = new XsltSettings(enableDocumentFunction: true, enableScript: false);
                    // 如果 XSLT 里自带 <xsl:import ...> 或 <xsl:include ...>，需要在第四个参数里指定 XmlResolver
                    // 这里设为 null，禁止任何外部 DTD/实体解析
                    xslt.Load(Omml2LatexXsltPath, xsltSettings, new XmlUrlResolver());

                    // 3. 准备一个 StringWriter 来接收转换结果
                    using (StringWriter sw = new StringWriter())
                    using (XmlWriter xmlWriter = XmlWriter.Create(sw, xslt.OutputSettings))
                    {
                        // 4. 执行转换：输入是 xmlReader，输出写到 xmlWriter（最终就是 StringWriter）
                        xslt.Transform(xmlReader, xmlWriter);
                        // 5. 获取转换后的字符串（可能包含换行、缩进；我们去除前后多余空白）
                        string latexResult = sw.ToString().Trim();
                        return latexResult;
                    }
                }
            }
            catch (XsltException xe)
            {
                throw new InvalidOperationException("在执行 XSLT 转换时出错，请检查 OMML 和 XSLT 文件的兼容性。", xe);
            }
        }

        public static string ConvertLatexToOmml(string latex)
        {
            try
            {
                // 1. 准备一个 XmlReader 从 ommlXml 中读取 OMML 片段

                using (StringReader stringReader = new StringReader($"<formula><latex>{SecurityElement.Escape(latex)}</latex></formula>"))
                using (XmlReader xmlReader = XmlReader.Create(stringReader))
                {
                    // 2. 准备一个 XslCompiledTransform，并载入 XSLT
                    XslCompiledTransform xslt = new XslCompiledTransform();

                    // XsltSettings(trustedXsltScript = true) 允许在 XSLT 中使用脚本
                    var xsltSettings = new XsltSettings(enableDocumentFunction: true, enableScript: false);
                    // 如果 XSLT 里自带 <xsl:import ...> 或 <xsl:include ...>，需要在第四个参数里指定 XmlResolver
                    // 这里设为 null，禁止任何外部 DTD/实体解析
                    xslt.Load(Omml2LatexXsltPath, xsltSettings, new XmlUrlResolver());
                    // 3. 准备一个 StringWriter 来接收转换结果
                    using (StringWriter sw = new StringWriter())
                    using (XmlWriter xmlWriter = XmlWriter.Create(sw, xslt.OutputSettings))
                    {
                        // 4. 执行转换：输入是 xmlReader，输出写到 xmlWriter（最终就是 StringWriter）
                        xslt.Transform(xmlReader, xmlWriter);
                        // 5. 获取转换后的字符串（可能包含换行、缩进；我们去除前后多余空白）
                        string latexResult = sw.ToString().Trim();
                        return latexResult;
                    }
                }
            }
            catch (XsltException xe)
            {
                throw new InvalidOperationException("在执行 XSLT 转换时出错，请检查 OMML 和 XSLT 文件的兼容性。", xe);
            }
        }
    }
}
