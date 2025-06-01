using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ioh
{


    public class CustomDocument
    {
        /// <summary>
        /// 文档里的所有段落，每个元素对应一个段落的纯文本内容
        /// </summary>
        public List<ParagraphData> Paragraphs { get; set; } = new List<ParagraphData>();

        /// <summary>
        /// 文档里的所有图片，使用文件名或相对路径来标识。图片二进制数据你可以另外序列化，也可以直接保存到磁盘，从 JSON 里只存路径/文件名。
        /// </summary>
        public List<string> images { get; set; } = new List<string>();

        /// <summary>
        /// 所有表格。每个表格由若干行（List&lt;List&lt;string&gt;&gt;）组成，行里的每个单元格以纯文本形式存储。
        /// </summary>
        public List<Table> Tables { get; set; } = new List<Table>();

        /// <summary>
        /// 所有公式。示例这里把公式以字符串形式存到列表里，实际保存时可以是 OMML（Office Math Markup Language）的 XML 字符串。
        /// </summary>
        public List<Formula> Formulas { get; set; } = new List<Formula>();
    }
}
