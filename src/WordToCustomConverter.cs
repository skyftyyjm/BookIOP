using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Drawing.Pictures;

using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using System.Xml.Linq;
namespace ioh
{
    public class WordToCustomConverter
    {
        public static void SaveCustomToWord(CustomDocument cd, string docxPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document))
            {
                // 创建主文档部分
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                mainPart.Document.Body  = new Body();

                // 保存段落
                foreach (var paragraph in cd.Paragraphs)
                {
                    Paragraph para = new Paragraph(new Run(new Text(paragraph.InputParagraphText)));
                    mainPart.Document.Body.Append(para);
                }

                // 保存表格
                foreach (var table in cd.Tables)
                {
                    DocumentFormat.OpenXml.Wordprocessing.Table tbl = new DocumentFormat.OpenXml.Wordprocessing.Table();
                    //foreach (var row in table)
                    //{
                    //    TableRow tableRow = new TableRow();
                    //    foreach (var cell in row)
                    //    {
                    //        TableCell tableCell = new TableCell(new Paragraph(new Run(new Text(cell))));
                    //        tableRow.Append(tableCell);
                    //    }
                    //    tbl.Append(tableRow);
                    //}
                    mainPart.Document.Body.Append(tbl);
                }

                // 保存公式
                foreach (var formula in cd.Formulas)
                {
                    // 这里假设有一个方法将公式转换为OfficeMath对象
                    //OfficeMath officeMath = MathConverter.ConvertLatexToOmml(formula.InputText);
                    //body.Append(officeMath);
                }

                // 保存图片
                foreach (var imagePath in cd.images)
                {
                    AddImageToBody(mainPart, imagePath);
                }

                mainPart.Document.Save();
            }
        }

        private static void AddImageToBody(MainDocumentPart mainPart, string imagePath)
        {

            var type = GetImageType(System.IO.Path.GetExtension(imagePath));
            ImagePart imagePart = mainPart.AddImagePart(type);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            AddImageToParagraph(mainPart, mainPart.GetIdOfPart(imagePart));
        }

        private static void AddImageToParagraph(MainDocumentPart mainPart, string relationshipId)
        {
            var element = new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 792000L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = 1U, Name = "New Image" },
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                    { Id = 0U, Name = "Picture" },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip(
                                        new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                            new DocumentFormat.OpenXml.Drawing.BlipExtension()
                                            { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                                        )
                                    )
                                    {
                                        Embed = relationshipId,
                                        CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle()
                                    )
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 990000L, Cy = 792000 }
                                    ),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                    )
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );

    
            mainPart.Document.Body.Append(new Paragraph(new Run(element)));
        }

        public static CustomDocument ParseWordToCustom(string docxPath)
        {
            if (!File.Exists(docxPath))
                throw new FileNotFoundException("找不到指定的 Word 文件：" + docxPath);

            string imageOutputFolder = System.IO.Path.GetTempPath();



            CustomDocument result = new CustomDocument();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
            {
                // 1. 处理段落/表格/公式（只遍历 Body 下的直接子元素；如果段落中嵌表格、嵌图片，也可以递归遍历）
                var body = wordDoc.MainDocumentPart.Document.Body;

                // 将 Body 下面的所有元素按顺序遍历，区分 Paragraph / Table / OfficeMath 等
                foreach (var element in body.Elements())
                {
                    // —— 段落
                    if (element is Paragraph para)
                    {
                        // 获取纯文本（Run + Text 组合）
                        string text = para.InnerText ?? string.Empty;
                        text = text.Trim();
                        // 只要非空就存为一个“段落”
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            ParagraphData paragraphData = new ParagraphData
                            { 
                                InputParagraphText = text,
                                InputTitleText = text,
                                InputParagraphRtfText = text,
                                InputAnnotationText = text,
                                InputAnnotationTextRtf = text,
                            };
                            result.Paragraphs.Add(paragraphData);
                        }
                    }
                    // —— 表格
                    else if (element is DocumentFormat.OpenXml.Wordprocessing.Table tbl)
                    {
                        // 新建一个表格的二维字符串列表
                        List<List<string>> tableData = new List<List<string>>();

                        // TableRow = 行，TableCell = 单元格
                        foreach (var row in tbl.Elements<TableRow>())
                        {
                            List<string> rowData = new List<string>();
                            foreach (var cell in row.Elements<TableCell>()) {
                                // 取单元格的文本（如果单元格内部还有段落，会把所有 Run.Text 都拼在一起）
                                string cellText = cell.InnerText ?? string.Empty;
                                rowData.Add(cellText.Trim());
                            }
                            tableData.Add(rowData);
                        }

                        //var newData = new Table
                        //{
                        //    InputTableTitle = "table",
                        //    InputTableText = WebForm.InputTableText,
                        //    FontSize = "小四",
                        //    FontType = "方正细等线",
                        //    EnglishFontType = "231",
                        //};

                        //result.Tables.Add(tableData);
                    }
                    // —— 公式（OfficeMath）
                    else if (element is OfficeMath officeMath)
                    {
                        // 将整个 <m:oMath> 节点以 Raw XML 存储
                        string ommlXml = officeMath.OuterXml;
                        string mmlXml = MathConverter.ConvertOmmlToMML(ommlXml);
                        string latex = MathConverter.ConvertOmmlToLatex(mmlXml); // 可选：转换为 LaTeX 字符串

                        var newData = new Formula
                        {
                            InputText = latex, // 从表单获取段落文本
                            Number = "",//公式编号
                            InputMathMl = mmlXml,//mathml公式
                            fileImage = "",//图片文件base64
                            ObjectLeftSide = 0.0,
                            FontType = null,
                            FontSize = null,
                            LeftSide = 0,
                            PresegmentSpace = 0,
                            PostsegmentSpace = 0,
                            Latex2MathMLvalue = "2",
                        };
                        result.Formulas.Add(newData);
                    }
                    // 如果 Body 里还有其他 OfficeMath 嵌套（比如在段落里），也可以单独再检索一次所有 Descendants<OfficeMath>。
                }

                // 2. 处理嵌套在段落/表格内部的 OfficeMath（有可能部分公式并不是 Body 直接子元素，而是嵌套在 Run 里）
                //    所以再全局搜索一次 Descendants<OfficeMath>
                IEnumerable<OfficeMath> nestedMaths = body.Descendants<OfficeMath>();
                foreach (var om in nestedMaths)
                {
                    string xml = om.OuterXml;

                    string mmlXml = MathConverter.ConvertOmmlToMML(xml);
                    string latex = MathConverter.ConvertOmmlToLatex(mmlXml); // 可选：转换为 LaTeX 字符串

                    var newData = new Formula
                    {
                        InputText = latex, // 从表单获取段落文本
                        Number ="",//公式编号
                        InputMathMl = mmlXml,//mathml公式
                        fileImage = "",//图片文件base64
                        ObjectLeftSide = 0.0,
                        FontType = null,
                        FontSize = null,
                        LeftSide = 0,
                        PresegmentSpace = 0,
                        PostsegmentSpace = 0,
                        Latex2MathMLvalue = "2",
                    };
                    result.Formulas.Add(newData);
                }

                // 3. 处理图片
                // Word 中的图片通常是一个 Drawing 或者一个 Picture（旧版），但最终都对应到 MainDocumentPart.ImageParts
                int imageIndex = 1;
                foreach (ImagePart imgPart in wordDoc.MainDocumentPart.ImageParts)
                {
                    // 生成一个唯一文件名，比如 img1.png、img2.jpeg 等
                    string extension = GetImageExtension(imgPart.ContentType);
                    string fileName = $"img_{imageIndex}{extension}";
                    string fullPath = System.IO.Path.Combine(imageOutputFolder, fileName);

                    // 将图片流写到磁盘
                    using (var stream = imgPart.GetStream())
                    using (var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fs);
                    }

                    // 在 CustomDocument 只存文件名（或相对路径）。注意：若 Word 里同一张图片插入了多次，实际上 ImageParts 里只会有一次，
                    // 但只要不重复导出就可以；若你想保留“多次引用”的信息，需要再扫描文档中每个 Drawing 里引用的 RelationshipId
                    result.images.Add(fullPath);
                    imageIndex++;
                }
            }

            return result;
        }

        /// <summary>
        /// 根据 ContentType 返回合适的文件扩展名（.png/.jpeg/.gif/.bmp 等）
        /// </summary>
        private static string GetImageExtension(string contentType)
        {
            switch (contentType.ToLower())
            {
                case "image/png": return ".png";
                case "image/jpeg": return ".jpg";
                case "image/jpg": return ".jpg";
                case "image/gif": return ".gif";
                case "image/bmp": return ".bmp";
                case "image/tiff": return ".tiff";
                default: return ".img";
            }
        }

        private static PartTypeInfo GetImageType(string contentType)
        {
            switch (contentType.ToLower())
            {
                case ".png": return ImagePartType.Png;
                case ".jpg": return ImagePartType.Jpeg;
                case ".gif": return ImagePartType.Gif;
                case ".bmp": return ImagePartType.Bmp;
                case ".tiff": return ImagePartType.Tiff;
                default: return ImagePartType.Svg;
            }
        }
    }

}
