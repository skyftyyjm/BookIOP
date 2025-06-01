using System;
using System.IO;
using System.IO.Compression;
using ioh;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;

class Program
{
    static void convertToZip(string filePath, string zipFilePath)
    {
        string workPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(workPath);
        string imagePath = Path.Combine(workPath, "images");
        Directory.CreateDirectory(imagePath);

        string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
        var mainJsonData = new
        {
            ProjectName = fileName,
            ProjectType = "图书",
            Template = fileName
        };
        // 将 main.json 数据序列化为 JSON
        string mainJson = JsonConvert.SerializeObject(mainJsonData, Formatting.Indented);
        string mainJsonPath = Path.Combine(workPath, "main.json");
        File.WriteAllText(mainJsonPath, mainJson);
        var mapJsonData = new
        {
            graphNodeId = 0,
            cusId = 0,
            cusIndexId = 0,
            legends = new List<string>(),
            nodes = new List<string>(),
            relations = new List<string>(),
        };
        // 将 main.json 数据序列化为 JSON
        string mapJson = JsonConvert.SerializeObject(mapJsonData, Formatting.Indented);

        string questionJsonPath = Path.Combine(workPath, "question.json");
        string mapJsonPath = Path.Combine(workPath, "map.json");
        File.WriteAllText(mapJsonPath, mapJson);
        File.WriteAllText(questionJsonPath, JsonConvert.SerializeObject(TemplateStyleData.Styles, Formatting.Indented)); // 写入json

        List<InputData> InputDataList = new List<InputData>();

        int myId = 1;

        var titleData = new TitleData
        {
            Id = myId, // 设置新 ID
            SortId = myId,
            InputTitleId = 0,
            Types = "标题",
            InputTitleClass = "结构标题", // 标题类型
            InputTitleStyle = "一级", // 标题样式名称
            SelectText = "标题1",
            InputTitleText = "标题1", // 标题内容
            InputTitleRtfText = "标题1",//标题Rtf内容
            BeforeNumericUpDown = 0,//标题前间距
            AfterNumericUpDown = 0,//标题后间距
            InputTitleRetract = 0,//标题缩进
            InputTitleAfterInterval = 0,//标题序号后间距
            fontType = "方正标雅宋",
            EnglishFontType = "方正标雅宋",
            Alignment = "居中",
            fontsize = "八号",
        };
        InputDataList.Add(titleData);
        myId++;

        CustomDocument cd = WordToCustomConverter.ParseWordToCustom(filePath);
        for (int i = 0; i < cd.Paragraphs.Count; i++)
        {
            cd.Paragraphs[i].Id = myId;
            cd.Paragraphs[i].SortId = myId;
            cd.Paragraphs[i].ParagraphStyleId = 1;
            cd.Paragraphs[i].Types = "段落";
            cd.Paragraphs[i].PagraphType = "通用段落";
            cd.Paragraphs[i].CurrentId = 1;
            cd.Paragraphs[i].CommentId = 1;
            cd.Paragraphs[i].CurrentId = 1;
            cd.Paragraphs[i].FormulaId = 1;
            InputDataList.Add(cd.Paragraphs[i]);
            myId++;
        }

        for (int i = 0; i < cd.Formulas.Count; i++)
        {
            cd.Formulas[i].Id = myId;
            cd.Formulas[i].SortId = myId;
            cd.Formulas[i].Types = "公式";
            InputDataList.Add(cd.Formulas[i]);
            myId++;
        }


        for (int i = 0; i < cd.images.Count; i++)
        {
            string imagesName = System.IO.Path.GetFileName(cd.images[i]);
            string imageDestPath = Path.Combine(imagePath, imagesName);

            System.IO.File.Copy(cd.images[i], imageDestPath);
            ImagesUrl imagesUrl = new ImagesUrl
            {
                Ratio = 100,
                imageUrl = imagesName,
            };
            var newData = new ImageData
            {
                Id = myId, // 设置新 ID
                SortId = myId, // 设置新 ID
                Types = "图片",
                InputImageTitle = imagesName,
                ImagesName = imagesName,//图片名称
                InputImageUrl = new List<ImagesUrl> {
                    imagesUrl
                }
            };

            InputDataList.Add(newData);
            myId++;
        }


        for (int i = 0; i < cd.Tables.Count; i++)
        {
            cd.Tables[i].Id = myId;
            cd.Tables[i].SortId = myId;
            cd.Tables[i].Types = "表格";

            InputDataList.Add(cd.Tables[i]);
            myId++;
        }
        string inputDataListJson = JsonConvert.SerializeObject(InputDataList, Formatting.Indented);

        string projectJsonPath = Path.Combine(workPath, "project.json");
        File.WriteAllText(projectJsonPath, inputDataListJson); // 写入json
  
        if (System.IO.File.Exists(zipFilePath))
        {
            System.IO.File.Delete(zipFilePath);
        }
        ZipFile.CreateFromDirectory(workPath, zipFilePath);
    }

    static void convertToDocx(string filePath, string docxFilePath)
    {
        string workPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(workPath);
        ZipFile.ExtractToDirectory(filePath, workPath);

        string mainJsonPath = Path.Combine(workPath, "main.json"); // 读取 main.json 文件
        string questionJsonPath = Path.Combine(workPath, "question.json");
        string mapJsonPath = Path.Combine(workPath, "map.json");
        string projectJsonPath = Path.Combine(workPath, "project.json");

        string mapJsonData = File.ReadAllText(mapJsonPath);
        string mainJsonData = File.ReadAllText(mainJsonPath);

        string questionJsonData = File.ReadAllText(questionJsonPath);
        TemplateStyleData.Styles = JsonConvert.DeserializeObject<List<StyleCategory>>(questionJsonData);

        string projectJsonData = File.ReadAllText(projectJsonPath);
        var inputDataList = JsonConvert.DeserializeObject<List<InputData>>(projectJsonData, new InputDataConverter());
        CustomDocument cd = new CustomDocument
        {
            Paragraphs = new List<ParagraphData>(),
            Formulas = new List<Formula>(),
            images = new List<string>(),
            Tables = new List<Table>()
        };
        foreach (var item in inputDataList)
        {
            if (item is ParagraphData paragraph)
            {
                cd.Paragraphs.Add(paragraph);
            }
            else if (item is Formula formula)
            {
                cd.Formulas.Add(formula);
            }
            else if (item is ImageData imageData)
            {
                string imagePath = Path.Combine(workPath, "images" ,imageData.ImagesName);
                cd.images.Add(imagePath);
            }
            else if (item is Table table)
            {
                cd.Tables.Add(table);
            }
        }
       WordToCustomConverter.SaveCustomToWord(cd, docxFilePath);
    }

    static void Main(string[] args)
    {
        // Generate a random directory name
        string inFilePath;
        string outFilePath;
        bool toZip = false;

        if (args.Length == 1)
        {
            inFilePath = args[0];
            if (inFilePath.EndsWith(".docx"))
            {
                outFilePath = inFilePath + ".zip";
                toZip = true;
            } else if (inFilePath.EndsWith(".zip"))
            {
                outFilePath = inFilePath.Substring(0, inFilePath.Length - 4); // Remove the .zip extension
                if (!outFilePath.EndsWith(".docx"))
                {
                    outFilePath = outFilePath + ".docx";
                }
                toZip = false;
            } else
            {
                return;
            }

        } else if (args.Length == 2)
        {
            inFilePath = args[0];
            if (inFilePath.EndsWith(".docx"))
            {
                toZip = true;
            }
            outFilePath = args[1];
        }
        else
        {
            Console.Write("err");
            return;
        }

        if (toZip)
        {
            convertToZip(inFilePath, outFilePath);


        } else
        {
            convertToDocx(inFilePath, outFilePath);

        }
        Console.Write(outFilePath);




        // Compress the directory into a zip file


    }
}
