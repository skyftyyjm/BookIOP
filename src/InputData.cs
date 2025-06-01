using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Vml.Office;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ioh
{
   

    public interface InputData
    {
        int Id { get; set; }
        int SortId { get; set; }
        string Types { get; set; }
    }

    public class LatexData
    {
        public int id { get; set; }
        public string LatexCode { get; set; }
        public string MathMlCode { get; set; }
        public string baseImage { get; set; }
        public int ImagePosition { get; set; }
        public int ImageWidth { get; set; } // 新增宽度属性
        public int ImageHeight { get; set; } // 新增高度属性
    }
    public class TitleData : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "标题";
        public string InputTitleClass { get; set; }//标题类别
        public int InputTitleId { get; set; }//标题ID
        public string SelectText { get; set; } //标题纯文本
        public string InputTitleStyle { get; set; }//标题名称
        public int InputTitleStyleId { get; set; }//标题名称Id
        public double InputTitleRetract { get; set; }//标题缩进，标题
        public double InputTitleAfterInterval { get; set; }//标题序号后间距，标题
        public string InputFontColorR { get; set; }//字体颜色
        public string InputFontColorG { get; set; }//字体颜色
        public string InputFontColorB { get; set; }//字体颜色

        public string TitleNumberColorR { get; set; }//序号颜色
        public string TitleNumberColorG { get; set; }//序号颜色
        public string TitleNumberColorB { get; set; }//序号颜色
        public string Bold { get; set; }//加粗
        public string Italic { get; set; }//斜体
                                          //上间距
        public string MarginTop { get; set; }
        //下间距
        public string MarginBottom { get; set; }
        //标题前间距
        public double BeforeNumericUpDown { get; set; }
        //标题后间距
        public double AfterNumericUpDown { get; set; }
        public string Underline { get; set; }//下划线
                                             //序号后是否换行
        public string NumberLineFeed { get; set; }
        public string InputTitleText { get; set; }//标题文本
        public string InputTitleRtfText { get; set; }//标题文本RTF
        public string Hidden { get; set; }//隐藏，标题
        public string titleCount { get; set; }//标题
        public string equivalentClass { get; set; }//独立标题等同级别
        public string InputTrue { get; set; }//标题
        public string fontType { get; set; }//字体
        public string EnglishFontType { get; set; }//英文字体
        public string fontsize { get; set; }//字体大小
        public string Alignment { get; set; }//对齐方式
    }
    public class ParagraphNotes
    {
        public int Id { get; set; } // 注释ID
        public int TextStart { get; set; }
        public int TextLength { get; set; }
        public string Content { get; set; }

        public string Text { get; set; }
        public int MarkerStart { get; set; }
    }
    public class ParagraphData : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public int ParagraphStyleId { get; set; } //段落样式ID
        public string Types { get; set; } = "段落";

        public int CurrentId { get; set; }
        public Dictionary<int, LatexData> LatexDataMap { get; set; }
        public Dictionary<int, int> ImageIdMap { get; set; }
        public string InputParagraphText { get; set; }
        public string InputTitleText { get; set; }
        public string PagraphType { get; set; }
        public string InputParagraphRtfText { get; set; }
        public List<FontTypes> InputFontTypes { get; set; }
        public static int currentId { get; set; } // 静态变量，初始值为 1
        public string FontType { get; set; }//中文字体
        public string EnglishFontType { get; set; }//英文字体

        public string FontSize { get; set; }//字体大小

        public string Alignment { get; set; }//对齐方式
        public double LineSpace { get; set; }//行距
        public double PresegmentSpace { get; set; }//段前间距
        public double PostsegmentSpace { get; set; }//段后间距
        public double LeftSide { get; set; }//左侧间距

        public string InputAnnotationText { get; set; }//注释文本

        public string InputAnnotationTextRtf { get; set; }//注释文本RTF

        public List<ParagraphNotes> comments { get; set; }//注释列表
        public int CommentId { get; set; }
        public List<ParagraphFormula> formulas { get; set; }//公式列表
        public int FormulaId { get; set; }

    }
    public class specialParagraphData : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "特殊段落";
        public string InputParagraphText { get; set; }
        public string InputTitleText { get; set; }
        public string PagraphType { get; set; }

        public List<FontTypes> InputFontTypes { get; set; }
        public string FontType { get; set; }//中文字体

        public string FontSize { get; set; }//字体大小

        public string Alignment { get; set; }//对齐方式
        public string LineSpace { get; set; }//行距
        public string PresegmentSpace { get; set; }//段前间距
        public string PostsegmentSpace { get; set; }//段后间距
        public string LeftSide { get; set; }//左侧间距
    }

    public class ImageData : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "图片";
        public string InputImageTitle { get; set; }
        public int ImageTitleId { get; set; }
        public int ViceImageTitleStyleId { get; set; }//副图标题样式Id
        public int CaptionStyleId { get; set; }//图注样式Id
        public double ImageInterval { get; set; }//图前间距
        public StyleDetail TitleStyle { get; set; }//标题样式
        public UseParagraphStyle ViceImageTitleStyle { get; set; }//副图标题样式
        public UseParagraphStyle CaptionStyle { get; set; }//图注样式
        public string ImageBitmap { get; set; }//图片位图，插图
        public string ImageVectorgraph { get; set; }//图片矢量图，插图
        public string ImageCount { get; set; }
        public string ImagesName { get; set; }//图片名称
        public List<ImagesUrl> InputImageUrl { get; set; }//图片地址
        public string InputAannotation { get; set; }//图片z注释
        public string InputAannotationRtf { get; set; }//图片注释RTF

    }


    public enum LegendType { Point, Line }
    public enum LineType { Solid, Dashed }
    //public enum LineSize { f025, f050, f075, f100, f125, f150 }
    //public enum LineSize { 2.5f, f050, f075, f100, f125, f150 }

    public enum Arrow { None, oneWay, twoWay }
    public class Legend
    {

    }

    public class ImagesUrl
    {

        public string imageUrl { get; set; }//图片地址
        public string name { get; set; }//图片名称
        public decimal Ratio { get; set; }//图片百分比
    }


    public class Formula : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "公式";

        public string InputText { get; set; }//latex公式
        public string InputMathMl { get; set; }//mathml公式

        public string Number { get; set; }//公式编号
        public string fileImage { get; set; }//图片文件base64
        public string FontType { get; set; }//中文字体

        public string FontSize { get; set; }//字体大小

        public string Latex2MathMLvalue { get; set; }

        public double LeftSide { get; set; }//左侧间距

        public double ObjectLeftSide { get; set; }
        public double PresegmentSpace { get; set; }//段前间距
        public double PostsegmentSpace { get; set; }//段后间距



    }

    public class Video : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "音视频";
        public string InputVideoTitle { get; set; }



        public string InputVideoType { get; set; }
        public string InputVideoUrl { get; set; }

    }


    public class Model : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "模型";
        public string InputModelTitle { get; set; }
        public string InputModelUrl { get; set; }
        public string ModelCount { get; set; }
        public string InputModelImageUrl { get; set; }
        public string InputModelMtlUrl { get; set; }

    }
    public class Table : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "表格";
        public string InputTableTitle { get; set; }
        public string InputTableText { get; set; }
        public string TableCount { get; set; }
        public StyleDetail TitleStyle { get; set; }
        public string FontType { get; set; }//中文字体
        public string EnglishFontType { get; set; }//英文字体
        public string FontSize { get; set; }//字体大小
        public string TableEndLine { get; set; }//表格后边距
        public string MarginLine { get; set; }
        public string PaddingLine { get; set; }


    }
    public class UseExercises
    {
        public int Id { get; set; }
        public string Types { get; set; }//类型
        public string InputExercisesTitle { get; set; }

        public List<Exercise> InputContent { get; set; }
    }
    public class ParagraphFormula
    {
        public int Id { get; set; } // 注释ID
        public int TextStart { get; set; }
        public int TextLength { get; set; }
        public string Latex2MathMLvalue { get; set; }
        public string LatexCode { get; set; }
        public string MathMl { get; set; }

        public string ImageName { get; set; }
        public int MarkerStart { get; set; }
    }
        public class Exercise
    {
        public int ExercisesId { get; set; }
        public string ExercisesType { get; set; }
        public string ExercisesTitle { get; set; }//题目标题
        public string ExercisesTitleRtf { get; set; } // RTF格式的题目标题
        public string ExercisesTitleFormat { get; set; } //解析后的题目标题
        public List<ExerciseOption> ExercisesOption { get; set; }
        public string exercisesExplainTextRtf { get; set; }//解析的RTF格式内容
        public string exercisesExplainTextFormat { get; set; }//解析 后的解析
        public string exercisesExplainText { get; set; }//解析
        public object ExercisesAnswer { get; set; } //答案 可以是 int 

        public Dictionary<int, List<ParagraphFormula>> formulaDict { get; set; }// 用于存储每个RichTextBox对应的公式列表
        public Dictionary<int, int> oldTextLengthDict { get; set; }// 用于存储每个RichTextBox的旧文本长度
        public Dictionary<int, int> oldSelectionDict { get; set; }// 用于存储每个RichTextBox的旧光标位置

    }

    public class ExerciseOption
    {
        public int OptionId { get; set; }
        public string OptionContent { get; set; }//纯文本内容
        public string OptionContentRtf { get; set; } // RTF格式的选项内容
        public string OptionFormatContent { get; set; }//解析后的内容
    }

    public class StyleCategory
    {
        public string Name { get; set; }
        public List<StyleType> Styles { get; set; }
    }
    public class StyleType
    {
        public string Type { get; set; }
        public List<StyleDetail> StyleDetails { get; set; } = new List<StyleDetail>();
        public List<UseParagraphStyle> ParagraphStyle { get; set; } = new List<UseParagraphStyle>();
        public List<UseImageStyle> ImageStyle { get; set; } = new List<UseImageStyle>();//图片样式
        public List<UseExercisesStyle> ExercisesStyle { get; set; } = new List<UseExercisesStyle>();//习题样式

        public List<UseExcelStyle> ExcelStyles { get; set; } = new List<UseExcelStyle>();//表格样式
        public List<UseModelStyle> ModelStyles { get; set; } = new List<UseModelStyle>();//模型样式
        public List<UseVideoStyle> VideoStyles { get; set; } = new List<UseVideoStyle>();//音视频样式
        public List<UseFormulaStyle> FormulaStyles { get; set; } = new List<UseFormulaStyle>();//公式样式
    }

    public class StyleDetail
    {
        public int Id { get; set; }
        public string Name { get; set; }//名称 
        public string FontType { get; set; }//中文字体类型
        public string TitleEnglishFontType { get; set; }//英文字体，标题
        public string fontsize { get; set; }//字体大小，标题
        public string BoldType { get; set; }//加粗，标题
        public string ItalicType { get; set; }//斜体，标题
        public string UnderlineType { get; set; }//下划线，标题
        public string FontColorR { get; set; }//字体颜色R，标题
        public string FontColorG { get; set; }//字体颜色G，标题
        public string FontColorB { get; set; }//字体颜色B，标题
        public string Alignment { get; set; }// 对齐方式，标题
        public double TitleRetract { get; set; }//标题缩进，标题
        public double MarginTop { get; set; }//上间距，标题，
        public double MarginBottom { get; set; }//下间距，标题，
        public string BeforeCharacter { get; set; }//前缀字符，标题
        public string SerialNumber { get; set; }//序号，标题
        public string AfterCharacter { get; set; }//后缀字符，标题
        public double TitleAfterInterval { get; set; }//标题序号后间距，标题
        public string NumberLineFeed { get; set; }//序号后换行,标题
        public string SerialNumberColorR { get; set; }//序号颜色，标题
        public string SerialNumberColorG { get; set; }//序号颜色，标题
        public string SerialNumberColorB { get; set; }//序号颜色，标题
        public string LineFeed { get; set; }//标题后换行,标题

        public string Scope { get; set; }//作用域，标题
        public string Rank { get; set; }//等同级别，标题
        public string equivalentClass { get; set; }   //独立标题等同级别，标题
        public string Hidden { get; set; }//隐藏，标题


    }
    public class UseParagraphStyle//段落样式
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string ChineseFontType { get; set; } // 中文字体
        public string FontType { get; set; }//西文字体
        public string FontSize { get; set; }//字体大小
        public string Aligning { get; set; }//对齐方式
        public double LeftSide { get; set; }//左侧间距
        public double LineSpace { get; set; }//行距
        public double PresegmentSpace { get; set; }//段前间距
        public double PostsegmentSpace { get; set; }//段后间距
    }
    public class UseImageStyle//图片样式
    {
        public int ImageTitle { get; set; }//图片标题id
        public int ViceImageTitleStyle { get; set; }//图片副标题id
        public int CaptionStyle { get; set; }//图注id
        public double ImageInterval { get; set; }//图片间距，插图
        public string ImageBitmap { get; set; }//图片位图，插图
        public string ImageVectorgraph { get; set; }//图片矢量图，插图
    }
    public class UseExercisesStyle//习题样式
    {
        public int TitleStyleId { get; set; }//标题样式ID
        public int QuestionStemStyleId { get; set; }//题干样式段落Id
        public int OptionStyleId { get; set; }    //选项样式段落Id
    }


    public class UseExcelStyle//表格样式
    {
        public int TitleStyleId { get; set; }//标题样式ID
        public string ChineseFontType { get; set; }//中文字体
        public string EnglishFontType { get; set; } //英文字体
        public string FontSize { get; set; } //字体大小
        public double EndLineSpace { get; set; } //表后行间距
        public double OutSideLine { get; set; } //表外框线
        public double InsideLine { get; set; } //表内框线
    }
    public class UseModelStyle
    {
        public int TitleStyleId { get; set; }//标题样式ID
    }
    public class UseVideoStyle
    {
        public string VideoStlyleName { get; set; }//视频样式名称
        public string VideoMp3Name { get; set; }//音频样式名称
    }
    public class UseFormulaStyle
    {
        public string FontType { get; set; }//字体类型
        public string FontSize { get; set; }//字体大小

        public double MarginLeft { get; set; }//左缩进

        public double PresegmentSpace { get; set; }
        public double PostsegmentSpace { get; set; }//段后下间距

    }
    public class ExercisesStyle : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "习题";
        public string InputExercisesTitle { get; set; }//习题标题
        public string ExercisesCount { get; set; }//标题序号

        public List<Exercise> InputContent { get; set; }
        public string fontType { get; set; }

        public int QuestionStemStyleId { get; set; }//题干样式段落Id
        public int OptionStyleId { get; set; }    //选项样式段落Id
        public int TitleStyleId { get; set; } // 标题样式Id

        public StyleDetail TitleStyle { get; set; }//标题样式
        public UseParagraphStyle QuestionStemStyle { get; set; }

        public UseParagraphStyle OptionStyle { get; set; }

    }

    public class FontTypes
    {
        public int Id { get; set; }
        public string name { get; set; }
        public string FontType { get; set; }
    }


    public class QuessionTitle : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "习题标题";
        public string QuessionTitleClass { get; set; }//习题标题等级
        public string QuessionTitleText { get; set; }//习题标题
        public string titleCount { get; set; }//习题标题序号
    }
    public class QuessionChoose : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "习题单选";
        public string ExercisesTitle { get; set; }//题目
        public string ExercisesTitleRtf { get; set; } // RTF格式的题目标题
        public string ExercisesTitleFormat { get; set; } //解析后的题目标题
        public List<ExerciseOption> ExercisesOption { get; set; }//选项
        public string exercisesExplainTextRtf { get; set; }//解析
        public string exercisesExplainTextFormat { get; set; }//解析 后的解析
        public string exercisesExplainText { get; set; }//解析
        public object ExercisesAnswer { get; set; } // 可以是 int 或 List<int> 或 string

        public Dictionary<int, List<ParagraphFormula>> formulaDict { get; set; }// 用于存储每个RichTextBox对应的公式列表
        public Dictionary<int, int> oldTextLengthDict { get; set; }// 用于存储每个RichTextBox的旧文本长度
        public Dictionary<int, int> oldSelectionDict { get; set; }// 用于存储每个RichTextBox的旧光标位置
    }
    public class QuessionMultipleChoice : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "习题多选";
        public string ExercisesTitle { get; set; }//题目
        public string ExercisesTitleRtf { get; set; } // RTF格式的题目标题
        public string ExercisesTitleFormat { get; set; } //解析后的题目标题
        public List<ExerciseOption> ExercisesOption { get; set; }//选项
        public string exercisesExplainTextRtf { get; set; }//解析
        public string exercisesExplainTextFormat { get; set; }//解析 后的解析
        public string exercisesExplainText { get; set; }//解析
        public object ExercisesAnswer { get; set; } // 可以是 int 或 List<int> 或 string
        /*  public List<Exercise> InputContent { get; set; }*/

        public Dictionary<int, List<ParagraphFormula>> formulaDict { get; set; }// 用于存储每个RichTextBox对应的公式列表
        public Dictionary<int, int> oldTextLengthDict { get; set; }// 用于存储每个RichTextBox的旧文本长度
        public Dictionary<int, int> oldSelectionDict { get; set; }// 用于存储每个RichTextBox的旧光标位置

    }

    public class QuessionSubjectivity : InputData
    {
        public int Id { get; set; }
        public int SortId { get; set; }
        public string Types { get; set; } = "习题主观";
        /* public List<Exercise> InputContent { get; set; }*/
        public string ExercisesTitle { get; set; }//题目
        public string ExercisesTitleRtf { get; set; } // RTF格式的题目标题
        public string ExercisesTitleFormat { get; set; } //解析后的题目标题
        public List<ExerciseOption> ExercisesOption { get; set; }//选项
        public string exercisesExplainTextRtf { get; set; }//解析
        public string exercisesExplainTextFormat { get; set; }//解析 后的解析
        public string exercisesExplainText { get; set; }//解析
        public object ExercisesAnswer { get; set; } // 可以是 int 或 List<int> 或 string

        public Dictionary<int, List<ParagraphFormula>> formulaDict { get; set; }// 用于存储每个RichTextBox对应的公式列表
        public Dictionary<int, int> oldTextLengthDict { get; set; }// 用于存储每个RichTextBox的旧文本长度
        public Dictionary<int, int> oldSelectionDict { get; set; }// 用于存储每个RichTextBox的旧光标位置
    }



}
