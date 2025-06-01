using ioh;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ioh
{
    public static class TemplateStyleData
    {
        public static List<StyleCategory> Styles { get; set; } = new List<StyleCategory>
        {
        new StyleCategory
        {
            Name = "标题样式",
            Styles = new List<StyleType>
            {
                new StyleType
                {
                    Type = "结构标题",
                    StyleDetails = new List<StyleDetail>
                    {
                        new StyleDetail {
                            Id = 1,
                            Name = "一级",
                            FontType = "方正标雅宋",
                            fontsize = "小二",
                            FontColorR ="0",
                            FontColorG ="0",
                            FontColorB ="0",
                            BeforeCharacter = "第",
                            SerialNumber ="一二三",
                            AfterCharacter ="章",
                            Alignment="居中",
                            MarginTop =0,//段前上间距
                            MarginBottom = 0,//段后下间距
                            BoldType="是",
                            UnderlineType="否",
                            SerialNumberColorR = "0",//序号颜色
                            SerialNumberColorG = "0",
                            SerialNumberColorB = "0",
                            LineFeed = "否",
                            NumberLineFeed = "否",
                            Scope = "否"
                        },
                        new StyleDetail {
                             Id = 2,
                            Name = "二级",
                            FontType = "方正标雅宋",
                            fontsize = "小三",
                             FontColorR ="0",
                            FontColorG ="0",
                            FontColorB ="0",
                            BeforeCharacter = "第",
                            SerialNumber ="一二三",
                            AfterCharacter ="节",
                            Alignment="居中",
                            MarginTop =0,//段前上间距
                            MarginBottom = 0,//段后下间距
                            //Effectiveness="是",
                            ItalicType="否",
                            BoldType="是",UnderlineType="否",
                            SerialNumberColorR = "0",//序号颜色
                            SerialNumberColorG = "0",
                            SerialNumberColorB = "0",
                            LineFeed = "否",
                            NumberLineFeed = "否",Scope = "否"
                        },
                        new StyleDetail {
                                  Id = 3,
                            Name = "三级",
                            FontType = "方正标雅宋",
                            fontsize = "小三" ,
                            FontColorR ="0",
                            FontColorG ="0",
                            FontColorB ="0",
                            BeforeCharacter = "",
                            SerialNumber ="一、二、三、",
                            AfterCharacter ="",
                            Alignment="左对齐",
                             MarginTop =0,//段前上间距
                            MarginBottom = 0,//段后下间距
                            //Effectiveness="是",
                            ItalicType="否",
                            BoldType="是",UnderlineType="否",
                            SerialNumberColorR = "0",
                            SerialNumberColorG = "0",
                            SerialNumberColorB = "0",
                            LineFeed = "否",
                            NumberLineFeed = "否",Scope = "否"
                        },
                        new StyleDetail {
                                  Id = 4,
                            Name = "四级", FontType = "方正标雅宋", fontsize = "四号",
                              FontColorR ="0",
                            FontColorG ="0",
                            FontColorB ="0",
                            BeforeCharacter = "", SerialNumber ="(一)(二)(三)",
                            AfterCharacter ="",
                            Alignment="左对齐",
                              MarginTop =0,//段前上间距
                            MarginBottom = 0,//段后下间距
                          //  Effectiveness="是",
                            ItalicType="否",
                            BoldType="是",UnderlineType="否",
                            SerialNumberColorR = "0",
                            SerialNumberColorG = "0",
                            SerialNumberColorB = "0",
                            LineFeed = "否",
                            NumberLineFeed = "否",Scope = "否"
                        },
                        new StyleDetail { Name = "五级", FontType = "方正教材规范楷体", fontsize = "四号",
                           FontColorR ="0",
                            FontColorG ="0",
                            FontColorB ="0",
                            BeforeCharacter = "", SerialNumber ="1.2.3.",
                            AfterCharacter ="",Alignment="左对齐",
                                       MarginTop =0,//段前上间距
                            MarginBottom = 0,//段后下间距
                            //Effectiveness="是",
                            ItalicType="否",
                            BoldType="是",UnderlineType="否",
                            SerialNumberColorR = "0",
                            SerialNumberColorG = "0",
                            SerialNumberColorB = "0",
                            LineFeed = "否",
                            NumberLineFeed = "否",Scope = "否"

                        },
                        new StyleDetail {
                                  Id = 5,
                            Name = "六级", FontType = "方正教材规范楷体", fontsize = "六号",
                                FontColorR ="0",
                            FontColorG ="0",
                            FontColorB ="0",
                            BeforeCharacter = "", SerialNumber ="1.2.3.",
                            AfterCharacter ="",Alignment="左对齐",
                                       MarginTop =40,//段前上间距
                            MarginBottom = 40,//段后下间距
                            //Effectiveness="是",
                            ItalicType="否",
                            BoldType="是",UnderlineType="否",
                            SerialNumberColorR = "0",
                            SerialNumberColorG = "0",
                            SerialNumberColorB = "0",
                            LineFeed = "否",
                            NumberLineFeed = "否",Scope = "否"
                        }
                    }
                },
                new StyleType
                {
                    Type = "顺序标题",
                    StyleDetails = new List<StyleDetail>
                    {
                        new StyleDetail {       Id = 1,Name = "插图", FontType = "方正教材规范楷体", fontsize = "小五",BeforeCharacter = "", SerialNumber ="1.2.3.",
                            AfterCharacter ="",Rank="四级",equivalentClass="四级", Scope = "一级" },
                        new StyleDetail {      Id = 2, Name = "表格", FontType = "方正教材规范楷体", fontsize = "小五",BeforeCharacter = "", SerialNumber ="1.2.3.",
                            AfterCharacter ="",Rank="四级",equivalentClass="四级", Scope = "一级" }
                    }
                },
                new StyleType
                {
                    Type = "独立标题",
                    StyleDetails = new List<StyleDetail>
                    {
                       new StyleDetail {       Id = 1,Name = "前言", FontType = "方正标雅宋", fontsize = "二号",Alignment="居中",Rank="一级",equivalentClass="一级", Scope = "一级"}
                    }
                }
            }
        },
        new StyleCategory { Name = "段落样式", Styles = new List<StyleType>{
               new StyleType{
                     Type = "通用段落",
                      ParagraphStyle = new List<UseParagraphStyle>
                      {
                            new UseParagraphStyle {
                                Id = 1,
                                Name = "默认",
                                ChineseFontType = "方正细等线",
                                FontType="Times",
                                FontSize = "四号",
                                Aligning ="左对齐",
                                LeftSide =2,
                                PresegmentSpace=0,
                                PostsegmentSpace = 0,
                                LineSpace=2
                            },
                      }
               }


        } },
        new StyleCategory { Name = "公式样式", Styles = new List<StyleType>{
               new StyleType{
                     Type = "通用公式",
                      FormulaStyles = new List<UseFormulaStyle>
                      {
                            new UseFormulaStyle {
                            FontType = "方正教材规范楷体",
                            FontSize = "一号",
                            MarginLeft=5,
                            PresegmentSpace=2.5,
                            PostsegmentSpace=2.5,
                            },
                      }
               }
        }

        },
        new StyleCategory {
            Name ="插图样式",
            Styles = new List<StyleType>{
            new StyleType{
                Type = "插图样式",
               ImageStyle = new List< UseImageStyle>
                {
                   new  UseImageStyle
                   {
                        ImageTitle = 1,//图题
                        ViceImageTitleStyle =1,//副图题
                        CaptionStyle =1,//图注样式
                        ImageInterval = 0.0,//图片间距，插图
                        ImageBitmap ="自适应",//图片位图，插图
                        ImageVectorgraph  ="自适应",//图片矢量图，插图
                    }
                }
            }
            }
        },
        new StyleCategory {
            Name ="表格样式",
            Styles = new List<StyleType>{
            new StyleType{
                Type = "表格样式",
                ExcelStyles =new List<UseExcelStyle>
                {
                    new UseExcelStyle{
                         TitleStyleId =0,
                        ChineseFontType="方正细等线",
                        EnglishFontType ="Time New Roman",
                        FontSize="小四",
                         EndLineSpace=2.0,
                         OutSideLine=2.0,
                         InsideLine=2.0
                     }
                }
            }
            }
        },
        new StyleCategory {
            Name = "模型样式",
            Styles = new List<StyleType>{
                new StyleType{
                    Type="模型样式",
                    ModelStyles = new List<UseModelStyle>
                    {
                         new UseModelStyle{
                               TitleStyleId=0,
                    }
                    }
                }
            }
        },
         new StyleCategory {
            Name = "音视频样式",
            Styles = new List<StyleType>{
                new StyleType{
                    Type="音视频样式",
                    VideoStyles = new List<UseVideoStyle>
                    {
                         new UseVideoStyle{
                             VideoStlyleName ="默认样式",
                             VideoMp3Name ="默认样式",
                            }
                    }
                }
            }
        },
        new StyleCategory {
            Name = "习题样式",
            Styles = new List<StyleType>{
                new StyleType{
                    Type="习题样式",
                    ExercisesStyle = new List<UseExercisesStyle>
                    {
                         new  UseExercisesStyle{
                                TitleStyleId =1,
                                QuestionStemStyleId=1,
                                OptionStyleId =1
                              }
                    }
                }
            }
        },
    };


    
   
       
        
    
    }
};
