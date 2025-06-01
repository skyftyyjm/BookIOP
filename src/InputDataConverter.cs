using ioh;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ioh
{
    public class InputDataConverter : JsonConverter//自定义JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(InputData);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            var jsonObject = JObject.Load(reader);
            string type = jsonObject["Types"].ToString();

            InputData inputData = type switch
            {
                "标题" => new TitleData(),
                "段落" => new ParagraphData(),
                "图片" => new ImageData(), // 添加对图片类型的支持
                "公式" => new Formula(), // 也可以添加其他类型
                "音视频" => new Video(),   // 添加音视频类型
                "表格" => new Table(),
                "模型" => new Model(),
                "习题" => new ExercisesStyle(),
                "习题标题" => new QuessionTitle(),
                "习题单选" => new QuessionChoose(),
                "习题多选" => new QuessionMultipleChoice(),
                "习题主观" => new QuessionSubjectivity(),


                _ => throw new NotSupportedException($"不支持的类型: {type}")
            };

            serializer.Populate(jsonObject.CreateReader(), inputData);
            return inputData;
        }










        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            // 实现序列化逻辑（如果需要）
            throw new NotImplementedException();
        }
    }
}
