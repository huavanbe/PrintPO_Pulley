using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.Helpers
{
    class JsonHelper
    {
        //save file to json
        public static void SaveJsonObject(object data, string fileName)
        {
            using (StreamWriter writer = File.CreateText(AppDomain.CurrentDomain.BaseDirectory + @"DATA\" + fileName))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(writer, data);
            }

        }


        public static void SaveJson2(object data, string filePath)
        {
            using (StreamWriter writer = File.CreateText(filePath))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(writer, data);
            }

        }
        //read json file as text
        public static string ReadFileToString(string filePath)
        {
            string text = "";
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                text = streamReader.ReadToEnd();
            }
            return text;
        }

        //read json file to object
        public static T ReadJsonToObject<T>(string fileName)
        {
            T data;
            using (StreamReader reader = File.OpenText(AppDomain.CurrentDomain.BaseDirectory + @"DATA\" + fileName))
            {
                JsonSerializer serializer = new JsonSerializer();
                data = (T)serializer.Deserialize(reader, typeof(T));
            }
            return data;
        }

        public static T ReadJson2<T>(string filePath)
        {
            T data;
            using (StreamReader reader = File.OpenText(filePath))
            {
                JsonSerializer serializer = new JsonSerializer();
                data = (T)serializer.Deserialize(reader, typeof(T));
            }
            return data;
        }
    }
}
