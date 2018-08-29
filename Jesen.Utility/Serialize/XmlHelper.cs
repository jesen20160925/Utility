using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Jesen.Utility.Serialize
{
    /// <summary>
    /// 使用序列化器完成的XML序列化帮助类
    /// </summary>
    public class XmlHelper
    {
        #region 通过XmlSerializer序列化实体
        /// <summary>
        /// 通过XmlSerializer序列化实体
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <returns></returns>
        public static string ToXml<T>(T t) where T : new()
        {
            XmlSerializer xmlSerializer = new XmlSerializer(t.GetType());
            Stream stream = new MemoryStream();
            xmlSerializer.Serialize(stream, t);
            stream.Position = 0;
            StreamReader reader = new StreamReader(stream);
            string text = reader.ReadToEnd();
            return text;
        }
        #endregion

        #region 字符串序列化成XML
        /// <summary>
        /// 字符串序列化成XML
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="content"></param>
        /// <returns></returns>
        public static T ToObject<T>(string content) where T : new()
        {
            using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(content)))
            {
                XmlSerializer xmlFormat = new XmlSerializer(typeof(T));
                return (T)xmlFormat.Deserialize(stream);
            }
        }
        #endregion

        #region 文件反序列化成实体
        /// <summary>
        /// 文件反序列化成实体
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static T FileToObject<T>(string fileName) where T : new()
        {
            string CurrentXMLPath = ConfigurationManager.AppSettings["CurrentXMLPath"];
            fileName = Path.Combine(CurrentXMLPath, typeof(T).Name + @".xml");
            using (Stream fStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                XmlSerializer xmlFormat = new XmlSerializer(typeof(T));
                return (T)xmlFormat.Deserialize(fStream);
            }
        } 
        #endregion
    }
}
