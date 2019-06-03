using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace RoS_BOTWindowsService
{
    /// <summary>
    /// Config By C.Y. Fang
    /// </summary>
    [Serializable]
    [XmlRoot(ElementName = "List")]
    public class Config
    {
        /// <summary>
        /// StringWriterWithUTF8Encoding
        /// </summary>
        private sealed class StringWriterWithUTF8Encoding : StringWriter
        {
            /// <summary>
            /// UTF8Encoding
            /// </summary>
            private UTF8Encoding UTF8Encoding { get { return new UTF8Encoding(); } }
            /// <summary>
            /// Encoding
            /// </summary>
            public override Encoding Encoding => UTF8Encoding;
        }

        /// <summary>
        /// Config path
        /// </summary>
        public static readonly String ConfigPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Config.xml";
        /// <summary>
        /// Read xml to string
        /// </summary>
        /// <returns></returns>
        private static readonly string Read = File.ReadAllText(ConfigPath, Encoding.UTF8);
        /// <summary>
        /// User
        /// </summary>
        [XmlElement(ElementName = "User")]
        public string User { get; set; }
        /// <summary>
        /// Password
        /// </summary>
        [XmlElement(ElementName = "Password")]
        public string Password { get; set; }
        /// <summary>
        /// Timeout
        /// </summary>
        [XmlElement(ElementName = "Timeout")]
        public string Timeout { get; set; }
        /// <summary>
        /// vmrun.exe path
        /// </summary>
        [XmlElement(ElementName = "VMRUN_Path")]
        public string VMrunPath { get; set; }
        /// <summary>
        /// VM path
        /// </summary>
        [XmlElement(ElementName = "MachinePath")]
        public string MachinePath { get; set; }
        /// <summary>
        /// Interval
        /// </summary>
        [XmlElement(ElementName = "Interval")]
        public string Interval { get; set; }


        /// <summary>
        /// Deserialize
        /// </summary>
        /// <param name="serializedObj"></param>
        /// <returns></returns>
        public Config Deserialize()
        {
            string str = Read;
            if (str == null || str.Equals(string.Empty))
                return null;

            var serializer = new XmlSerializer(typeof(Config));
            using (var stringReader = new StringReader(str))
            {
                using (var xmlTextReader = new XmlTextReader(stringReader))
                {
                    var config = (Config)serializer.Deserialize(xmlTextReader);
                    return config;
                }
            }
        }
    }

}
