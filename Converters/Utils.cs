using System;
using System.IO;
using System.Text;
using Hexx.Core;

namespace Hexx.Converters
{
    internal static class Utils
    {
        /// <summary>
        /// 파일이 위치할 디렉토리를 확보합니다.
        /// </summary>
        /// <param name="filePath">위치</param>
        public static void SecurePathDirectory(string filePath)
        {
            string extension = Path.GetExtension(filePath);

            if (extension == string.Empty)
            {
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
            }
            else
            {
                string outDir = Path.GetDirectoryName(filePath);

                if (!Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }
            }
        }

        /// <summary>
        /// 각 줄 앞에 텍스트를 추가합니다.
        /// </summary>
        /// <param name="text">소스 텍스트</param>
        /// <param name="additionalText">추가할 텍스트</param>
        public static string AddPerEachLine(string text, string additionalText)
        {
            StringBuilder builder = new StringBuilder();
            foreach (string line in text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
            {
                builder.AppendLine($"{additionalText}{line}");
            }
            return builder.ToString().TrimEnd();
        }
    }
}
