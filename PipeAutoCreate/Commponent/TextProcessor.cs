using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PipeAutoCreate.Commponent
{
    public static class TextProcessor
    {
        public static List<double> ExtractNumbers(string input)
        {
            List<double> numbers = new List<double>();

            // 定义匹配数字的正则表达式
            Regex regex = new Regex(@"\d+(\.\d+)?");

            // 匹配输入字符串
            MatchCollection matches = regex.Matches(input);

            // 将所有匹配的数字添加到列表中
            foreach (Match match in matches)
            {
                // 使用尝试转换，以处理没有小数部分的情况
                double number;
                if (double.TryParse(match.Value, out number))
                {
                    numbers.Add(number);
                }
            }

            return numbers;
        }

    }
}
