using System.Net.Http;
using System.Text.RegularExpressions;

namespace AutomaticChassisCoordinatingCalculator
{
    class Core
    {
        string price;
        string productName;
        string result;

        public string Price
        {
            get
            {
                return price;
            }

            set
            {
                price = value;
            }
        }

        public string ProductName
        {
            get
            {
                return productName;
            }

            set
            {
                productName = value;
            }
        }

        public string Result
        {
            get
            {
                return result;
            }

            set
            {
                result = value;
            }
        }

        public void getResponse(string URL)
        {
            if (URL.Contains("tmall"))
            {
                price = null;
            }
            var urlString = URL.Replace("item.jd.com", "item.m.jd.com/product");
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("UserAgent", "Mozilla/5.0 (Linux; Android 5.1.1; Nexus 6 Build/LYZ28E) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.23 Mobile Safari/537.36");
            this.Result = httpClient.GetStringAsync(urlString).Result.ToString();
        }

        public void getProductName()
        {
            var result = Result;
            Regex r1 = new Regex("<title>[^>]+>");
            if (r1.IsMatch(result))
                productName = r1.Match(result).Value;
            productName = Regex.Replace(productName, @"\s", "");
            productName = productName.Replace("<title>", "");
            productName = productName.Replace("-京东</title>", "");
            productName = productName.Replace("-淘宝网</title>", "");
            productName = productName.Replace("-tmall.com天猫</title>", "");
        }

        public void getPrice()
        {
            var result = Result;
            Regex r1 = new Regex(@"<input type=""hidden"" id=""jdPrice"" name=""jdPrice"" value=""[^>]+>");
            Regex r2 = new Regex(@"<input type=""hidden"" name=""current_price"" value= ""[^>]+>");
            if (r1.IsMatch(result))
            {
                price = r1.Match(result).Value;
                Regex re = new Regex("(?<=\").*?(?=\")", RegexOptions.None);
                MatchCollection mc = re.Matches(price);
                foreach (Match ma in mc)
                {
                    price = ma.Value;
                }
            }
            else if (r2.IsMatch(result))
            {
                price = r2.Match(result).Value;
                Regex re = new Regex("(?<=\").*?(?=\")", RegexOptions.None);
                MatchCollection mc = re.Matches(price);
                foreach (Match ma in mc)
                {
                    price = ma.Value;
                }
            }
            else
            {
                price = "0";
            }
        }
    }
}
