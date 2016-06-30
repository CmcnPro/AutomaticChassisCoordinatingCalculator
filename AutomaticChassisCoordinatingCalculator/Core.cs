using System.Net.Http;
using System.Text.RegularExpressions;

namespace AutomaticChassisCoordinatingCalculator
{
    class Core
    {
        string price;
        string productName;
        string result;
        string brandName;
        string model;

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

        public string BrandName
        {
            get
            {
                return brandName;
            }

            set
            {
                brandName = value;
            }
        }

        public string Model
        {
            get
            {
                return model;
            }

            set
            {
                model = value;
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

        public void getBrandName(string URL)
        {
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("UserAgent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36");
            this.Result = httpClient.GetStringAsync(URL).Result.ToString();
            var result = Result;
            Regex r1 = new Regex(">品牌</td><td>[^>]+>");//JD
            Regex r2 = new Regex(">品牌[^>]+>");//TB,TM
            if (r1.IsMatch(result))
            {
                brandName = r1.Match(result).Value;
                brandName = brandName.Replace(">品牌</td><td>", "");
                brandName = brandName.Replace("</td>", "");
            }else
            if (r2.IsMatch(result))
            {
                brandName = r2.Match(result).Value;
                brandName = brandName.Replace(">品牌:&nbsp;", "");
                brandName = brandName.Replace("</li>", "");
                Regex re = new Regex("(?<=\").*?(?=\")", RegexOptions.None);
                MatchCollection mc = re.Matches(brandName);
                foreach (Match ma in mc)
                {
                    brandName = ma.Value;
                    if (URL.Contains("tmall"))
                    {
                        string[] _list = brandName.Split(' ');
                        brandName = _list[0];
                    }
                }
            }
        }

        public void getModel(string URL)
        {
            var result = Result;
            Regex r1 = new Regex(">型号</td><td>[^>]+>");//JD
            Regex r2 = new Regex("<title>[^>]+>");//TB
            if (r1.IsMatch(result))
            {
                model = r1.Match(result).Value;
                model = model.Replace(">型号</td><td>", "");
                model = model.Replace("</td>", "");
            }
            else if (r2.IsMatch(result))
            {
                model = r2.Match(result).Value;
                Regex re = new Regex("(?<=\").*?(?=\")", RegexOptions.None);
                MatchCollection mc = re.Matches(model);
                foreach (Match ma in mc)
                {
                    model = ma.Value;
                }
                model = Regex.Replace(model, @"\s", "");
                model = model.Replace("<title>", "");
                model = model.Replace("-淘宝网</title>", "");
                model = model.Replace("-tmall.com天猫</title>", "");
            }
        }

        public void getPrice()
        {
            var result = Result;
            Regex r1 = new Regex(@"<input type=""hidden"" id=""jdPrice"" name=""jdPrice"" value=""[^>]+>");//JD
            Regex r2 = new Regex(@"<input type=""hidden"" name=""current_price"" value= ""[^>]+>");//TB
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
