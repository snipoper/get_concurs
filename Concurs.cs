using HtmlAgilityPack;
using OfficeOpenXml;
using Microsoft.AspNetCore.Mvc;

namespace FSAPortal.Controllers
{
    public class Concurs : Controller
    {

        public class Contract
        {
            public string Sheet { get; set; }
            public string Name { get; set; }
            public string Date1 { get; set; }
            public string Date2 { get; set; }
            public string Price { get; set; }
            public string Link { get; set; }
            public string Tip { get; set; }
            public string Tip2 { get; set; }

            public Contract(string sheet, string name, string date1, string date2, string price, string link, string tip, string tip2)
            {
                Sheet = sheet;
                Name = name;
                Date1 = date1;
                Date2 = date2;
                Price = price;
                Link = link;
                Tip = tip;
                Tip2 = tip2;
            }
        }

        [HttpGet]
        public async Task<ActionResult> MainControllerTest(string checkedInput0, string checkedInput1, string checkedInput2, string checkedInput3, string checkedInput4, string checkedInput5, string checkedInput6, string checkedInput7)
        {
            List<string> searchList = new List<string>();

                if (checkedInput0 != null)
                {
                    searchList.Add(checkedInput0.Replace(" ", "%20"));
                };
                if (checkedInput1 != null)
                {
                    searchList.Add(checkedInput1.Replace(" ", "%20"));
                };
                if (checkedInput2 != null)
                {
                    searchList.Add(checkedInput2.Replace(" ", "%20"));
                };
                if (checkedInput3 != null)
                {
                    searchList.Add(checkedInput3.Replace(" ", "%20"));
                };
                if (checkedInput4 != null)
                {
                    searchList.Add(checkedInput4.Replace(" ", "%20"));
                };
                if (checkedInput5 != null)
                {
                    searchList.Add(checkedInput5.Replace(" ", "%20"));
                };
                if (checkedInput6 != null)
                {
                    searchList.Add(checkedInput6.Replace(" ", "%20"));
                };
                if (checkedInput7 != null)
                {
                    searchList.Add(checkedInput7.Replace(" ", "%20"));
                };

            List<string> rtsTenderList = new List<string>
            {
                "search",
                "platforms/rts-tender",
                "platforms/sberbank-ast",
                "platforms/etp-tender",
                "platforms/federatsiya",
                "platforms/rosatom",
                "platforms/fabrikant",
                "platforms/otc",
                "platforms/b2b-center",
                "platforms/ast-goz",
                "platforms/etp-gpb",
                "platforms/eetp-roseltorg",
                "platforms/etp-tek-torg",
                "platforms/sibur",
                "platforms/megafon",
                "platforms/estp",
                "platforms/tsentr-realizatsii",
                "platforms/komos",
                "platforms/talan",
                "platforms/etp-mmvb",
                "platforms/bashneft",
                "platforms/etp-rzhd",
                "platforms/etprf",
                "platforms/uralkalij",
                "platforms/etp-223",
                "platforms/tender-pro",
                "platforms/etp-torgi-223",
                "platforms/etp-referi",
                "platforms/vtb-tsentr",
                "platforms/slavneft",
                "platforms/ets",
                "platforms/gazneftetorg",
                "platforms/gazprombank",
                "platforms/etp-rad",
                "platforms/nis",
                "platforms/agz-rt",
                "platforms/safmar",
                "platforms/etp-tmk",
                "platforms/torgi-gov-ru",
                "platforms/atts",
                "platforms/etp-lukojl",
                "platforms/lotekspert",
                "platforms/akd",
                "platforms/auktsiony-dalnego-vostoka",
                "platforms/etp-electro-torgi",
                "platforms/etp-bashzakaz",
                "platforms/etp-svfu",
                "platforms/em-mo",
                "platforms/etp-rosseti",
                "platforms/etp-spetsstrojtorg",
                "platforms/etp-tzs-elektra",
                "platforms/etp-setonline",
                "platforms/etp-alrosa",
                "platforms/etp-birzha-sankt-peterburg",
                "platforms/etp-oborontorg",
                "platforms/etp-eltoks",
                "platforms/etpzakupki-tatar",
                "platforms/etp-tattehmedfarm",
                "platforms/etp-onlinecontract",
                "platforms/krymskaya-etp-torgi82-ru",
                "platforms/etp-oao-setevaya-kompaniya",
                "platforms/etp-itender",
                "platforms/etp-kfu",
                "platforms/etp-eltorg",
                "platforms/etp-region-ast",
                "platforms/etp-moskollektor",
                "platforms/etp-teclot",
                "platforms/etp-avtodor",
                "platforms/rosetp",
                "platforms/etp-rtrs",
                "platforms/etp-asgor",
                "platforms/etp-aeroflot",
                "platforms/etp-dveuk",
                "platforms/etp-tatavtodor",
                "platforms/etp-kvadra",
                "platforms/etp-mosvodokanal",
                "platforms/etp-auktsion-tsentr",
                "platforms/etp-rb2b",
                "platforms/etp-auktsiony-sibiri",
                "platforms/etp-npo-verhnevolzhskij-torgovyj-soyuz",
                "platforms/etp-vladzakupki",
                "platforms/etp-ozk",
                "platforms/etp-stolitsa",
                "platforms/etp-zakupki21-ru",
                "platforms/ams-servis",
                "platforms/etp-erus",
                "platforms/etp-elektronnyj-kapital",
                "platforms/portal-postavshchikov-moskvy",
                "platforms/vetp",
                "platforms/regtorg",
                "platforms/seltim",
                "platforms/rhtorg",
                "platforms/etp-gruppy-lsr",
                "platforms/segezhagroup",
                "platforms/national-etp",
                "platforms/berezka",
                "platforms/etp-zakazrf",
                "platforms/etp-ets"
            };
            List<string> tenderUrlList = new List<string>();

            List<Contract> resultList = new List<Contract>();

            int countSearch = 0;

            List<Task> tasks = new List<Task>();
            int howmany = 0;

            foreach (string search in searchList)
            {

                tasks.Add(Task.Run(async () =>
                    {
                    foreach (string tender in rtsTenderList)
                    {
                        
                            try
                            {
                                using (var client = new HttpClient())
                                {
                                    string url = "https://zakupki360.ru/" + tender + "?query=" + search;
                                    string html = "";
                                    howmany = howmany + 1;

                                    Console.WriteLine((rtsTenderList.Count * searchList.Count) - howmany);

                                    try
                                    {
                                        html = await client.GetStringAsync(url);
                                    }
                                    catch
                                    {

                                        html = await client.GetStringAsync(url);
                                    }
                                    var doc = new HtmlDocument();
                                    doc.LoadHtml(html);

                                    var contractElements = doc.DocumentNode.SelectNodes("//div[@class='content__content']//a[@class='app-passive-link']");
                                    var contractElementsCatch = doc.DocumentNode.SelectNodes("//div[@class='content__header']");

                                    string contractType1 = "";

                                    foreach (var contractElement in contractElements)
                                    {
                                        string contractHref = contractElement.GetAttributeValue("href", "");

                                        if (tenderUrlList.Contains(contractHref))
                                        {
                                            continue;
                                        }

                                        tenderUrlList.Add(contractHref);
                                        try
                                        {
                                            url = "https://zakupki360.ru" + contractHref;
                                            html = await client.GetStringAsync(url);
                                        }
                                        catch
                                        {
                                            continue;
                                        }

                                        var docInHref = new HtmlDocument();
                                        docInHref.LoadHtml(html);

                                        int index0 = contractElements.IndexOf(contractElement);
                                        string getInnerText = contractElementsCatch[index0].InnerText;

                                        if(getInnerText.Contains("223-ФЗ")){
                                            contractType1 = "223-ФЗ";
                                        }
                                        else if (getInnerText.Contains("Коммерческие")){
                                            contractType1 = "Коммерческие";
                                        }
                                        else {
                                            contractType1 = getInnerText;
                                            continue;
                                        }

                                        var contractType21 = docInHref.DocumentNode.SelectNodes("//div[@class='data info__data']").ToList();

                                        var contractType2 = contractType21[4].InnerText;

                                        //string contractName = docInHref.DocumentNode.SelectSingleNode("//h1[@class='dossier__title']")?.InnerText;
                                        var contractNameList = docInHref.DocumentNode.SelectNodes("//h1[@class='dossier__title']").ToList();
                                        var contractName = contractNameList[0].InnerText;
                                        int index = contractName.LastIndexOf("(на сумму");
                                        try
                                        {
                                            contractName = contractName.Remove(index).Trim();
                                        }
                                        catch { }
                                        //contractName = Regex.Replace(contractName, " \\(.*\\)$", "");

                                        string contractStartDate = docInHref.DocumentNode.SelectSingleNode("//div[@class='dossier__column data']")?.InnerText.Trim();
                                        contractStartDate = contractStartDate.Substring(0, 20).Substring(contractStartDate.Length - 10);
                                        var contractStartDate2 = docInHref.DocumentNode.SelectNodes("//span[@class='ng-star-inserted']").ToList();
                                        var contractStartDate22 = "";
                                        try
                                        {
                                            contractStartDate22 = contractStartDate2[2].InnerText;
                                        }
                                        catch
                                        {
                                            contractStartDate22 = "";
                                        }

                                        string contractPriceGet = contractType21[0].InnerText;
                                        string contractPrice = contractPriceGet.Replace("&nbsp;", " ");

                                        string contractContent = docInHref.DocumentNode.SelectSingleNode("//meta[@itemprop='url']")?.Attributes["content"]?.Value;
                                        if (contractContent == "")
                                        {
                                            contractContent = url;
                                        }

                                        lock (resultList)
                                        {
                                            resultList.Add(new Contract(search, contractName, contractStartDate, contractStartDate22, contractPrice, contractContent, contractType2, contractType1));
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                        
                    }
                }));
                countSearch++;
            }

            await Task.WhenAll(tasks);

            resultList.Sort((r1, r2) =>
            {
                int index1 = searchList.IndexOf(r1.Sheet);
                int index2 = searchList.IndexOf(r2.Sheet);
                return index1.CompareTo(index2);
            });

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var templateFile = new FileInfo("wwwroot/Конкурсы/Шаблон закупки.xlsx");
            var outputFilePath = "wwwroot/Конкурсы/Итог выгрузка.xlsx";

            using (var package = new ExcelPackage(templateFile))
            {
                var workbook = package.Workbook;
                // Заполнение данными
                int count = 2;
                int oldIndex = 0;
                try
                {
                    foreach (var item in resultList)
                    {
                        int index = searchList.IndexOf(item.Sheet);
                        if (oldIndex != index && index != 0)
                        {
                            oldIndex = index;
                            count = 2;
                        }

                        String sheetName = searchList[index].Replace("%20", " ");
                        if (sheetName == checkedInput0)
                        {
                            index = 0;
                        }
                        if (sheetName == checkedInput1)
                        {
                            index = 1;
                        }
                        if (sheetName == checkedInput2)
                        {
                            index = 2;
                        }
                        if (sheetName == checkedInput3)
                        {
                            index = 3;
                        }
                        if (sheetName == checkedInput4)
                        {
                            index = 4;
                        }
                        if (sheetName == checkedInput5)
                        {
                            index = 5;
                        }
                        if (sheetName == checkedInput6)
                        {
                            index = 6;
                        }
                        if (sheetName == checkedInput7)
                        {
                            index = 7;
                        }

                        var sheet = workbook.Worksheets[index];
                        
                        sheet.Name = sheetName;

                        var dataRow1 = sheet.Row(count);
                        sheet.Cells[count, 1].Value = item.Name;
                        sheet.Cells[count, 2].Value = item.Date1;
                        sheet.Cells[count, 3].Value = item.Date2;
                        sheet.Cells[count, 4].Value = item.Price;
                        sheet.Cells[count, 5].Value = item.Link;

                        ExcelHyperLink hyperlink = (ExcelHyperLink)(sheet.Cells[count, 5].Hyperlink = new ExcelHyperLink(item.Link));

                        sheet.Cells[count, 6].Value = item.Tip;
                        sheet.Cells[count, 7].Value = item.Tip2;

                        count++;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }

                var file = new FileInfo(outputFilePath);
                package.SaveAs(file);
            }

            return new EmptyResult();
        }
    }
}