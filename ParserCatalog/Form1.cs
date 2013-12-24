using System.Collections.Specialized;
using System.Drawing;
using System.Threading;
using System.Web;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using WatiN.Core;
using Form = System.Windows.Forms.Form;

namespace ParserCatalog
{
    public partial class Form1 : Form
    {
        private List<Shop> shops;
        private List<ShopBig> shopBigs;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            shops = new List<Shop>();
            shopBigs = new List<ShopBig>();
            treeView1.Nodes.Clear();
            shops.Add(new Shop() { Name = "Leggi", Url = "http://leggi.com.ua/" });
            shops.Add(new Shop() { Name = "Trikobach", Url = "http://trikobakh.com" });
            shops.Add(new Shop() { Name = "Твоё", Url = "http://tvoe.ru/collection/" });
            shops.Add(new Shop() { Name = "ОЗКАН", Url = "http://www.ozkanwear.com" });
            //shops.Add(new Shop() { Name = "Trimedwedya", Url = "http://www.trimedwedya.ru/index.php/kategorii-tovarov" });
            shops.Add(new Shop() { Name = "S-trikbel", Url = "http://s-trikbel.ru/" });
            shops.Add(new Shop() { Name = "Butterfly-dress", Url = "http://butterfly-dress.com" });
            shops.Add(new Shop() { Name = "Roomdecor", Url = "http://roomdecor.su" });
            shops.Add(new Shop() { Name = "Aventum", Url = "http://aventum.cz" });
            shops.Add(new Shop() { Name = "Sportoptovik", Url = "http://www.sportoptovik.ru/" });
            shops.Add(new Shop() { Name = "Nashipupsi", Url = "http://nashipupsi.ru" });
            shops.Add(new Shop() { Name = "Спорт-Опт", Url = "http://xn----0tbbbddeld.xn--p1ai/catalog/" });
            shops.Add(new Shop() { Name = "Адель", Url = "http://td-adel.ru/catalog/bedclothes/" });
            shops.Add(new Shop() { Name = "Artvision", Url = "http://artvision-opt.ru" });
            shops.Add(new Shop() { Name = "OptEconom", Url = "http://www.opt-ekonom.ru" });
            shops.Add(new Shop() { Name = "Naksa", Url = "http://naksa.ru" });
            shops.Add(new Shop() { Name = "Nobi54", Url = "http://www.nobi54.ru" });
            shops.Add(new Shop() { Name = "Lemming", Url = "http://lemming.su/katalog" });
            shops.Add(new Shop() { Name = "Piniolo", Url = "http://www.piniolo.ru/" });
            shops.Add(new Shop() { Name = "Witerra", Url = "http://witerra.ru/" });
            shops.Add(new Shop() { Name = "Gipnozstyle", Url = "http://ru.gipnozstyle.ru" });
            shops.Add(new Shop() { Name = "Noski-a42", Url = "http://noski-a42.ru/" });
            shops.Add(new Shop() { Name = "Iv-trikotage", Url = "http://iv-trikotage.ru" });
            shops.Add(new Shop() { Name = "Shop-nogti", Url = "http://shop-nogti.ru" });
            shops.Add(new Shop() { Name = "Npopt", Url = "http://npopt.ru" });
            shops.Add(new Shop() { Name = "Optovik-centr", Url = "http://optovik-centr.ru" });
            shops.Add(new Shop() { Name = "Japan-cosmetic", Url = "http://japan-cosmetic.biz" }); ;
            //nameCB.Items.AddRange(shops.Select(x => x.Name).ToArray());

            shopBigs.Add(new ShopBig()
            {
                Name = "Maximum",
                Url = "http://maximumufa.ru/represent/_represent_catalog/index.php",
                XPath = "//section/article/a",
                Host = "http://maximumufa.ru/represent/_represent_catalog/"
            });
            shopBigs.Add(new ShopBig() { Name = "Otoys", Url = "http://www.otoys.ru/", XPath = "//div[contains(concat(' ', @id, ' '), ' toy_types ')]/ul/li/a" });
            Helpers.GetCatalog(ref shopBigs);
            //Загрузка магазина в TreeView
            int i = 0;
            foreach (var big in shopBigs)
            {
                treeView1.Nodes.Add(big.Name);
                foreach (var cat in big.CatalogList)
                {
                    treeView1.Nodes[i].Nodes.Add(cat.Name);
                }
                i++;
            }

            foreach (var sh in shops)
            {
                treeView1.Nodes.Add(sh.Name);
            }
        }

        private void Start_Click(object sender, EventArgs e)
        {
            Start.Enabled = false;
            Start.Text = "Подождите...";
            var pars = new List<Site>();
            var leggi = new List<string>();
            foreach (TreeNode aNode in treeView1.Nodes)
            {
                var t1 = new List<Category>();
                foreach (TreeNode node in aNode.Nodes)
                {
                    if (node.Checked)
                    {
                        var url = shopBigs.FirstOrDefault(x => x.Name == aNode.Text).CatalogList.FirstOrDefault(x => x.Name == node.Text).Url;
                        t1.Add(new Category() { Name = node.Text, Url = url });
                    }
                }
                if (t1.Any())
                    pars.Add(new Site() { Name = aNode.Text, Categories = t1, Catalog = true });
                else
                {
                    var sh = shops.Where(x => x.Name == aNode.Text).ToList();
                    if (aNode.Checked && sh != null && sh.Any())
                    {
                        pars.Add(new Site() { Name = aNode.Text, Categories = new List<Category>() { new Category() { Name = sh[0].Name, Url = sh[0].Url } } });
                    }
                }
            }
            var stL = new List<string>();
            var st = new Stopwatch();
            st.Start();
            countStripStatus.Text = "Скачено 0 из " + pars.Count;
            var errors = new List<string>();
            Parallel.ForEach(pars, site =>
            {
                try
                {
                    var cL = new List<Category>();
                    var shopUrl = site.Categories[0].Url;
                    if (!site.Catalog)
                    {
                        var catList = new List<Category>();

                        var client = new System.Net.WebClient();
                        var data = client.OpenRead(shopUrl);
                        var reader = new StreamReader(data, Encoding.GetEncoding("windows-1251"));
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var page = new HtmlAgilityPack.HtmlDocument();
                        page.LoadHtml(s);

                        //get catalogs
                        var arr = new string[]
                    {
                        "categ", "catal", "woman", "man", "katalog", "kategorii", "platja", "aksessuary", "roomdecor",
                        "folder", "collect"
                    };

                        var query = "//ul/li/a";
                        if (shopUrl.Contains("trimedwedya") || shopUrl.Contains("artvision-opt"))
                            query = "//ul/li/ul/li/a";
                        else if (shopUrl.Contains("butterfly-dress"))
                            query = "//ul/li/ul/li/div/a";
                        else if (shopUrl.Contains("s-trikbel"))
                            query = "//li[contains(concat(' ', @class, ' '), ' name ')]/a";
                        else if (shopUrl.Contains("roomdecor"))
                            query = "//li/ul/li/a";
                        else if (shopUrl.Contains("nashipupsi"))
                            query = "//a[contains(concat(' ', @href, ' '), 'folder')]";
                        else if (shopUrl.Contains("opt-ekonom"))
                            query = "//span[contains(concat(' ', @class, ' '), ' inner ')]/a";
                        else if (shopUrl.Contains("lemming"))
                            query = "//span/a";
                        else if (shopUrl.Contains("piniolo"))
                            query = "//li[contains(concat(' ', @class, ' '), ' item ')]/a";
                        else if (shopUrl.Contains("witerra"))
                            query = "//td[contains(concat(' ', @class, ' '), ' boxText ')]/a";
                        else if (shopUrl.Contains("ru.gipnozstyle"))
                            query = "//div[contains(concat(' ', @class, ' '), ' twocol ')]/a";
                        else if (shopUrl.Contains("shop-nogti"))
                            query = "//div/div/div/a";
                        else if (shopUrl.Contains("iv-trikotage"))
                            query = "//div[contains(concat(' ', @class, ' '), ' menu_spec ')]/ul/li/a";
                        else if (shopUrl.Contains("optovik-centr"))
                            query = "//a[contains(concat(' ', @class, ' '), ' mainlevel_frontpage_categories ')]";
                        else if (shopUrl.Contains("japan-cosmetic"))
                            query = "//div[contains(concat(' ', @class, ' '), 'moduletableproizv')]/div/a";

                        var cats = page.DocumentNode.SelectNodes(query);
                        foreach (var cat in cats)
                        {
                            if (cat.Attributes.Count > 0)
                            {
                                var link = cat.Attributes["href"].Value;
                                bool good = arr.Any(ar => link.Contains(ar));
                                if (link.Contains("s-trikbel") || shopUrl.Contains("artvision-opt") ||
                                    shopUrl.Contains("opt-ekonom") || shopUrl.Contains("witerra") ||
                                    shopUrl.Contains("ru.gipnozstyle") || shopUrl.Contains("trikotage") ||
                                    shopUrl.Contains("npopt") || shopUrl.Contains("japan-cosmetic"))
                                    good = true;
                                if (link.Contains("roomdecor") && (link.Contains("6195") || link.Contains("6159")))
                                    good = false;
                                if (shopUrl.Contains("xn----0tbbbddeld.xn--p1ai") || shopUrl.Contains("td-adel"))
                                {
                                    var t1 = cat.ParentNode.InnerHtml;
                                    if (t1.Contains("<ul") || link.Contains("new"))
                                        good = false;
                                }

                                if (good)
                                {
                                    if (shopUrl.Contains("www.trimedwedya.ru"))
                                        link = "http://www.trimedwedya.ru" + link;
                                    else if (shopUrl.Contains("td-adel"))
                                        link = "http://td-adel.ru" + link;
                                    else if (shopUrl.Contains("xn----0tbbbddeld.xn--p1ai"))
                                        link = "http://xn----0tbbbddeld.xn--p1ai/" + link;
                                    else if (shopUrl.Contains("lemming.su"))
                                        link = "http://lemming.su" + link;
                                    else if (!link.Contains(shopUrl))
                                        link = shopUrl + link;
                                    catList.Add(new Category() { Name = cat.InnerText, Url = WebUtility.HtmlDecode(link) });
                                }
                            }
                        }
                        var temp = new HashSet<string>(catList.Select(x => x.Url));

                        if (temp.Count != catList.Count)
                        {
                            foreach (var t in temp)
                            {
                                foreach (var g in catList)
                                {
                                    if (t.Contains(g.Url))
                                    {
                                        cL.Add(g);
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            cL = catList;
                        }
                    }
                    else
                        cL = site.Categories;
                    if (shopUrl.Contains("tvoe"))
                        GetTvoi(cL.Select(x => x.Url));
                    else if (shopUrl.Contains("ozkan"))
                        GetOzkan(cL);
                    else if (shopUrl.Contains("leggi"))
                    {
                        leggi=cL.Select(x => x.Url).ToList();
                    }
                    else if (shopUrl.Contains("trikobakh"))
                        GetTrikobakh(cL);
                    else if (shopUrl.Contains("trimedwedya"))
                        GetTrimedwedya(cL);
                    else if (shopUrl.Contains("s-trikbel"))
                        GetTrikbel(cL);
                    else if (shopUrl.Contains("butterfly-dress"))
                        GetButterfly(cL);
                    else if (shopUrl.Contains("aventum"))
                        GetAventum(cL);
                    else if (shopUrl.Contains("sportoptovik"))
                        GetSportoptovik(cL);
                    else if (shopUrl.Contains("roomdecor"))
                        GetRoomdecor(cL);
                    else if (shopUrl.Contains("nashipupsi"))
                        GetNashipupsi(cL);
                    else if (shopUrl.Contains("xn----0tbbbddeld.xn--p1ai"))
                    {
                        cL.RemoveAt(0);
                        GetSportOpt(cL);
                    }
                    else if (shopUrl.Contains("artvision-opt"))
                        GetArtvision(cL);
                    else if (shopUrl.Contains("td-adel"))
                    {
                        cL.RemoveAt(0);
                        GetAdel(cL);
                    }
                    else if (shopUrl.Contains("opt-ekonom"))
                        GetOptEconom(cL);
                    else if (shopUrl.Contains("naksa"))
                        GetNaksa(cL);
                    else if (shopUrl.Contains("nobi54"))
                        GetNobi(cL);
                    else if (shopUrl.Contains("lemming"))
                        GetLemming(cL);
                    else if (shopUrl.Contains("piniolo"))
                        GetPiniolo(cL);
                    else if (shopUrl.Contains("witerra"))
                        GetWiterra(cL);
                    else if (shopUrl.Contains("gipnozstyle"))
                    {
                        cL.RemoveAt(0);
                        GetGipnozstyle(cL);
                    }
                    else if (shopUrl.Contains("noski-a42"))
                        GetNoski(cL);
                    else if (shopUrl.Contains("trikotage"))
                    {
                        if (!cL.Any())
                            cL.Add(new Category() { Url = "http://iv-trikotage.ru/" });
                        GetTrikotage(cL);
                    }
                    else if (shopUrl.Contains("shop-nogti"))
                        GetShopNogti(cL);
                    else if (cL[0].Url.Contains("npopt"))
                        GetNpopt(cL);
                    else if (cL[0].Url.Contains("optovik-centr"))
                        GetOptovikCentr(cL);
                    else if (cL[0].Url.Contains("japan-cosmetic"))
                        GetJapanCosmetic(cL);
                    else if (cL[0].Url.Contains("maximum"))
                        GetMaximum(cL);
                    else if (cL[0].Url.Contains("otoys"))
                        GetOtoys(cL);
                    stL.Add(st.Elapsed.ToString());
                }
                catch (Exception ex)
                {
                    errors.Add(site.Name+" - Ошибка: "+ex.Message);
                }
            });
            timeStripStatus.Text = "Время парсинга " + st.Elapsed;
            countStripStatus.Text = "Загружено " + (pars.Count - errors.Count-1) + " из "+pars.Count;
            
            if (leggi.Any())
            {
                GetLeggi(leggi);
            }
            st.Stop();
            stL.Add(st.Elapsed.ToString());
            Start.Enabled = true;
            timeStripStatus.Text = "Время парсинга " + st.Elapsed;
            countStripStatus.Text = "Загружено " + (pars.Count - errors.Count) + " сайтов";
            if (errors.Count > 0)
            {
                var ss = "";
                foreach (var s in errors)
                {
                    ss += s + "\r\n";
                }
                MessageBox.Show(ss.Trim());
            }
            Start.Text = "Начать парсинг";
        }

        private void GetOtoys(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.otoys.ru/", new NameValueCollection());

            var folder = path.Text + @"\" + "Otoys";
            if (photoCheck.Checked)
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
            }
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.otoys.ru/",
                    "//div[contains(concat(' ', @class, ' '), ' markTitle ')]/a",
                    "//div[contains(concat(' ', @class, ' '), ' paging ')]/a[contains(text(), 'Последняя')] | //div[contains(concat(' ', @class, ' '), ' paging ')]/a[last()]", "/page_", null);


                if (prod.Count == 0)
                    continue;
                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 15 == 0)
                    {
                        Thread.Sleep(7000);
                    }
                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' mark_header ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' artMarket ')]").Replace("Aртикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' goodPrice ')]").Replace("р.", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' description ')]/p", "", new List<string>() { "Категория" }).Replace("<!-- no autotypograph --><!-- no autoclear -->", "").Replace("Уведомить о поступлении.", "").Replace("Обращаем Ваше внимание, чтобы забрать товар необходимо оформить заказ и получить подтверждение от оператора.", "").Trim();
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'market_images')]");
                    
                    var tt = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' categoryLinks ')]");
                    if (tt != null)
                    {
                        var t = Regex.Split(tt[0].InnerHtml, "<br>");
                        foreach (var s in t)
                        {
                            var page = new HtmlAgilityPack.HtmlDocument();
                            page.LoadHtml(s);
                            cat = Helpers.GetItemsInnerText(page, "//a", "", null, "/");
                            break;
                        }
                    }
                    var photo = "";
                    if (photoCheck.Checked)
                    {
                        var p = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'market_images')]/img", "", "", "", "", "src");
                        photo = Helpers.SavePhoto(p, folder);
                    }
                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs,
                        Photo = photo
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Otoys.xlsx", photoCheck.Checked);
            StatusStrip("Otoys");
            
        }

        private void GetMaximum(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://maximumufa.ru/", new NameValueCollection());

            var folder = path.Text + @"\" + "Maximum";
            if (photoCheck.Checked)
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
            }
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://maximumufa.ru",
                    "//td[2]/a",
                    "//p[contains(concat(' ', @class, ' '), ' commonPager ')]/a[contains(text(), 'Вывести всё')]", null, "", "http://maximumufa.ru/represent/_represent_catalog/");

                if (prod.Count == 0)
                {
                    var cat = Helpers.GetProductLinks(catalog.Url, cook, "http://maximumufa.ru", "//section/ul/li/a", null, "http://maximumufa.ru/represent/_represent_catalog/");
                    if (cat.Any())
                    {
                        var temp = new List<string>();
                        foreach (var c in cat)
                        {
                            var tr = Helpers.GetProductLinks(c, cook, "http://maximumufa.ru", "//td/a", "//p[contains(concat(' ', @class, ' '), ' commonPager ')]/a[contains(text(), 'Вывести всё')]", null, "", "http://maximumufa.ru/represent/_represent_catalog/");
                            if (tr.Any())
                                temp.AddRange(tr.ToList());
                            else
                            {
                                var cat2 = Helpers.GetProductLinks(catalog.Url, cook, "http://maximumufa.ru", "//section/ul/li/a", null, "http://maximumufa.ru/represent/_represent_catalog/");
                                if (cat.Any())
                                {
                                    foreach (var c2 in cat2)
                                    {
                                        var tr2 = Helpers.GetProductLinks(c2, cook, "http://maximumufa.ru",
                                            "//a[contains(concat(' ', @rel, ' '), ' prettyPhoto ')]",
                                            "//p[contains(concat(' ', @class, ' '), ' commonPager ')]/a[contains(text(), 'Вывести всё')]", null, "", "http://maximumufa.ru/represent/_represent_catalog/");
                                        if (tr2.Any())
                                            temp.AddRange(tr2.ToList());
                                    }
                                }
                            }
                        }
                        prod = new HashSet<string>(temp);
                    }
                }

                if (prod.Count == 0)
                    continue;
                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 15 == 0)
                    {
                        Thread.Sleep(7000);
                    }
                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' first ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' info ')]/table/tr").Replace("Артикул", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' price ')]").Replace("Р", "").Replace("=", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' info ')]/table/tr", "", new List<string>() { "Артикул" },";;");
                    desc = string.Join("_", desc.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)).Replace("_", " ").Replace(";;",", ").Replace("\r\n","");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'prettyPhoto')]", "", "http://maximumufa.ru");
                    if (phs.Any())
                        phs = phs.Where(x => !x.Contains("no_photo")).ToList();
                    cat = Helpers.GetItemsInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' commonBreadcrumb ')]/a", "", new List<string>() { "Каталог" }, "/");
                    var photo = "";
                    if (photoCheck.Checked)
                    {
                        var p = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'prettyPhoto')]/img", "", "http://maximumufa.ru", "", "", "src");
                        if (p.Any())
                            p = p.Where(x => !x.Contains("no_photo")).ToList();
                        photo = Helpers.SavePhoto(p, folder);
                    }
                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs,
                        Photo = photo
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Maximum.xlsx",photoCheck.Checked);
            StatusStrip("Maximum");
            
        }

        private void GetJapanCosmetic(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://japan-cosmetic.biz/", new NameValueCollection());
            int i = 0;
            var re = new List<string>();
            foreach (var catalog in list)
            {
                if (catalog.Url.IndexOf(".htm") == -1)
                    continue; //  
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://japan-cosmetic.biz", "//td/a[contains(text(), '[Подробнее...]')]", "//ul[contains(concat(' ', @class, ' '), 'pagination')]/li/a", null);
                i += prod.Count;

                if (prod.Count == 0)
                {
                    re.Add(catalog.Url);
                    continue;
                }

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 8 == 0)
                    {
                        Thread.Sleep(5000);
                    }

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' name ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' identifier ')]").Replace("Арт.:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' productPrice ')]").Replace("руб.", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' description ')]/p | //span[contains(concat(' ', @itemprop, ' '), ' description ')]/ul/li", "", null, " ");
                    var desc2 = Helpers.GetItemsInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' brand ')]", "", null, "\r\n");
                    desc = (desc + "\r\n" + desc2).Trim();
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]");

                    cat = Helpers.GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), ' pathway ')]", "", new List<string>() { "На главную" }, "/");

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\JapanCosmetic.xlsx");
            StatusStrip("JapanCosmetic");
        }

        private void GetOptovikCentr(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://optovik-centr.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://optovik-centr.ru", "//td/span[contains(concat(' ', @class, ' '), ' h3 ')]/a", "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);

                if (prod.Count == 0)
                {
                    var cat = Helpers.GetProductLinks(catalog.Url, cook, "http://optovik-centr.ru", "//a[contains(concat(' ', @class, ' '), ' subcat ')]", null);
                    if (cat.Any())
                    {
                        var temp = new List<string>();
                        foreach (var c in cat)
                        {
                            var tr = Helpers.GetProductLinks(c, cook, "http://optovik-centr.ru", "//td/span[contains(concat(' ', @class, ' '), ' h3 ')]/a", "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);
                            temp.AddRange(tr.ToList());
                        }
                        prod = new HashSet<string>(temp);
                    }
                }

                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' h3 ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//td/b").Replace("Артикул:", "").Replace(".", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'price_')]").Replace("руб.", "");
                    desc = Helpers.GetItemsInnerHtml(doc2, "//div[contains(concat(' ', @class, ' '), ' browseProductDescription ')]", "", null, " ");
                    var d = Regex.Split(desc, "<br>");
                    if (d.Any())
                    {
                        desc = "";
                        foreach (var s in d)
                        {
                            //if (s.Contains("Размер"))
                            //{
                            //    size = s.Replace("Размер:", "").Replace("Размеры:", "").Replace(",", ";").Trim();
                            //}
                            //else if (s.Contains("Цвет"))
                            //{
                            //    col = s.Replace("Цвет:", "").Replace("Цвета:", "").Replace(",", ";").Trim();
                            //}
                            //else 
                            if (!string.IsNullOrEmpty(s.Trim()))
                                desc += s + "\r\n";
                        }
                        desc = desc.Trim();
                    }

                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//td/a[contains(concat(' ', @onclick, ' '), 'optovik-centr')]");

                    cat = Helpers.GetItemsInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/a", "", new List<string>() { "Главная", "Каталог" }, "/");

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\OptovikCentr.xlsx");
            StatusStrip("OptovikCentr");
        }

        private void GetNpopt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://npopt.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://npopt.ru", "//div/h3/a", "//div[contains(concat(' ', @class, ' '), ' bottom ')]/ul/li/a", null, catalog.Url);

                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' center_side ')]/h2");
                    var artic = title.Substring(title.Length - 6);
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'price_')]").Replace("руб.", "").Trim();
                    desc = Helpers.GetItemsInnerHtml(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", null, " ");
                    var d = Regex.Split(desc, "<br>");
                    if (d.Count() > 1)
                    {
                        desc = "";
                        foreach (var s in d)
                        {
                            if (s.Length > 0)
                                desc += s + "\r\n";
                        }
                        desc = desc.Trim();
                    }
                    else
                    {
                        desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", null, "\r\n");
                    }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), ' img-popup ')]", "//div[contains(concat(' ', @class, ' '), 'description')]/p/img", "http://npopt.ru", "http://npopt.ru");
                    if (desc.Length == 0)
                        phs.AddRange(Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), 'description')]/p/p/img", "", "", "", "", "src"));
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/ul/li/a", "", new List<string>() { "Главная", "Каталог товаров" }, "/");

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Npopt.xlsx");
            StatusStrip("Npopt");
        }

        private void GetNoski(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://noski-a42.ru/", new NameValueCollection());
            var folder = path.Text + @"\" + "Noski";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "?page=all", cook, "http://noski-a42.ru/", "//div[contains(concat(' ', @class, ' '), ' product_info ')]/h3/a", null);
                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url + "?page=all", null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/h1");
                    var artic = "";
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'price_')]");
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", new List<string>() { "Внимание" }, "\r\n");

                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    var dphs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), ' image ')]/a");
                    if (dphs.Any())
                        phs.Add(dphs[0]);
                    dphs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), ' images ')]/a");
                    if (dphs.Any())
                        phs.AddRange(dphs);
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' path ')]/a", "", new List<string>() { "Главная" }, "/");
                    size = Helpers.GetItemsInnerText(doc2, "//label[contains(concat(' ', @class, ' '), ' variant_name ')]", "", null, "; ");
                    var data = Helpers.GetItemsAttributt(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/h1", "", "data-product", null, ";");
                    var sizePrice = Helpers.GetItemsAttributt(doc2, "//input[contains(concat(' ', @name, ' '), ' variant ')]", "chg_price(" + data + ",'", "onclick", null, "; ").Replace("')", "").Trim();
                    var photo = "";
                    if (photoCheck.Checked)
                    {
                        var p = Helpers.GetPhoto(doc2, "//img[contains(concat(' ', @src, ' '), '300x300')]", "", "", "",
                            "", "src");
                        photo = Helpers.SavePhoto(p, folder);
                    }
                    if (sizePrice.IndexOf(";") > 0)
                    {
                        var arrPrice = Regex.Split(sizePrice, "; ");
                        var arrSize = Regex.Split(size, "; ");
                         var url=res;
                        for (int i = 0; i < arrPrice.Length; i++)
                        {
                            if (i != 0)
                            {
                                title = desc = url = cat = string.Empty;
                                phs = new List<string>();
                            }
                            products.Add(new Product()
                            {
                                Url = url,
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = arrPrice[i],
                                CategoryPath = cat,
                                Size = arrSize[i],
                                Photos = phs,
                                Photo = photo
                            });
                        }
                    }
                    else
                    {
                        if (size.IndexOf(";") > 0)
                            size = size.Remove(size.IndexOf(";"));
                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs,
                            Photo = photo
                        });
                    }
                    //}
                    //catch (Exception ex) { }
                }

            }

            Helpers.SaveToFile(products, path.Text + @"\Noski.xlsx",photoCheck.Checked);
            StatusStrip("Noski");
            
        }

        private void GetShopNogti(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://shop-nogti.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                if (prod.Count == 0)
                {
                    var doc1 = Helpers.GetHtmlDocument(catalog.Url, "http://shop-nogti.ru", null, cook);
                    var cat1 = Helpers.GetPhoto(doc1, "//tr/td/a", "", "http://shop-nogti.ru", "", "categ");
                    var temp = new List<string>();
                    if (cat1.Any())
                    {
                        foreach (var c1 in cat1)
                        {
                            var tr = Helpers.GetProductLinks(c1, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                            if (tr.Any())
                                temp.AddRange(tr.ToList());
                            else
                            {
                                var doc2 = Helpers.GetHtmlDocument(c1, catalog.Url, null, cook);
                                var cat2 = Helpers.GetPhoto(doc2, "//tr/td/a", "", "http://shop-nogti.ru", "", "categ");
                                if (cat2.Any())
                                {
                                    foreach (var c2 in cat2)
                                    {
                                        var tr2 = Helpers.GetProductLinks(c2, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                                        if (tr2.Any())
                                            temp.AddRange(tr2.ToList());
                                        else
                                        {
                                            var doc3 = Helpers.GetHtmlDocument(c2, c1, null, cook);
                                            var cat3 = Helpers.GetPhoto(doc3, "//tr/td/a", "", "http://shop-nogti.ru", "", "categ");
                                            if (cat3.Any())
                                            {
                                                foreach (var c3 in cat3)
                                                {
                                                    var tr3 = Helpers.GetProductLinks(c3, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                                                    if (tr3.Any())
                                                        temp.AddRange(tr3.ToList());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    prod = new HashSet<string>(temp);
                }
                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//td/h1");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' productPrice ')]").Replace("руб.", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", new List<string>() { "Внимание" }, "\r\n");
                    var tre = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/ul/li", "", null, "\r\n");

                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]");
                    var cd = Regex.Split(res.Replace("http://shop-nogti.ru/shop/details/", ""), "/");
                    if (cd != null)
                    {
                        foreach (var c in cd)
                        {
                            var it = 0;
                            var number = Int32.TryParse(c, out it);
                            if (!number && !c.Contains(".html") && !c.Contains("'"))
                            {
                                var temp = Helpers.GetItemInnerText(doc2, "//a[contains(concat(' ', @href, ' '), '" + c + "')]");
                                if (!string.IsNullOrEmpty(temp))
                                    cat += temp + "/";
                            }
                        }
                        if (cat.Length > 0)
                            cat = cat.Substring(0, cat.Length - 1);
                    }


                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\ShopNogti.xlsx");
            StatusStrip("ShopNogti");
        }

        private void GetTrikotage(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://iv-trikotage.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://iv-trikotage.ru", "//a[contains(concat(' ', @class, ' '), ' fast_pro_title ')]", null);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, "http://iv-trikotage.ru", null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @id, ' '), ' pro_fast_name ')]");
                    var artic = Helpers.GetItemsInnerText(doc2, "//tr/td/div", "Артикул:", null);
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' price ')]");
                    var table = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' detali ')]");
                    desc = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' desc ')]").Replace("Описание:", "").Trim();
                    if (table != null)
                    {
                        foreach (var t in table)
                        {
                            if (t.InnerText.Contains("Артикул"))
                                artic = t.InnerText.Replace("Артикул:", "").Trim();
                        }
                    }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var photos = doc2.DocumentNode.SelectNodes("//li/a/img");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => "http://iv-trikotage.ru" + p.Attributes["src"].Value.Replace("68_","1200_")));
                    else
                    {
                        photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' zoomPup ')]/img");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => p.Attributes["src"].Value.Replace("68_", "1200_")));
                    }
                    if (phs.Any())
                    {
                        phs = phs.Where(x => !x.Contains("home")).ToList();
                    }
                    var cats = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' category ')]"); ;
                    if (cats != null)
                    {
                        foreach (var c in cats)
                        {
                            if (!c.InnerText.Contains("Главная") && !string.IsNullOrEmpty(c.InnerText.Trim()) && !c.InnerText.Contains(title))
                                cat += c.InnerText.Trim() + "/";
                        }
                        if (cat.Length > 0)
                            cat = cat.Substring(0, cat.Length - 1);
                    }

                    var siz = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' size ')]/span[contains(concat(' ', @class, ' '), ' txt_count ')]");
                    if (siz != null)
                    {
                        foreach (var fd in siz)
                        {
                            if (!string.IsNullOrEmpty(fd.InnerText.Trim()))
                                size += fd.InnerText.Trim() + "; ";
                        }
                        size = size.Substring(0, size.Length - 2);
                    }
                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Trikotage.xlsx");
            StatusStrip("Trikotage");
        }

        private void GetGipnozstyle(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://ru.gipnozstyle.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://ru.gipnozstyle.ru", "//div[contains(concat(' ', @class, ' '), ' sm ')]/a", null);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' naimenovanie ')]");
                    var artic = "";
                    var price = "";
                    var table = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' detali ')]");
                    if (table != null)
                    {
                        foreach (var t in table)
                        {
                            if (t.InnerText.Contains("Артикул"))
                                artic = t.InnerText.Replace("Артикул:", "").Trim();
                            else if (t.InnerText.Contains("Цена"))
                                price = t.InnerText.Replace("Цена:", "").Replace("рублей", "").Replace(" ", "").Replace(".", ",").Trim();
                            else if (!string.IsNullOrEmpty(t.InnerText.Trim()) && !t.InnerText.Contains("Цвет") && !t.InnerText.Contains("Размер"))
                                desc += t.InnerText.Trim() + "\r\n";
                        }
                        desc = desc.Trim();
                    }
                    if (desc.Length == 0) { desc = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' opisanie ')]"); }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' sixcol ')]/div/a/img");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => p.Attributes["src"].Value.Replace("_sm","")));
                    var photos2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' sixcol ')]/a/img");
                    if (photos2 != null)
                        phs.AddRange(photos2.Select(p => p.Attributes["src"].Value.Replace("_sm", "")));
                    cat = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' font ')]");

                    var cols = doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @id, ' '), ' colors ')]/option");
                    var siz = doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @id, ' '), ' razmer ')]/option");
                    if (cols != null)
                    {
                        foreach (var fd in cols)
                        {
                            if (!string.IsNullOrEmpty(fd.Attributes["value"].Value.Trim()))
                                col += fd.Attributes["value"].Value.Trim() + "; ";
                        }
                        col = col.Substring(0, col.Length - 2);
                    }
                    if (siz != null)
                    {
                        foreach (var fd in siz)
                        {
                            if (!string.IsNullOrEmpty(fd.Attributes["value"].Value.Trim()))
                                size += fd.Attributes["value"].Value.Trim() + "; ";
                        }
                        size = size.Substring(0, size.Length - 2);
                    }
                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Gipnozstyle.xlsx");
            StatusStrip("Gipnozstyle");
        }

        private void GetWiterra(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://witerra.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "&sort=products_sort_order&page=all", cook, "", "//td[contains(concat(' ', @class, ' '), ' productListing-data ')]/a", null);

                var catal = Helpers.GetProductLinks(catalog.Url, cook, "", "//td[contains(concat(' ', @class, ' '), ' smallText ')]/a", null);
                var temp = new List<string>();
                foreach (var c in catal)
                {
                    var product = Helpers.GetProductLinks(c + "&sort=products_sort_order&page=all", cook, "",
                        "//td[contains(concat(' ', @class, ' '), ' productListing-data ')]/a", null);
                    if (product.Any())
                        temp.AddRange(product);
                    else
                    {
                        var catal2 = Helpers.GetProductLinks(c, cook, "", "//td[contains(concat(' ', @class, ' '), ' smallText ')]/a", null);
                        Thread.Sleep(4000);
                        foreach (var c2 in catal2)
                        {
                            var product2 = Helpers.GetProductLinks(c2 + "&sort=products_sort_order&page=all", cook, "",
                                "//td[contains(concat(' ', @class, ' '), ' productListing-data ')]/a", null);
                            if (product2.Any())
                                temp.AddRange(product2);
                        }
                    }
                    var catal3 = Helpers.GetProductLinks(c, cook, "", "//td[contains(concat(' ', @class, ' '), ' smallText ')]/a", null);
                    foreach (var c3 in catal3)
                    {
                        var product3 = Helpers.GetProductLinks(c3 + "&sort=products_sort_order&page=all", cook, "",
                        "//td[contains(concat(' ', @class, ' '), ' productListing-data ')]/a", null);
                        if (product3.Any())
                            temp.AddRange(product3);
                    }

                }
                temp.AddRange(prod);
                prod = new HashSet<string>(temp);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    try
                    {
                        var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                        var col = "";
                        var size = "";
                        var desc = "";
                        var cat = "";
                        var phs = new List<string>();
                        var title = Helpers.GetItemInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' pageHeading ')]");
                        var artic = "";
                        if (!string.IsNullOrEmpty(title))
                        {
                            artic = title.Substring(title.ToLower().IndexOf("арт") + 3, title.Length - (title.ToLower().IndexOf("арт") + 3)).Replace(".", "").Trim();
                        }
                        if (string.IsNullOrEmpty(artic))
                            artic = title;
                        var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')]", 1).Replace("руб.", "").Replace("Базовая цена:", "").Trim();
                        if (string.IsNullOrEmpty(price))
                            price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')]").Replace("руб.", "").Replace(" ", "").Replace(".", ",").Trim();
                        var desc1 = doc2.DocumentNode.SelectNodes("//td[contains(concat(' ', @class, ' '), ' main ')]/p");
                        if (desc1 != null)
                        {
                            foreach (var d in desc1)
                            {
                                desc += d.InnerText.Trim() + "\r\n";
                            }
                            desc = desc.Trim();
                        }
                        else
                        {
                            desc = Helpers.GetItemInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' main ')]/table",1);
                            if (desc.Contains("<!--"))
                                desc = desc.Remove(desc.IndexOf("Предыдущий")).Trim();
                        }
                        var desc2 = Helpers.GetItemsInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' main ')]/div[contains(concat(' ', @style, ' '), 'justify')]", "", null);
                        if (desc2.Length > 0)
                            desc = (desc + " " + desc2).Trim();

                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' zoom ')]");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => p.Attributes["href"].Value));

                        //cat = HttpUtility.HtmlDecode(catalog.Name);
                        cat = Helpers.GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), ' headerNavigation ')]", "", new List<string>() { "Главная" }, "/");
                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });

                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Witerra.xlsx");
            StatusStrip("Witerra");
        }

        private void GetPiniolo(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.piniolo.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "?page=0", cook, "http://www.piniolo.ru/", "//a[contains(concat(' ', @class, ' '), ' link-pv-name ')]", null);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' product-name ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), ' skuValue ')]");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    var price = Helpers.GetItemInnerText(doc2, "//strong/span/span").Replace("р.", "").Replace("р", "").Replace(" ", "").Replace(".", ",").Trim();
                    if (string.IsNullOrEmpty(price))
                        price = Helpers.GetItemInnerText(doc2, "//p/span/span/span").Replace("р.", "").Replace("р", "").Replace(" ", "").Replace(".", ",").Trim();
                    if(string.IsNullOrEmpty(price))
                        price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' price ')]").Replace("руб.", "").Replace("р", "").Replace(" ", "").Replace(".", ",").Trim();
                    var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' tab-content ')]/p");
                    if (desc1 == null || string.IsNullOrEmpty(desc1[0].InnerHtml.Trim()))
                        desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' tab-content ')]");
                    if (desc1 != null && !string.IsNullOrEmpty(desc1[0].InnerHtml.Trim()))
                    {
                        var temp = Regex.Split(desc1[0].InnerHtml, "<br");
                        foreach (var d in temp)
                        {
                            var t = "";
                            if (d.Contains("style=") && !d.Contains("Цена"))
                            {
                                var tr = d.Substring(d.IndexOf(">") + 1, d.Length - (d.IndexOf(">") + 1));
                                var doc3 = new HtmlAgilityPack.HtmlDocument();
                                doc3.LoadHtml(HttpUtility.HtmlDecode(tr));
                                t = doc3.DocumentNode.InnerText;
                            }
                            if (d.Contains("Размеры"))
                            {
                                if (string.IsNullOrEmpty(t))
                                    size = d.Replace("Размеры:", "").Replace(".", "").Replace("с", "").Replace(">", "").Trim();
                                else
                                    size = t.Replace("Размеры:", "").Replace(".", "").Replace("с", "").Replace(">", "").Trim();
                            }
                            else if (d.Contains("Цвет"))
                            {
                                if (string.IsNullOrEmpty(t))
                                    col = d.Replace("Цвет:", "").Replace(".", "").Replace(">", "").Trim();
                                else
                                    col = t.Replace("Цвет:", "").Replace(".", "").Replace(">", "").Trim();
                            }
                            else if (!d.Contains("Цена") && !string.IsNullOrEmpty(d.Trim()))
                            {
                                if (string.IsNullOrEmpty(t))
                                    desc += d.Replace(">", "").Trim() + "\r\n";
                                else
                                    desc += t.Replace(">", "").Trim() + "\r\n";
                            }
                        }
                        if (desc.Contains("<"))
                            desc = desc.Remove(desc.IndexOf("<"));
                        desc = desc.Trim();
                    }

                    var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @id, ' '), ' zoom ')]");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => "http://www.piniolo.ru/" + p.Attributes["href"].Value));

                    cat = Helpers.GetEncodingCategory(catalog.Name);

                    var table2 =
                        doc2.DocumentNode.SelectNodes("//p/span");
                    if (table2 != null)
                    {
                        foreach (var fd in table2)
                        {
                            if (fd.InnerText.Contains("размеры"))
                                size = fd.InnerText.Replace("размеры:", "").Replace(",", ";").Trim();
                        }
                    }

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Piniolo.xlsx");
            StatusStrip("Piniolo");
        }

        private void GetLemming(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://lemming.su/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://lemming.su", "//p/a", null);
                prod.Remove(prod.ToList()[0]);
                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var client = new System.Net.WebClient();
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                    client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                    client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                    client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                    client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data);
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var price = "";
                    var size = "";
                    var col = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var title = "";
                    var phs = new List<string>();
                    var title1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' item-page ')]/h3");
                    if (title1 != null)
                    {
                        title = title1[0].InnerText.Trim();
                        artic = title;
                    }
                    else
                    {
                        title1 = doc2.DocumentNode.SelectNodes("//p/strong/span");
                        if (title != null)
                        {
                            title = title1[0].InnerText.Trim();
                            artic = title;
                        }
                    }

                    var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' item-page ')]/div/div/div/p[contains(concat(' ', @style, ' '), ' text-align: left; ')]");//[0].InnerText.Trim();
                    if (desc1 != null)
                    {
                        foreach (var d in desc1)
                        {
                            desc += d.InnerText.Trim() + " ";
                        }
                    }
                    else
                    {
                        desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' item-page ')]/p");
                        if (desc1 != null)
                        {
                            foreach (var d in desc1)
                            {
                                if (!d.InnerText.Contains("Расцветки:") && !d.InnerText.Contains("уточняйте по телефону") && !string.IsNullOrEmpty(d.InnerText.Trim()) && !d.InnerText.Contains("Цвет №") && !d.InnerText.Contains("бланк заказа") && !d.InnerText.Contains("<!--"))
                                    desc += d.InnerText.Trim() + " ";
                            }
                        }
                    }
                    //desc = desc.Replace("Расцветки:", "").Replace("", "").Trim();
                    var photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' item-page ')]/p/a");
                    if (photos != null)
                    {
                        foreach (var photo in photos)
                        {
                            if (!photo.InnerText.Contains("var"))
                                phs.Add("http://lemming.su" + photo.Attributes["href"].Value);
                        }
                    }
                    var photos2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' sigplus-gallery ')]/ul/li/a");
                    if (photos2 != null)
                    {
                        foreach (var photo in photos2)
                        {
                            if (!photo.InnerText.Contains("var"))
                                phs.Add("http://lemming.su" + photo.Attributes["href"].Value);
                        }
                    }

                    if (res.Contains("katalog-vesna-osen"))
                        cat = "Весенняя коллекция";
                    else
                        cat = "Зимняя коллекция";

                    var table2 =
                        doc2.DocumentNode.SelectNodes("//p/span");
                    if (table2 != null)
                    {
                        foreach (var fd in table2)
                        {
                            if (fd.InnerText.Contains("размеры"))
                                size = fd.InnerText.Replace("размеры:", "").Replace(",", ";").Trim();
                        }
                    }

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Lemming.xlsx");
            StatusStrip("Lemming");
        }

        private void GetNobi(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.nobi54.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.nobi54.ru", "//div[contains(concat(' ', @class, ' '), ' prod-info ')]/h2/a",
                    "//div[contains(concat(' ', @class, ' '), ' pagination ')]/a", "?page=", null);
                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var client = new System.Net.WebClient();
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                    client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                    client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                    client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                    client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data);
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var price = "";
                    var size = "";
                    var col = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var title = "";
                    var phs = new List<string>();
                    var title1 = doc2.DocumentNode.SelectNodes("//h1[contains(concat(' ', @class, ' '), ' title ')]");
                    if (title1 != null)
                    {
                        title = title1[0].InnerText.Trim();
                        artic = title.Substring(title.IndexOf("арт.") + 4);
                        title = title.Replace(artic, "").Replace("арт.", "");
                    }

                    var price1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @id, ' '), ' total-price ')]");
                    if (price1 != null)
                        price = price1[0].InnerText.Replace("руб", "").Replace(" ", "").Trim();
                    var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' description editor-content ')]/p");//[0].InnerText.Trim();
                    if (desc1 != null)
                    {
                        foreach (var d in desc1)
                        {
                            if (!string.IsNullOrEmpty(d.InnerText.Trim()))
                                desc += d.InnerText.Trim() + ". ";
                        }
                        desc = desc.Replace("..", ".");
                    }

                    var photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' photo ')]/a");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => p.Attributes["href"].Value));

                    var win = Encoding.GetEncoding("windows-1251");
                    byte[] winBytes = win.GetBytes(catalog.Name);
                    cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);

                    var table2 =
                        doc2.DocumentNode.SelectNodes("//select");
                    if (table2 != null)
                    {
                        var fds = Regex.Split(table2[0].InnerHtml, ">");
                        foreach (var fd in fds)
                        {
                            var begin = fd.IndexOf("<");
                            if (begin < 0)
                                begin = fd.Length;
                            var temp = fd.Substring(0, begin).Trim();
                            if (!string.IsNullOrEmpty(temp) && !temp.Contains("Выбер"))
                                size += temp + "; ";
                        }
                        size = size.Substring(0, size.Length - 2);
                    }

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });

                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Nobi.xlsx");
            StatusStrip("Nobi");
        }

        private void GetNaksa(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://naksa.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var client2 = new System.Net.WebClient();
                client2.Headers.Add(HttpRequestHeader.Cookie, cook);
                client2.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client2.Headers.Add(HttpRequestHeader.Referer, "http://naksa.ru/");
                client2.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client2.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data2 = client2.OpenRead(catalog.Url);
                var reader2 = new StreamReader(data2, Encoding.UTF8);
                string s2 = reader2.ReadToEnd();
                data2.Close();
                reader2.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s2);
                var prLink = new List<string>();
                var descList = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' prdbrief_name ')]/a");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://naksa.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                        var doc3 = new HtmlAgilityPack.HtmlDocument();
                        doc3.LoadHtml(p.ParentNode.ParentNode.InnerHtml);
                        var temp = doc3.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' prdbrief_brief_description ')]");
                        if (temp != null)
                            descList.Add(temp[0].InnerText);
                        else
                            descList.Add(" ");
                    }

                    var pages = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' no_underline ')]");
                    if (pages != null && pages.Count > 1)
                    {
                        var preLink = new List<string>();
                        foreach (var pag in pages)
                        {
                            var link2 = WebUtility.HtmlDecode(pag.Attributes["href"].Value);
                            if (!preLink.Contains(link2))
                            {
                                var web2 = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load("http://naksa.ru" + link2);
                                var a2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' prdbrief_name ')]/a");
                                if (a2 != null)
                                {
                                    foreach (var p in a2)
                                    {
                                        prLink.Add("http://naksa.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                        var doc3 = new HtmlAgilityPack.HtmlDocument();
                                        doc3.LoadHtml(p.ParentNode.ParentNode.InnerHtml);
                                        var temp = doc3.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' prdbrief_brief_description ')]");
                                        if (temp != null)
                                            descList.Add(temp[0].InnerText);
                                        else
                                            descList.Add(" ");
                                    }
                                }
                                preLink.Add(link2);
                            }
                        }
                    }
                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;
                int z = 0;
                foreach (var res in prod)
                {
                    try
                    {

                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var price = "";
                        var size = "";
                        var col = "";
                        var desc = "";
                        var cat = "";
                        var artic = "";
                        var title = "";
                        var phs = new List<string>();
                        var title1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' cpt_product_name ')]/h1");
                        if (title1 != null)
                            title = title1[0].InnerText.Trim();
                        var artic1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' productCode ')]");
                        if (artic1 != null)
                            artic = artic1[0].InnerText.Trim();
                        else
                            artic = title;
                        var price1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' totalPrice ')]");
                        if (price1 != null)
                            price = price1[0].InnerText.Replace("руб.", "").Replace(".", ",").Trim();
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' cpt_product_description ')]");
                        if (desc1 != null)
                            desc = desc1[0].InnerText.Trim();
                        else
                            desc = descList[z];
                        var photos = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @id, ' '), ' img-current_picture ')]");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => "http://naksa.ru" + p.Attributes["src"].Value));

                        var cats = doc2.DocumentNode.SelectNodes("//table/tr/td/a");
                        if (cats != null)
                        {
                            foreach (var cat1 in cats)
                            {
                                if (!cat1.InnerText.Contains("Главная") && !string.IsNullOrEmpty(cat1.InnerText.Trim()) && !cat1.InnerText.Contains("печати"))
                                    cat += cat1.InnerText.Trim() + "/";
                            }
                            cat = cat.Substring(0, cat.Length - 1);
                        }

                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });
                        z++;
                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Haksa.xlsx");
            StatusStrip("Haksa");
        }

        private void GetOptEconom(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.opt-ekonom.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var client2 = new System.Net.WebClient();
                client2.Headers.Add(HttpRequestHeader.Cookie, cook);
                client2.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client2.Headers.Add(HttpRequestHeader.Referer, "http://www.opt-ekonom.ru/");
                client2.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client2.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data2 = client2.OpenRead(catalog.Url);
                var reader2 = new StreamReader(data2, Encoding.UTF8);
                string s2 = reader2.ReadToEnd();
                data2.Close();
                reader2.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s2);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' button_detail ')]");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://www.opt-ekonom.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                    var pages = doc.DocumentNode.SelectNodes("//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a");
                    if (pages != null && pages.Count > 1)
                    {
                        var preLink = new List<string>();
                        foreach (var pag in pages)
                        {
                            var link2 = WebUtility.HtmlDecode(pag.Attributes["href"].Value);
                            if (!preLink.Contains(link2))
                            {
                                var web2 = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load("http://www.opt-ekonom.ru" + link2);
                                var a2 = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' button_detail ')]");
                                if (a2 != null)
                                {
                                    foreach (var p in a2)
                                    {
                                        prLink.Add("http://www.opt-ekonom.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                    }
                                }
                                preLink.Add(link2);
                            }
                        }
                    }

                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    try
                    {
                    if (countRequest % 8 == 0)
                    {
                        Thread.Sleep(5000);
                    }
                    var client = new System.Net.WebClient();
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                    client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                    client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                    client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                    client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data);
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var price = "";
                    var size = "";
                    var col = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var title = "";
                    var phs = new List<string>();
                    var title1 = doc2.DocumentNode.SelectNodes("//h1");
                    if (title1 != null)
                        title = title1[1].InnerText.Trim();
                    var artic1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @id, ' '), ' product_code ')]");
                    if (artic1 != null)
                    {
                        artic = artic1[0].InnerText.Trim();
                        title = title.Replace("(Код: " + artic + ")", "");
                    }
                    else
                        artic = title;
                    var price1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @id, ' '), ' block_price ')]");
                    if (price1 != null)
                        price = price1[0].InnerText.Replace("RUB", "").Replace(".", ",").Trim();
                    var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' jshop_prod_description ')]");//[0].InnerText.Trim();
                    if (desc1 != null)
                    {
                        var tb = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' jshop_prod_description ')]/table");
                        var tb2 = "";
                        if (tb != null)
                            tb2 = tb[0].InnerText;
                        if (!string.IsNullOrEmpty(tb2))
                            desc += desc1[0].InnerText.Replace(tb2, "").Trim();
                        else
                            desc += desc1[0].InnerText.Trim();
                        desc = desc.Replace("Таблица размеров", "").Trim();
                        if (desc.Contains("<!--"))
                        {
                            desc = desc.Remove(desc.IndexOf("<!--")).Trim();
                        }
                    }
                    var photos = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @class, ' '), ' jshop_img_thumb ')]");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => p.Attributes["src"].Value.Replace("thumb", "full")));
                    else
                    {
                        photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' lightbox ')]");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => p.Attributes["href"].Value));
                    }
                    var win = Encoding.GetEncoding("windows-1251");
                    byte[] winBytes = win.GetBytes(catalog.Name);
                    cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);
                    var table2 =
                        doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @id, ' '), ' jshop_attr_id1 ')]");
                    if (table2 != null)
                    {
                        var fds = Regex.Split(table2[0].InnerHtml, ">");
                        foreach (var fd in fds)
                        {
                            var begin = fd.IndexOf("<");
                            if (begin < 0)
                                begin = fd.Length;
                            var temp = fd.Substring(0, begin).Trim();
                            if (!string.IsNullOrEmpty(temp) && !temp.Contains("Выбер"))
                                size += temp + "; ";
                        }
                        size = size.Substring(0, size.Length - 2);
                    }

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs
                    });
                    countRequest++;
                    }
                    catch (Exception ex) { Thread.Sleep(8000); }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\OptEconom.xlsx");
            StatusStrip("OptEconom");
        }

        private void GetAdel(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://td-adel.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var client2 = new System.Net.WebClient();
                client2.Headers.Add(HttpRequestHeader.Cookie, cook);
                client2.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client2.Headers.Add(HttpRequestHeader.Referer, "http://td-adel.ru/");
                client2.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client2.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data2 = client2.OpenRead(catalog.Url);
                var reader2 = new StreamReader(data2, Encoding.UTF8);
                string s2 = reader2.ReadToEnd();
                data2.Close();
                reader2.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s2);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' spacer-link ')]");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://td-adel.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                    var a2 = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', @id, ' '), ' more-list1 ')]");
                    if (a2 != null)
                    {
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead("http://td-adel.ru" + a2[0].Attributes["href"].Value);
                        var reader = new StreamReader(data, Encoding.UTF8);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(s);
                        var aa = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' spacer-link ')]");
                        if (aa != null)
                        {
                            foreach (var p in aa)
                            {
                                prLink.Add("http://td-adel.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                            }
                        }
                    }
                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var price = "";
                        var size = "";
                        var col = "";
                        var desc = "";
                        var cat = "";
                        var artic = "";
                        var title = "";
                        var phs = new List<string>();
                        var title1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' tovar-card__left ')]/h1");
                        if (title1 != null)
                            title = title1[0].InnerText.Trim();
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' tovar-card__paramets ')]");//[0].InnerText.Trim();
                        var desc2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' tovar-card__text ')]");
                        if (desc2 != null && !string.IsNullOrEmpty(desc2[0].InnerText.Replace("При сумме заказа от 50000 рублей предоставляется скидка 3%", "")
                                            .Replace("При сумме заказа от 100000 рублей предоставляется скидка 5%", "").Trim()))
                            desc += desc2[0].InnerText.Trim() + ". ";
                        if (desc1 != null)
                            desc += desc1[0].InnerText.Replace("Комментарий:", "").Trim();
                        //if (artic1 != null)
                        //    artic = artic1[0].InnerText.Replace("Код:", "").Trim();
                        //else
                        desc = desc.Replace("При сумме заказа от 50000 рублей предоставляется скидка 3%", "")
                                            .Replace("При сумме заказа от 100000 рублей предоставляется скидка 5%", "");
                        artic = title;
                        var price1 = doc2.DocumentNode.SelectNodes("//p[contains(concat(' ', @class, ' '), ' b-product__price ')]");
                        if (price1 != null)
                            price = price1[0].InnerText.Replace("руб.", "").Replace(" ", "").Trim();

                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' carousel-href ')]");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => "http://td-adel.ru" + p.Attributes["href"].Value));
                        var cat1 = catalog.Url.Replace("http://td-adel.ru/catalog/", "");
                        var cat2 = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @href, ' '), '" + cat1 + "')]");
                        if (cat2 != null)
                        {
                            cat = cat2[0].InnerText.Trim();
                            var cts = cat2[0].ParentNode.ParentNode.ParentNode.ChildNodes["a"];

                            if (cts != null)
                            {
                                cat = cts.InnerText.Trim() + "/" + cat;
                                var par = cts.ParentNode.ParentNode.ParentNode.ChildNodes["a"];
                                if (par != null)
                                {
                                    cat = par.InnerText.Trim() + "/" + cat;
                                    var pr2 = par.ParentNode.ParentNode.ParentNode.ChildNodes["a"];
                                    if (pr2 != null)
                                        cat = pr2.InnerText.Trim() + "/" + cat;
                                }
                            }

                        }
                        var table =
                            doc2.DocumentNode.SelectNodes("//li[contains(concat(' ', @class, ' '), ' li_mult ')]");
                        if (table != null)
                        {
                            foreach (var tr in table)
                            {
                                if (tr.InnerText.Trim() != string.Empty)
                                {
                                    var title2 = "";
                                    var desc3 = "";
                                    var tit1 =
                                        doc2.DocumentNode.SelectNodes(
                                            "//div[contains(concat(' ', @class, ' '), ' tovar-card__title ')]");
                                    if (tit1 != null)
                                    {
                                        title2 = tit1[0].InnerText.Trim();
                                        if (title2.ToLower().Contains("размер"))
                                            size = title2.ToLower().Replace("размер", "").Trim();
                                    }
                                    if (!string.IsNullOrEmpty(title.Trim()))
                                        title2 = title + " " + title2;
                                    artic = title2;
                                    var des =
                                        doc2.DocumentNode.SelectNodes(
                                            "//div[contains(concat(' ', @class, ' '), ' tovar-card__text1 ')]");
                                    if (des != null)
                                        desc3 = des[0].InnerText.Trim();
                                    if (!string.IsNullOrEmpty(desc.Trim()))
                                        desc3 = desc + ". " + desc3;
                                    desc3 =
                                        desc3.Replace("..", ".")
                                            .Replace("\n", "")
                                            .Replace("\t", "")
                                            .Replace("  ", "").Trim();
                                    var pr =
                                        doc2.DocumentNode.SelectNodes(
                                            "//input[contains(concat(' ', @name, ' '), ' price ')]");
                                    if (pr != null)
                                        price = pr[0].Attributes["value"].Value.Replace(".", ",").Trim();
                                    products.Add(new Product()
                                    {
                                        Url = res,
                                        Article = artic,
                                        Color = col,
                                        Description = desc3,
                                        Name = title2,
                                        Price = price,
                                        CategoryPath = cat,
                                        Size = size,
                                        Photos = phs
                                    });
                                }
                            }
                        }
                        else
                        {
                            price = Helpers.GetItemInnerText(doc2,
                                "//div[contains(concat(' ', @class, ' '), ' tovar-card__price ')]").Replace(".-", "").Trim();
                            products.Add(new Product()
                            {
                                Url = res,
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = price,
                                CategoryPath = cat,
                                Size = size,
                                Photos = phs
                            });
                        }



                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Adel.xlsx");
            StatusStrip("Adel");
        }

        private void GetArtvision(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://artvision-opt.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                Thread.Sleep(10000);
                var client2 = new System.Net.WebClient();
                client2.Headers.Add(HttpRequestHeader.Cookie, cook);
                client2.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client2.Headers.Add(HttpRequestHeader.Referer, "http://artvision-opt.ru/");
                client2.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client2.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data2 = client2.OpenRead(catalog.Url);
                var reader2 = new StreamReader(data2, Encoding.UTF8);
                string s2 = reader2.ReadToEnd();
                data2.Close();
                reader2.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s2);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' b-product-line__product-name-link ')]");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://artvision-opt.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }

                    var pages = doc.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), '  b-pager__link  ')]");
                    if (pages != null && pages.Count > 1)
                    {
                        var countLink = Convert.ToInt32(pages[1].InnerText.Trim());
                        for (var i = 2; i <= countLink; i++)
                        {
                            if (i % 3 == 0)
                            {
                                Thread.Sleep(8000);
                            }
                            var web2 = new HtmlWeb();
                            HtmlAgilityPack.HtmlDocument doc2 = web2.Load(catalog.Url + "/page_" + i);
                            var a2 = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' b-product-line__product-name-link ')]");
                            if (a2 != null)
                            {
                                foreach (var p in a2)
                                {
                                    prLink.Add("http://artvision-opt.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                }
                            }
                        }
                    }
                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;
                int countRequest = 0;
                foreach (var res in prod)
                {
                    try
                    {
                        if (countRequest % 3 == 0)
                        {
                            Thread.Sleep(11000);
                        }
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var price = "";
                        var size = "";
                        var col = "";
                        var desc = "";
                        var cat = "";
                        var artic = "";
                        var title = "";
                        var phs = new List<string>();
                        var title1 = doc2.DocumentNode.SelectNodes("//h1[contains(concat(' ', @class, ' '), ' b-product__name ')]");
                        if (title1 != null)
                            title = title1[0].InnerText.Trim();
                        var artic1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' b-product__sku ')]");
                        if (artic1 != null)
                            artic = artic1[0].InnerText.Replace("Код:", "").Trim();
                        else
                            artic = title;
                        var price1 = doc2.DocumentNode.SelectNodes("//p[contains(concat(' ', @class, ' '), ' b-product__price ')]");
                        if (price1 != null)
                            price = price1[0].InnerText.Replace("руб.", "").Replace(" ", "").Replace("Оптовыеирозничныеценыуточняйте", "").Trim();
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' b-user-content ')]");//[0].InnerText.Trim();
                        if (desc1 != null)
                            desc += desc1[0].InnerText.Trim();
                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' b-centered-image ')]");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => p.Attributes["href"].Value));
                        var cat1 = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' b-sidebar-menu__sublink_type_active ')]");
                        var cat2 = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' b-breadcrumb__link ')]");
                        if (cat1 != null)
                        {
                            cat = cat1[0].InnerText.Trim();
                            foreach (var c in cat2)
                            {
                                if (!c.InnerText.Contains("...") && !c.InnerText.Contains("Оптовый интернет-магазин") &&
                                    !c.InnerText.Contains("Товары и услуги") && !c.InnerText.Contains(cat))
                                    cat += "/" + c.InnerText.Trim();
                            }
                            cat = cat.Replace(".", "");
                        }
                        var table =
                            doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' b-product-info ')]/tr");
                        if (table != null)
                        {
                            foreach (var tr in table)
                            {
                                if (tr.InnerText.Contains("Размер"))
                                    size = tr.InnerText.Replace("Размер", "").Trim();
                                else if (tr.InnerText.Contains("Цвет"))
                                    col = tr.InnerText.Replace("Цвет", "").Trim();
                            }
                        }

                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });
                        countRequest++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Слишком частое обращение к северу http://artvision-opt.ru. Для исправления ошибки пройдите на указанный сайт и введите капчу. После разблокировки сайта нажмите кнопку ОК",
                                        "Предупреждение!!!",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Exclamation,
                                        MessageBoxDefaultButton.Button1);
                    }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Artvision.xlsx");
            StatusStrip("Artvision");
        }

        private void GetSportOpt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://xn----0tbbbddeld.xn--p1ai/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var client2 = new System.Net.WebClient();
                client2.Headers.Add(HttpRequestHeader.Cookie, cook);
                client2.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client2.Headers.Add(HttpRequestHeader.Referer, "http://xn----0tbbbddeld.xn--p1ai/");
                client2.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client2.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data2 = client2.OpenRead(catalog.Url + "?SHOWALL_1=1");
                var reader2 = new StreamReader(data2, Encoding.UTF8);
                string s2 = reader2.ReadToEnd();
                data2.Close();
                reader2.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s2);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' catalog-item-title ')]/a");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://xn----0tbbbddeld.xn--p1ai" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//h1[contains(concat(' ', @id, ' '), ' pagetitle ')]")[0].InnerText.Trim();
                        var artic = title;
                        var price = "";
                        var price1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' catalog-item-price ')]");
                        if (price1 != null)
                            price = price1[0].InnerText.Replace("руб", "").Replace(" ", "").Trim();
                        var size = "";
                        var col = "";
                        var desc = "";
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' catalog-detail-full-desc ')]");//[0].InnerText.Trim();
                        if (desc1 != null)
                        {
                            desc += desc1[0].InnerText.Replace("Описание", "").Trim();
                        }
                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @rel, ' '), ' catalog-detail-images ')]");
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                phs.Add("http://xn----0tbbbddeld.xn--p1ai" + p.Attributes["href"].Value);
                            }
                        }

                        var cat = "";
                        var win = Encoding.GetEncoding("windows-1251");
                        byte[] winBytes = win.GetBytes(catalog.Name);
                        cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);
                        var cats = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' selected ')]")[0].InnerText.Trim();
                        if (!cat.Equals(cats))
                        {
                            cat = cats + "/" + cat;
                        }
                        cat = HttpUtility.HtmlDecode(cat);
                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });
                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\SportOpt.xlsx");
            StatusStrip("SportOpt");
        }

        private void GetNashipupsi(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://nashipupsi.ru", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://nashipupsi.ru", "//p[contains(concat(' ', @class, ' '), ' product-name ')]/a",
                    "//div[contains(concat(' ', @class, ' '), ' shop2-pageist ')]/a", null);
                if (prod.Count == 0)
                    continue;
                int countRequest = 0;
                foreach (var res in prod)
                {
                    try
                    {
                        if (countRequest % 8 == 0)
                        {
                            Thread.Sleep(5000);
                        }
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' h_rb ')]")[0].InnerText.Trim();
                        var artic = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' product-code ')]/span")[0].InnerText.Trim();
                        if (artic.Contains("нет"))
                            artic = title;
                        var price = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' product-price ')]/span")[0].InnerText.Replace(" ", "").Trim();
                        var size = "";
                        var col = "";
                        var desc = "";
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' tabs-1 ')]/p");//[0].InnerText.Trim();
                        if (desc1 != null)
                        {
                            foreach (var d in desc1)
                            {
                                if (!string.IsNullOrEmpty(d.InnerText.Trim()))
                                    desc += d.InnerText + "\r\n ";
                            }
                            desc = desc.Trim();
                        }
                        else
                        {
                            desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' tabs-1 ')]");
                            if (desc1 != null)
                            {
                                foreach (var d in desc1)
                                {
                                    if (!string.IsNullOrEmpty(d.InnerText.Trim()))
                                        desc += d.InnerText + "\r\n";
                                }
                                desc = desc.Trim();
                            }
                        }
                        var photos = doc2.DocumentNode.SelectNodes("//ul/li/img");
                        var phs = new List<string>();
                        if (photos == null)
                        {
                            photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' product-image ')]/a");
                            if (photos != null)
                            {
                                foreach (var p in photos)
                                {
                                    phs.Add("http://nashipupsi.ru" + p.Attributes["href"].Value);
                                }
                            }
                        }
                        else
                        {
                            foreach (var p in photos)
                            {
                                var t1 = p.Attributes["onclick"].Value;
                                var begin = t1.IndexOf("/d/");
                                var result = t1.Substring(begin, t1.IndexOf("'", begin + 1) - begin);
                                phs.Add("http://nashipupsi.ru" + result);
                            }
                        }
                        var cat = "";
                        var cats = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' site-path ')]/a");
                        if (cats != null)
                        {
                            foreach (var c in cats)
                            {
                                if (!c.InnerText.Contains("Главная") && !c.InnerText.Contains("Интернет-магазин"))
                                    cat += c.InnerText.Trim() + "/";
                            }
                            if (cat.Length > 0)
                                cat = cat.Remove(cat.Length - 1);
                        }

                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });
                        countRequest++;
                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Nashipupsi.xlsx");
            StatusStrip("Nashipupsi");
        }

        private void GetSportoptovik(IEnumerable<Category> list)
        {
            var products = new List<Sport>();
            foreach (var catalog in list)
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(catalog.Url);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' block2 ')]/h2/a");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://www.sportoptovik.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }

                    var pages =
                        doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' peid ')]/ul/li/a");
                    if (pages != null && pages.Count > 1)
                    {
                        var preLink = new List<string>();
                        foreach (var pag in pages)
                        {
                            var link2 = WebUtility.HtmlDecode(pag.Attributes["href"].Value);
                            if (!preLink.Contains(link2))
                            {
                                var web2 = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load("http://www.sportoptovik.ru" + link2);
                                var a2 =
                                    doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' block2 ')]/h2/a");
                                if (a2 != null)
                                {
                                    foreach (var p in a2)
                                    {
                                        prLink.Add("http://www.sportoptovik.ru" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                    }
                                }
                                preLink.Add(link2);
                            }
                        }
                    }
                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;
                var categ1 = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' path ')]/a");
                var cat1 = "";
                if (categ1 != null && categ1.Count == 3)
                    cat1 = categ1[2].InnerText.Trim();
                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data, Encoding.ASCII);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' page_title ')]/h1")[0].InnerText.Trim();
                        var artic = title;
                        var price1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' goods_main ')]/div/p");
                        var price = "";
                        var price2 = "";
                        var price3 = "";
                        var price4 = "";
                        if (price1 != null)
                        {
                            //price = price1[0].InnerText.Replace("руб", "").Replace("ОПТ1 -", "").Replace(".", ",").Trim();
                            foreach (var p in price1)
                            {
                                if (p.InnerText.Contains("ОПТ1"))
                                {
                                    price = p.InnerText.Replace("руб", "").Replace("ОПТ1 -", "").Replace("\r", "").Trim();
                                }
                                else if (p.InnerText.Contains("ОПТ2"))
                                    price2 = p.InnerText.Replace("руб", "").Replace("ОПТ2 -", "").Replace("\r", "").Trim();
                                else if (p.InnerText.Contains("ОПТ3"))
                                    price3 = p.InnerText.Replace("руб", "").Replace("ОПТ3 -", "").Replace("\r", "").Trim();
                                else if (p.InnerText.Contains("ОПТ4"))
                                    price4 = p.InnerText.Replace("руб", "").Replace("ОПТ4 -", "").Replace("\r", "").Trim();
                            }
                            //[0].InnerText.Replace("руб", "").Replace("ОПТ1 -", "").Replace(".",",").Trim();
                        }
                        var temp = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' goods_main_descr ')]");
                        var size = "";
                        var col = "";
                        var desc = "";
                        if (temp != null)
                        {
                            var doc3 = new HtmlAgilityPack.HtmlDocument();
                            doc3.LoadHtml(temp[0].InnerHtml);
                            temp = doc3.DocumentNode.SelectNodes("//p");
                            foreach (var s1 in temp)
                            {
                                if (s1.InnerText.Contains("Цвет:"))
                                {
                                    var t1 = s1.InnerText.Substring(s1.InnerText.IndexOf("Цвет:"), s1.InnerText.Length - s1.InnerText.IndexOf("Цвет:"));
                                    col = t1.Replace("Цвет:", "").Trim();
                                }
                                else if (s1.InnerText.Contains("Размеры"))
                                {
                                    if (s1.InnerText.Contains("Размеры:"))
                                        size = s1.InnerText.Replace("Размеры:", "").Trim();
                                    else if (string.IsNullOrEmpty(size))
                                    {
                                        var begin = s1.InnerText.IndexOf(": ", s1.InnerText.IndexOf("Размеры")) + 1;
                                        var t1 = s1.InnerText.Substring(begin + 1, s1.InnerText.Length - begin);
                                        size = t1.Replace("Размеры:", "").Trim();
                                    }
                                }
                                else if (!s1.InnerText.Contains("if") && !string.IsNullOrEmpty(s1.InnerText.Trim())&&!s1.InnerText.Contains("Название"))
                                    desc += s1.InnerText.Trim() + ". ";
                            }
                            desc = desc.Replace("..", ".");
                        }

                        var photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' goods_main_img ')]/ul/li/a/img");
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                phs.Add("http://www.sportoptovik.ru" + p.Attributes["src"].Value);
                            }
                        }
                        var photos2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' img_bg ')]/img");
                        if (photos2 != null)
                        {
                            phs.Add("http://www.sportoptovik.ru" + photos2[0].Attributes["src"].Value);
                        }

                        //var cat = doc2.DocumentNode.SelectNodes("//ul/li/a/strong")[0].InnerText;
                        var cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' path ')]/a",
                            "", new List<string>() {"Главна"}, "/");
                        //if (!string.IsNullOrEmpty(cat1))
                        //    cat = cat1 + "/" + cat;

                        products.Add(new Sport()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            Price2 = price2,
                            Price3 = price3,
                            Price4 = price4,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });
                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Sportoptovik.xlsx");
            StatusStrip("Sportoptovik");
        }

        private void GetRoomdecor(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://roomdecor.su", new NameValueCollection());
            foreach (var catalog in list)
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(catalog.Url);
                var prLink = new List<string>();
                var last = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' wpsc_page_numbers_bottom ')]/a");
                var numberLastLink = 0;
                if (last != null)
                {
                    int i = 0;
                    foreach (var tt in last)
                    {
                        if (tt.InnerText.Contains("Последняя"))
                            numberLastLink = i;
                        i++;
                    }
                }
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "", "//h2[contains(concat(' ', @class, ' '), ' prodtitle ')]/a", "//div[contains(concat(' ', @class, ' '), ' wpsc_page_numbers_bottom ')]/a", "&paged=", null, numberLastLink);
                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' productcol ')]/h2")[0].InnerText;
                        var art = Regex.Split(
                            doc2.DocumentNode.SelectNodes(
                                "//div[contains(concat(' ', @class, ' '), ' product_description ')]/p")[0].InnerHtml,
                            "<br>");
                        var artic = "";
                        var size = "";
                        var col = "";
                        var desc = "";
                        foreach (var s1 in art)
                        {
                            if (s1.Contains("Артикул:"))
                                artic = s1.Replace("Артикул:", "").Trim();
                            else if (s1.Contains("Размеры:"))
                                size = s1.Replace("Размеры:", "").Trim();
                            else
                                desc += s1.Trim() + "\r\n";
                        }
                        if (string.IsNullOrEmpty(artic))
                            artic = title;
                        if (string.IsNullOrEmpty(desc.Trim()))
                        {
                            var t = doc2.DocumentNode.SelectNodes(
                                "//div[contains(concat(' ', @class, ' '), ' product_description ')]");
                            desc = t[0].InnerText.Trim();
                        }
                        var pr =
                            doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' currentprice ')]");
                        var price = "";
                        if (pr != null)
                            price = pr[0].InnerText.Replace("RUB", "").Replace(".", ",").Trim();
                        else
                        {
                            var t = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' wpsc_product_price ')]/input");
                            price = t[0].Attributes["value"].Value.Replace("RUB", "").Replace(".", ",").Trim();
                        }
                        var photos =
                            doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' imagecol ')]/a");
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                phs.Add(p.Attributes["href"].Value);
                            }
                        }

                        var win = Encoding.GetEncoding("windows-1251");
                        byte[] winBytes = win.GetBytes(catalog.Name);
                        var cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);

                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs
                        });
                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Roomdecor.xlsx");
            StatusStrip("Roomdecor");
        }

        private void GetAventum(IEnumerable<Category> list)
        {
            var products = new List<Product>();

            foreach (var catalog in list)
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(catalog.Url);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' prdbrief_name ')]/a");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add("http://aventum.cz" + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                }
                var prod = new HashSet<string>(prLink);
                if (prod.Count == 0)
                    continue;
                var cook = Helpers.GetCookiePost(prLink[0], new NameValueCollection() { { "current_currency", "3" } });
                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                        client.Headers.Add(HttpRequestHeader.Referer, catalog.Url);
                        client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                        client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' cpt_product_name ')]/h1")[0].InnerText;
                        var artic = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' productCode ')]")[0].InnerText;
                        var pr = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' totalPrice ')]");
                        var price = "";
                        if (pr != null)
                            price = pr[0].InnerText.Replace("руб.", "").Trim();
                        var photos = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @id, ' '), ' img-current_picture ')]");
                        var photos2 = doc2.DocumentNode.SelectNodes("//tr/td/a");
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                phs.Add("http://aventum.cz" + p.Attributes["src"].Value.Replace("_med",""));
                            }
                        }
                        if (photos2 != null)
                        {
                            foreach (var p in photos2)
                            {
                                phs.Add("http://aventum.cz" + p.Attributes["href"].Value);
                            }
                        }
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' cpt_product_description ')]/div/p");
                        var desc = "";
                        if (desc1 != null)
                        {
                            foreach (var d in desc1)
                            {
                                desc += d.InnerText.Trim() + "\r\n";
                            }
                            desc = desc.Substring(0, desc.Length - 4);
                        }
                        else
                        {
                            desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' cpt_product_description ')]/div");
                            if (desc1 != null)
                            {
                                foreach (var d in desc1)
                                {
                                    desc += d.InnerText.Trim() + "\r\n";
                                }
                                desc = desc.Trim();
                            }
                        }
                        var size = doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @name, ' '), ' opt_razmer ')]")[0].InnerText.Trim().Replace("\t", "").Replace("\r\n", ";").Replace(";;", "");

                        var color = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' cpt_product_params_fixed ')]/table/tr");
                        var col = "";
                        if (color != null)
                        {
                            foreach (var c in color)
                            {
                                if (c.InnerText.Contains("Цвет:"))
                                    col += c.InnerText.Replace("Цвет:", "").Trim() + "; ";
                            }
                            col = col.Substring(0, col.Length - 2);
                        }

                        var cat = Helpers.GetEncodingCategory(catalog.Name);

                        products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, CategoryPath = cat, Size = size, Photos = phs });
                    }
                    catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Aventum.xlsx");
            StatusStrip("Aventum");
        }

        private void GetButterfly(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost(list.ToList()[0].Url, new NameValueCollection() { { "submit", "%D0%98%D0%B7%D0%BC%D0%B5%D0%BD%D0%B8%D1%82%D1%8C%20%D0%B2%D0%B0%D0%BB%D1%8E%D1%82%D1%83" }, { "virtuemart_currency_id", "131" } });
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "/results,1-150", cook, "http://butterfly-dress.com", "//div[contains(concat(' ', @class, ' '), ' bfvmcatimage ')]/a",
                    "//a[contains(concat(' ', @class, ' '), ' pagenav ')]", null);
                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        client.Headers.Add(HttpRequestHeader.Cookie, cook);
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//td[contains(concat(' ', @class, ' '), ' bfvmdetailsright ')]/h1")[0].InnerText;
                        var artic = title;

                        var price = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' PricesalesPrice ')]")[0].InnerText.Replace("руб", "").Replace("Цена:", "").Trim();
                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' modal ')]");
                        //main-image img
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                if (!p.Attributes["href"].Value.Contains("recommend?tmpl=component"))
                                    phs.Add(p.Attributes["href"].Value);
                            }
                            if (phs.Count == 0)
                            {
                                var ph = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @id, ' '), ' medium-image ')]");
                                if (ph != null)
                                    phs.Add("http://butterfly-dress.com" + ph[0].Attributes["src"].Value);
                            }
                        }
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' main-image ')]/p");
                        var desc = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-description ')]").Replace("Описание", "").Trim();
                        if (desc1 != null)
                            desc = (desc + " " + desc1[0].InnerText.Trim()).Trim();
                        var siz = "";
                        var size = doc2.DocumentNode.SelectNodes("//ызфт[contains(concat(' ', @class, ' '), ' product-field-display')]/select/option");
                        var col = "";
                        if (size != null)
                        {
                            foreach (var c in size)
                            {
                                siz += c.InnerText.Trim() + "; ";
                            }
                            siz = siz.Substring(0, siz.Length - 2);
                        }

                        var cat = "";
                        var cats = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/a");
                        foreach (var c in cats)
                        {
                            var temp = c.InnerText;
                            if (!temp.Contains("Главная"))
                            {
                                cat += temp.Replace("Главная", "").Trim() + "/";
                            }
                        }
                        if (cat.Length > 0)
                            cat = cat.Substring(0, cat.Length - 1);
                        products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, CategoryPath = cat, Size = siz, Photos = phs });
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            Helpers.SaveToFile(products, path.Text + @"\Butterfly.xlsx");
            StatusStrip("Butterfly");
        }

        private void GetTrikbel(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            foreach (var catalog in list)
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(catalog.Url);
                var prLink = new List<string>();

                var a = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' name ')]/a");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }

                    var pages =
                        doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' links ')]/a");
                    if (pages != null && pages.Count > 1)
                    {
                        var preLink = new List<string>();
                        foreach (var pag in pages)
                        {
                            var link2 = WebUtility.HtmlDecode(pag.Attributes["href"].Value);
                            if (!preLink.Contains(link2))
                            {
                                var web2 = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load(link2);
                                var a2 =
                                    doc2.DocumentNode.SelectNodes(
                                        "//div[contains(concat(' ', @class, ' '), ' name ')]/a");

                                foreach (var p in a2)
                                {
                                    prLink.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                }
                                preLink.Add(link2);
                            }
                        }
                    }
                }
                var prod = new HashSet<string>(prLink);

                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        Stream data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' content ')]/h1")[0].InnerText;
                        var artic = title;

                        var price = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' price ')]")[0].InnerText.Replace("р.", "").Replace("Цена:", "").Trim();
                        var price2 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' price-new ')]");
                        if (price2 != null)
                            price = price2[0].InnerText.Replace("р.", "").Replace("Цена:", "").Trim();
                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' colorbox ')]");
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                phs.Add(p.Attributes["href"].Value);
                            }
                        }
                        var desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' tab-description ')]/p");
                        var desc = "";
                        if (desc1 != null)
                            desc = desc1[0].InnerText.Replace("&nbsp;", "").Trim();
                        else
                        {
                            desc1 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' tab-description ')]/div");
                            if (desc1 != null)
                            {
                                foreach (var d in desc1)
                                {
                                    desc += d.InnerText + "\r\n";
                                }
                                desc = desc.Trim();
                            }
                        }
                        var siz = "";
                        var size = doc2.DocumentNode.SelectNodes("//label[contains(concat(' ', @for, ' '), ' option-value')]");
                        var col = "";
                        if (size != null)
                        {
                            foreach (var c in size)
                            {
                                siz += c.InnerText.Trim() + "; ";
                            }
                            siz = siz.Substring(0, siz.Length - 2);
                        }

                        var cat = "";
                        var cats = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' breadcrumb ')]/a");
                        foreach (var c in cats)
                        {
                            var temp = c.InnerText;
                            if (!temp.Contains("Главная") && !temp.Contains(title))
                            {
                                cat += temp.Replace("Главная", "").Trim() + "/";
                            }
                        }
                        if (cat.Length > 0)
                            cat = cat.Substring(0, cat.Length - 1);
                        products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, CategoryPath = cat, Size = siz, Photos = phs });
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            Helpers.SaveToFile(products, path.Text + @"\Trikbel.xlsx");
            StatusStrip("Trikbel");
        }

        private void GetTrimedwedya(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = "";//GetCookiePost("http://www.trimedwedya.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.trimedwedya.ru", "//h2/a",//"//div[contains(concat(' ', @class, ' '), ' spacer ')]/div/a",
                    "//div[contains(concat(' ', @class, ' '), ' page_navi ')]/strong/a", null);

                foreach (var res in prod)
                {
                    try
                    {
                        var client = new System.Net.WebClient();
                        var data = client.OpenRead(res);
                        var reader = new StreamReader(data);
                        string s = reader.ReadToEnd();
                        data.Close();
                        reader.Close();
                        var doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                        var title =
                            doc2.DocumentNode.SelectNodes(
                                "//div[contains(concat(' ', @class, ' '), ' productdetails-view productdetails ')]/h1")[
                                    0].InnerText.Trim();
                        var artic = "";
                        if (title.IndexOf(" ") > 0)
                            artic = title.Substring(0, title.IndexOf(" "));
                        if (artic.Trim() == string.Empty)
                            artic = title;
                        var price =
                            doc2.DocumentNode.SelectNodes(
                                "//span[contains(concat(' ', @class, ' '), ' PricesalesPrice ')]")[0].InnerText.Replace(
                                    "руб", "").Trim();
                        var photos =
                            doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' main-image ')]/a");
                        var phs = new List<string>();
                        if (photos != null)
                        {
                            foreach (var p in photos)
                            {
                                phs.Add(p.Attributes["href"].Value);
                            }
                        }
                        var desc1 =
                            doc2.DocumentNode.SelectNodes(
                                "//div[contains(concat(' ', @class, ' '), ' product-short-description ')]");
                        var desc = "";
                        if (desc1 != null)
                            desc = desc1[0].InnerText.Trim();
                        var siz = "";
                        var color =
                            doc2.DocumentNode.SelectNodes(
                                "//label[contains(concat(' ', @class, ' '), ' other-customfield ')]");
                        var col = "";
                        if (color != null)
                        {
                            foreach (var c in color)
                            {
                                if (c.Attributes["for"].Value == "10")
                                    siz += c.InnerText.Trim() + "; ";
                                else if (c.Attributes["for"].Value == "5")
                                    col += c.InnerText.Trim() + "; ";
                            }
                            siz = siz.Substring(0, siz.Length - 2);
                            col = col.Substring(0, col.Length - 2);
                        }

                        var win = Encoding.GetEncoding("windows-1251");
                        byte[] winBytes = win.GetBytes(catalog.Name);
                        var cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);

                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = price,
                            CategoryPath = cat,
                            Size = siz,
                            Photos = phs
                        });
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            Helpers.SaveToFile(products, "Trimedwedya");
            StatusStrip("Trimedwedya");
        }

        private void GetTrikobakh(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            foreach (var catalog in list)
            {

                var prod = Helpers.GetProductLinks(catalog.Url, "", "http://trikobakh.com", "//a[contains(concat(' ', @class, ' '), ' prod_link ')]",
                    "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);

                if (prod.Count == 0)
                    continue;
                var cook = Helpers.GetCookiePost("http://trikobakh.com", new NameValueCollection() { { "Itemid", "1" }, { "option", "com_virtuemart" }, { "product_currency", "RUB" }, { "do_coupon", "yes" } });

                foreach (var res in prod)
                {
                    var client = new System.Net.WebClient();
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data);
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var title = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' katalog ')]/h1")[0].InnerText;
                    var artic = Regex.Replace(title, @"[^\d]", "");
                    var price = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' prod_price_new ')]")[0].InnerText.Replace("руб.", "").Replace(" ", "").Trim();
                    var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' MagicZoomPlus ')]");
                    var phs = new List<string>();
                    if (photos != null)
                    {
                        foreach (var p in photos)
                        {
                            //if (!p.Attributes["href"].Value.Contains("___"))
                            phs.Add("http://trikobakh.com/" + p.Attributes["href"].Value);
                        }
                    }
                    var desc1 = doc2.DocumentNode.SelectNodes("//td[contains(concat(' ', @class, ' '), ' product_fly_right ')]/p");
                    var desc = "";
                    var siz = "";
                    foreach (var d in desc1)
                    {
                        var temp = d.InnerText.Trim();
                        var ddd = d.SelectNodes("/a");
                        if (ddd == null)
                        {
                            if (!temp.Contains("Как ухаживать за трикотажными изделиями?") &&
                                !temp.Contains("Как определить размер?"))
                            {
                                desc += temp + ". ";

                                if (temp.Contains("Размер"))
                                {
                                    siz =
                                        temp.Substring(temp.IndexOf(":") + 1, temp.Length - temp.IndexOf(":") - 1)
                                            .Replace(";Выбрать", "")
                                            .Trim();
                                }
                            }
                        }
                        if(desc.Length>1)
                            desc = desc.Replace("..", ".");
                    }
                    var color = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' vmCartDetails ')]/select");
                    var col = "";

                    if (color != null)
                        col = color[0].InnerText.Replace("Выбрать", "").Trim().Replace("\n", "; ").Trim();
                    
                    var win = Encoding.GetEncoding("windows-1251");
                    byte[] winBytes = win.GetBytes(catalog.Name);
                    var cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);

                    products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, CategoryPath = cat, Size = siz, Photos = phs });
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Trikobakh.xlsx");
            StatusStrip("Trikobakh");
        }

        private void GetLeggi(IEnumerable<string> list)
        {

            var cook = "";
            var products = new List<Product>();
            using (var browser = new IE("http://leggi.com.ua/index.php?route=account/login"))
            {
                //browser.SetCookie("http://leggi.com.ua/", "currency=RUB;");
                var ex = browser.Link(Find.ByUrl("http://leggi.com.ua/index.php?route=account/logout"));
                if (!ex.Exists)
                {
                    browser.Link(Find.ByText("Рубль")).Click();

                    browser.TextField(Find.ByName("email")).TypeText("stakan__85@mail.ru");

                    browser.TextField(Find.ByName("password")).TypeText("arishka");
                    var lFields = browser.Links;
                    int i = 1;
                    foreach (var field in lFields)
                    {
                        if (field.ClassName == "button")
                        {
                            if (i == 2)
                                field.Click();
                            i++;
                        }
                    }
                }
                cook = browser.Eval("document.cookie");//.Replace("UAH","RUB");

                //browser.AutoClose = true;
            }

            foreach (var catalog in list)
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(catalog);
                var prLink = new List<string>();
                var a = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn ')]/table/tr/td/a[1]");
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                    if (prLink[0].Contains("category"))
                    {
                        var prLink2 = new List<string>();
                        foreach (var pr in prLink)
                        {
                            var web2 = new HtmlWeb();
                            HtmlAgilityPack.HtmlDocument doc2 = web2.Load(pr);
                            var a2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn ')]/table/tr/td/a[1]");
                            if (a2 != null)
                            {
                                foreach (var p in a2)
                                {
                                    prLink2.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                }
                            }
                        }
                        prLink = prLink2;
                        if (prLink[0].Contains("category"))
                        {
                            var prLink3 = new List<string>();
                            foreach (var pr in prLink)
                            {
                                var web2 = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load(pr);
                                var a2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn ')]/table/tr/td/a[1]");
                                if (a2 != null)
                                {
                                    foreach (var p in a2)
                                    {
                                        prLink3.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                    }
                                }
                            }
                            prLink = prLink3;
                        }
                    }
                    var pages = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' links ')]/a");
                    if (pages != null && pages.Count > 1)
                    {
                        var preLink = "";
                        foreach (var pag in pages)
                        {
                            var link2 = WebUtility.HtmlDecode(pag.Attributes["href"].Value);
                            if (link2 != preLink)
                            {
                                var web2 = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load(link2);
                                var a2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn ')]/table/tr/td/a[1]");

                                foreach (var p in a2)
                                {
                                    prLink.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                }
                                preLink = link2;
                            }
                        }
                    }
                }
                var prod = new HashSet<string>(prLink);

                foreach (var res in prod)
                {
                    var client = new System.Net.WebClient();
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                    client.Headers.Add(HttpRequestHeader.Accept, "text/html, application/xhtml+xml, */*");
                    client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko");
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data, Encoding.GetEncoding("windows-1251"));
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var title = doc2.DocumentNode.SelectNodes("//h1[contains(concat(' ', @class, ' '), ' titleName ')]")[0].InnerText;
                    //var desc = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' tab_description ')]")[0].InnerText.Replace("&nbsp;", "").Trim();
                    //if (desc.Length > 0)
                    //{
                    //    if (desc.Contains("<!--"))
                    //    {
                    //        desc = desc.Substring(desc.IndexOf("<!--"), desc.LastIndexOf("-->") + 3 - desc.IndexOf("<!--")).Trim();
                    //    }
                    //}
                    var desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' tab_description ')]/p", "", new List<string>() { "if", "регистрации" });
                    if (desc.Length == 0)
                    {
                        desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' tab_description ')]/div/p", "", new List<string>() { "if", "регистрации" });
                    }
                    var price = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn')]/div/table/tr/td/table/tr[1]/td[2]")[0].InnerText.Replace("р.", "").Replace(",", "").Replace(".", ",").Trim();
                    var artic = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn')]/div/table/tr/td/table/tr[3]/td[2]")[0].InnerText.Trim();
                    string size = "", col = "";
                    var countSelect = doc2.DocumentNode.SelectNodes("//table/tr/td/select[1]");
                    if (countSelect != null)
                    {
                        foreach (var c in countSelect)
                        {
                            var cc = c.ParentNode.InnerHtml;
                            var n = cc.Substring(0, cc.IndexOf("<")).Trim();
                            if (n.Contains("Размер"))
                            {
                                size = c.InnerText.Replace("  ", "").Replace("\r\n", "; ").Trim();
                                size = size.Substring(1, size.Length - 2);
                            }
                            else if (n.Contains("Цвет"))
                            {
                                col = c.InnerText.Replace("  ", "").Replace("\r\n", "; ").Trim();
                                col = col.Substring(1, col.Length - 2);
                            }
                        }
                    }
                    var cat = "";
                    var categ = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' categoriesContent')]/ul/li/a/b");

                    var categ2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' categoriesContent')]/ul/li/ul/li/a/b");
                    if (categ != null)
                        cat = categ[0].InnerText.Trim();
                    else if (categ2 != null)
                    {
                        var tr = categ2[0].ParentNode.ParentNode.ParentNode.ParentNode.FirstChild.InnerText;
                        cat += tr + "/" + categ2[0].InnerText.Trim();
                    }
                    else
                    {
                        var categ3 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' categoriesContent')]/ul/li/ul/li/ul/li/a/b");

                        var c1 = categ3[0].ParentNode.ParentNode.ParentNode.ParentNode.FirstChild.InnerText;
                        var c2 = categ3[0].ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.FirstChild.InnerText;
                        cat += c2 + "/" + c1 + "/" + categ3[0].InnerText.Trim();
                    }
                    cat = cat.Replace("&quot;", "");
                    var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' thickbox ')]");
                    var ph = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' centerColumn ')]/div/table/tr/td/a")[0].Attributes["href"].Value;

                    var phs = new List<string>() { ph };
                    if (photos != null)
                    {
                        foreach (var p in photos)
                        {
                            phs.Add(p.Attributes["href"].Value);
                        }
                        var temp = new HashSet<string>(phs);
                        phs = temp.ToList();
                    }
                    products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, Photos = phs, CategoryPath = cat, Size = size });
                }
            }
            Helpers.SaveToFile(products, path.Text + @"\Leggi.xlsx");
            StatusStrip("Leggi");
        }

        private void GetOzkan(IEnumerable<Category> list)
        {
            var products = new List<Ozkan>();
            var cook = Helpers.GetCookiePost("http://www.ozkanwear.com/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.ozkanwear.com/catalog/", "//div[contains(concat(' ', @class, ' '), ' model ')]/a",
                    "//div[contains(concat(' ', @id, ' '), ' pages_down ')]/p/a", null);
                foreach (var res in prod)
                {
                    var client = new System.Net.WebClient();
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data, Encoding.GetEncoding("windows-1251"));
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var title = doc2.DocumentNode.SelectNodes("//p[contains(concat(' ', @class, ' '), ' name_model ')]")[0].InnerText;
                    var artic = doc2.DocumentNode.SelectNodes("//p[contains(concat(' ', @class, ' '), ' class_code_model ')]/span")[0].InnerText;
                    var photos = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @class, ' '), ' small_pic ')]");
                    var phs = new List<string>();
                    if (photos != null)
                    {
                        foreach (var p in photos)
                        {
                            phs.Add("http://www.ozkanwear.com/catalog/" + p.Attributes["src"].Value);
                        }
                    }

                    var desc = doc2.DocumentNode.SelectNodes("//p[contains(concat(' ', @class, ' '), ' text_model ')]")[0].InnerText.Trim();
                    if (desc.Contains("Дополнительная информация отсутствует"))
                        desc = "";
                    var cat = "";
                    if (catalog.Url.Contains("id_type=2"))
                        cat = "Женщинам/";
                    else if (catalog.Url.Contains("id_type=1"))
                        cat = "Мужчинам/";
                    else if (catalog.Url.Contains("id_type=3"))
                        cat = "Мальчикам/";
                    else if (catalog.Url.Contains("id_type=4"))
                        cat = "Девочкам/";
                    cat += catalog.Name;



                    var price1 = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[8]");
                    var priceUpak = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[9]");
                    var color = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[4]");
                    var size = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[3]");
                    var name = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[2]");
                    var articSku = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[1]");
                    var nalich = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[7]");//no-x
                    var kol = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' price_tb')]/tr/td[6]");
                    if (price1 != null)
                    {
                        for (int i = 0; i < price1.Count; i++)
                        {
                            var col = color[i + 1].InnerText.Trim();
                            var siz = name[i + 1].InnerText.Trim() + " " + size[i + 1].InnerText.Trim();
                            var arS = articSku[i + 1].InnerText.Trim();
                            var nal = nalich[i].InnerText.Contains("X") ? "no_sale" : "";
                            var k = kol[i].InnerText;
                            var pr2 = priceUpak[i].InnerText.Replace(".", ",");
                            var pr1 = price1[i].InnerText.Replace(".", ",");
                            products.Add(new Ozkan()
                            {
                                Url = res,
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = pr1,
                                Photos = phs,
                                CategoryPath = cat,
                                Size = siz,
                                ArticleSKU = arS,
                                Kolichestvo = k,
                                Nalichie = nal,
                                PriceUpak = pr2
                            });
                        }
                    }
                    else
                    {
                        var price = doc2.DocumentNode.SelectNodes("//p[contains(concat(' ', @class, ' '), ' price_model ')]/span")[0].InnerText.Replace("у.е.", "").Replace("Цена:", "").Replace("&nbsp;", "").Trim();
                        products.Add(new Ozkan()
                        {
                            Url = res,
                            Article = artic,
                            Color = "",
                            Description = desc,
                            Name = title,
                            Price = price,
                            Photos = phs,
                            CategoryPath = cat,
                            Size = "",
                            ArticleSKU = "",
                            Kolichestvo = "",
                            Nalichie = "",
                            PriceUpak = ""
                        });
                    }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Ozkan.xlsx");
            StatusStrip("Ozkan");
        }

        private void GetTvoi(IEnumerable<string> list)
        {
            var products = new List<Product>();
            //get product
            var cook = Helpers.GetCookiePost("http://tvoe.ru/collection/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog, cook, catalog, "//dt/a",
                    "//div[contains(concat(' ', @class, ' '), ' right pages ')]/a", null);

                foreach (var res in prod)
                {
                    var client = new System.Net.WebClient();
                    var data = client.OpenRead(res);
                    var reader = new StreamReader(data, Encoding.GetEncoding("windows-1251"));
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc2 = new HtmlAgilityPack.HtmlDocument();
                    doc2.LoadHtml(HttpUtility.HtmlDecode(s));

                    var title = doc2.DocumentNode.SelectNodes("//dt/strong")[0].InnerText;
                    var d = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' info-block additional ')]");

                    var desc = doc2.DocumentNode.SelectNodes("//dd/p")[0].InnerText.Trim();
                    var price = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @id, ' '), ' share_price ')]")[0].InnerText.Replace("р.", "");
                    var artic = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' info-block colors')]/div[contains(concat(' ', @class, ' '), ' right ')]")[0].InnerText.Replace("Артикул:", "");
                    var color = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' info-block colors')]/div/ul/li/a");
                    var col = "";
                    if (color != null)
                    {
                        foreach (var c in color)
                        {
                            col += c.InnerText + "; ";
                        }
                        col = col.Substring(0, col.Length - 2);
                    }
                    var size = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' info-block sizes')]/div/ul/li/b");
                    var siz = "";
                    if (size != null)
                    {
                        foreach (var si in size)
                        {
                            siz += si.InnerText + "; ";
                        }
                        siz = siz.Substring(0, siz.Length - 2);
                    }

                    var cat = "";
                    if (catalog.Contains("woman"))
                        cat = "Женщинам/";
                    else
                        cat = "Мужчинам/";
                    var categ = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' left')]/ul/li");
                    if (categ != null)
                    {
                        foreach (var c in categ)
                        {
                            if (!c.InnerHtml.Contains("<a"))
                            {
                                cat += c.InnerText;
                                break;
                            }
                        }
                    }

                    var photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' previews ')]/ul/li/a");
                    var phs = new List<string>();
                    if (photos == null)
                    {
                        var photo = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @id, ' '), ' main-image ')]")[0].Attributes["src"].Value;
                        phs.Add("http://tvoe.ru" + photo);
                    }
                    else
                    {
                        foreach (var p in photos)
                        {
                            phs.Add("http://tvoe.ru" + p.Attributes["href"].Value);
                        }
                    }
                    products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, Photos = phs, CategoryPath = cat, Size = siz });
                }
            }
            Helpers.SaveToFile(products, path.Text + @"\Tvoi.xlsx");
            StatusStrip("Tvoi");
        }

        private void Open_Click(object sender, EventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path.Text = fbd.SelectedPath;
            }
        }

        // Updates all child tree nodes recursively.
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    // If the current node has child nodes, call the CheckAllChildsNodes method recursively.
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }

        // NOTE   This code can be added to the BeforeCheck event handler instead of the AfterCheck event.
        // After a tree node's Checked property is changed, all its child nodes are updated to the same value.
        private void node_AfterCheck(object sender, TreeViewEventArgs e)
        {
            // The code only executes if the user caused the checked state to change.
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    /* Calls the CheckAllChildNodes method, passing in the current 
                    Checked value of the TreeNode whose checked state changed. */
                    this.CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }

        private void StatusStrip(string nameTreeView)
        {
            //TreeNode tree = treeView1.Nodes.Cast<TreeNode>().FirstOrDefault(treeNode => treeNode.Text == nameTreeView);
            //if (tree != null)
            //{
            //    tree.ForeColor = Color.RoyalBlue;
            //    tree.Checked = false;
            //}
            //var end = countStripStatus.Text.Substring(countStripStatus.Text.IndexOf("из"), countStripStatus.Text.Length - countStripStatus.Text.IndexOf("из"));
            //var strip = countStripStatus.Text.Replace(end, "");
            //var numb = Convert.ToInt32(Regex.Replace(strip, @"[^\d]", ""));
            //countStripStatus.Text = "Скачено " + (numb + 1) + " " + end;
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    treeView1.Nodes.Clear();
        //    int i = 0;
        //    button1.Enabled = false;
        //    Start.Enabled = false;
        //    button1.Text = "Идёт проверка сайтов...";
        //    foreach (var big in shopBigs)
        //    {
        //        try
        //        {
        //            var client = new WebClient();
        //            var data = client.OpenRead(big.Url);
        //            treeView1.Nodes.Add(big.Name);
        //        }
        //        catch (Exception ex)
        //        {
        //            var node = new TreeNode(big.Name) { ForeColor = Color.Red, ToolTipText = "Сервер недоступен" };
        //            treeView1.Nodes.Add(node);
        //        }
        //        foreach (var cat in big.CatalogList)
        //        {
        //            treeView1.Nodes[i].Nodes.Add(cat.Name);
        //        }
        //        i++;
        //    }

        //    foreach (var sh in shops)
        //    {
        //        try
        //        {
        //            var client = new WebClient();
        //            var data = client.OpenRead(sh.Url);
        //            treeView1.Nodes.Add(sh.Name);
        //            //checkedListBox1.Items.Add(sh.Name, CheckState.Unchecked);
        //        }
        //        catch (Exception ex)
        //        {
        //            var node = new TreeNode(sh.Name) { ForeColor = Color.Red, ToolTipText = "Сервер недоступен" };
        //            treeView1.Nodes.Add(node);
        //            //checkedListBox1.Items.Add(sh.Name, CheckState.Indeterminate);
        //        }
        //        //Unchecked); //CheckState.Checked);
        //    }
        //    button1.Enabled = true;
        //    Start.Enabled = true;
        //    button1.Text = "Проверить сайты на доступность";
        //}

    }
}



