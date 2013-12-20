using System.Collections.Specialized;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Web;
using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using WatiN.Core;
using Form = System.Windows.Forms.Form;
using System.Globalization;
using Image = System.Drawing.Image;

namespace ParserCatalog
{
    public partial class Form1 : Form
    {
        private List<Shop> shops;
        public Form1()
        {
            InitializeComponent();
            shops = new List<Shop>();
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
            foreach (var sh in shops.Select(x => x.Name))
            {
                checkedListBox1.Items.Add(sh, CheckState.Unchecked); //CheckState.Checked);
            }

            //nameCB.SelectedIndex = 1;
        }

        private void Open_Click(object sender, EventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path.Text = fbd.SelectedPath;
            }
        }

        private void Start_Click(object sender, EventArgs e)
        {
            Start.Enabled = false;
            Start.Text = "Подождите...";

            var pr = new List<int>();
            foreach (var p in checkedListBox1.CheckedIndices)
            {
                pr.Add(Convert.ToInt32(p));
            }
            var stL = new List<string>();
            var st = new Stopwatch();
            st.Start();
            Parallel.ForEach(pr, num =>
            {
                var shop = shops[num];
                var products = new List<Product>();
                var catList = new List<Category>();

                var client = new System.Net.WebClient();
                var data = client.OpenRead(shop.Url);
                var reader = new StreamReader(data, Encoding.GetEncoding("windows-1251"));
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                var page = new HtmlAgilityPack.HtmlDocument();
                page.LoadHtml(s);

                //get catalogs
                var arr = new string[] { "categ", "catal", "woman", "man", "katalog", "kategorii", "platja", "aksessuary", "roomdecor", "folder", "collect" };

                var query = "//ul/li/a";
                if (shop.Url.Contains("trimedwedya") || shop.Url.Contains("artvision-opt"))
                    query = "//ul/li/ul/li/a";
                else if (shop.Url.Contains("butterfly-dress"))
                    query = "//ul/li/ul/li/div/a";
                else if (shop.Url.Contains("s-trikbel"))
                    query = "//li[contains(concat(' ', @class, ' '), ' name ')]/a";
                else if (shop.Url.Contains("roomdecor"))
                    query = "//li/ul/li/a";
                else if (shop.Url.Contains("nashipupsi"))
                    query = "//a[contains(concat(' ', @href, ' '), 'folder')]";
                else if (shop.Url.Contains("opt-ekonom"))
                    query = "//span[contains(concat(' ', @class, ' '), ' inner ')]/a";
                else if (shop.Url.Contains("lemming"))
                    query = "//span/a";
                else if (shop.Url.Contains("piniolo"))
                    query = "//li[contains(concat(' ', @class, ' '), ' item ')]/a";
                else if (shop.Url.Contains("witerra"))
                    query = "//td[contains(concat(' ', @class, ' '), ' boxText ')]/a";
                else if (shop.Url.Contains("ru.gipnozstyle"))
                    query = "//div[contains(concat(' ', @class, ' '), ' twocol ')]/a";
                else if (shop.Url.Contains("shop-nogti"))
                    query = "//div/div/div/a";
                else if (shop.Url.Contains("iv-trikotage"))
                    query = "//div[contains(concat(' ', @class, ' '), ' menu_spec ')]/ul/li/a";
                else if (shop.Url.Contains("optovik-centr"))
                    query = "//a[contains(concat(' ', @class, ' '), ' mainlevel_frontpage_categories ')]";
                else if (shop.Url.Contains("japan-cosmetic"))
                    query = "//div[contains(concat(' ', @class, ' '), 'moduletableproizv')]/div/a";

                var cats = page.DocumentNode.SelectNodes(query);
                foreach (var cat in cats)
                {
                    if (cat.Attributes.Count > 0)
                    {
                        var link = cat.Attributes["href"].Value;
                        bool good = arr.Any(ar => link.Contains(ar));
                        if (link.Contains("s-trikbel") || shop.Url.Contains("artvision-opt") || shop.Url.Contains("opt-ekonom") || shop.Url.Contains("witerra") || shop.Url.Contains("ru.gipnozstyle") || shop.Url.Contains("trikotage") || shop.Url.Contains("npopt") || shop.Url.Contains("japan-cosmetic"))
                            good = true;
                        if (link.Contains("roomdecor") && (link.Contains("6195") || link.Contains("6159")))
                            good = false;
                        if (shop.Url.Contains("xn----0tbbbddeld.xn--p1ai") || shop.Url.Contains("td-adel"))
                        {
                            var t1 = cat.ParentNode.InnerHtml;
                            if (t1.Contains("<ul") || link.Contains("new"))
                                good = false;
                        }

                        if (good)
                        {
                            if (shop.Url.Contains("www.trimedwedya.ru"))
                                link = "http://www.trimedwedya.ru" + link;
                            else if (shop.Url.Contains("td-adel"))
                                link = "http://td-adel.ru" + link;
                            else if (shop.Url.Contains("xn----0tbbbddeld.xn--p1ai"))
                                link = "http://xn----0tbbbddeld.xn--p1ai/" + link;
                            else if (shop.Url.Contains("lemming.su"))
                                link = "http://lemming.su" + link;
                            else if (!link.Contains(shop.Url))
                                link = shop.Url + link;
                            catList.Add(new Category() { Name = cat.InnerText, Url = WebUtility.HtmlDecode(link) });
                        }
                    }
                }
                var temp = new HashSet<string>(catList.Select(x => x.Url));
                var cL = new List<Category>();
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

                if (shop.Url.Contains("tvoe"))
                    GetTvoi(cL.Select(x => x.Url));
                else if (shop.Url.Contains("ozkan"))
                    GetOzkan(cL);
                else if (shop.Url.Contains("leggi"))
                    GetLeggi(cL.Select(x => x.Url));
                else if (shop.Url.Contains("trikobakh"))
                    GetTrikobakh(cL);
                else if (shop.Url.Contains("trimedwedya"))
                    GetTrimedwedya(cL);
                else if (shop.Url.Contains("s-trikbel"))
                    GetTrikbel(cL);
                else if (shop.Url.Contains("butterfly-dress"))
                    GetButterfly(cL);
                else if (shop.Url.Contains("aventum"))
                    GetAventum(cL);
                else if (shop.Url.Contains("sportoptovik"))
                    GetSportoptovik(cL);
                else if (shop.Url.Contains("roomdecor"))
                    GetRoomdecor(cL);
                else if (shop.Url.Contains("nashipupsi"))
                    GetNashipupsi(cL);
                else if (shop.Url.Contains("xn----0tbbbddeld.xn--p1ai"))
                {
                    cL.RemoveAt(0);
                    GetSportOpt(cL);
                }
                else if (shop.Url.Contains("artvision-opt"))
                    GetArtvision(cL);
                else if (shop.Url.Contains("td-adel"))
                {
                    cL.RemoveAt(0);
                    GetAdel(cL);
                }
                else if (shop.Url.Contains("opt-ekonom"))
                    GetOptEconom(cL);
                else if (shop.Url.Contains("naksa"))
                    GetNaksa(cL);
                else if (shop.Url.Contains("nobi54"))
                    GetNobi(cL);
                else if (shop.Url.Contains("lemming"))
                    GetLemming(cL);
                else if (shop.Url.Contains("piniolo"))
                    GetPiniolo(cL);
                else if (shop.Url.Contains("witerra"))
                    GetWiterra(cL);
                else if (shop.Url.Contains("gipnozstyle"))
                {
                    cL.RemoveAt(0);
                    GetGipnozstyle(cL);
                }
                else if (shop.Url.Contains("noski-a42"))
                    GetNoski(cL);
                else if (shop.Url.Contains("trikotage"))
                {
                    if (!cL.Any())
                        cL.Add(new Category() { Url = "http://iv-trikotage.ru/" });
                    GetTrikotage(cL);
                }
                else if (shop.Url.Contains("shop-nogti"))
                    GetShopNogti(cL);
                else if (cL[0].Url.Contains("npopt"))
                    GetNpopt(cL);
                else if (cL[0].Url.Contains("optovik-centr"))
                    GetOptovikCentr(cL);
                else if (cL[0].Url.Contains("japan-cosmetic"))
                    GetJapanCosmetic(cL);
                stL.Add(st.Elapsed.ToString());
            });
            st.Stop();
            stL.Add(st.Elapsed.ToString());
            Start.Enabled = true;
            Start.Text = "Начать парсинг";
        }
        private void GetJapanCosmetic(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://japan-cosmetic.biz/", new NameValueCollection());
            int i = 0;
            var re = new List<string>();
            foreach (var catalog in list)
            {
                if (catalog.Url.IndexOf(".htm") == -1)
                    continue; //  
                var prod = GetProductLinks(catalog.Url, cook, "http://japan-cosmetic.biz", "//td/a[contains(text(), '[Подробнее...]')]", "//ul[contains(concat(' ', @class, ' '), 'pagination')]/li/a", null);
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

                    var doc2 = GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' name ')]");
                    var artic = GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' identifier ')]").Replace("Арт.:", "").Trim();
                    var price = GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' productPrice ')]").Replace("руб.", "").Trim();
                    desc = GetItemsInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' description ')]/p | //span[contains(concat(' ', @itemprop, ' '), ' description ')]/ul/li", "", null, " ");
                    var desc2 = GetItemsInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' brand ')]", "", null, "\r\n");
                    desc = (desc + "\r\n" + desc2).Trim();
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]");

                    cat = GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), ' pathway ')]", "", new List<string>() { "На главную" }, "/");

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
            SaveToFile(products, "JapanCosmetic");
        }

        private void GetOptovikCentr(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://optovik-centr.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://optovik-centr.ru", "//td/span[contains(concat(' ', @class, ' '), ' h3 ')]/a", "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);

                if (prod.Count == 0)
                {
                    var cat = GetProductLinks(catalog.Url, cook, "http://optovik-centr.ru", "//a[contains(concat(' ', @class, ' '), ' subcat ')]", null);
                    if (cat.Any())
                    {
                        var temp = new List<string>();
                        foreach (var c in cat)
                        {
                            var tr = GetProductLinks(c, cook, "http://optovik-centr.ru", "//td/span[contains(concat(' ', @class, ' '), ' h3 ')]/a", "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);
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

                    var doc2 = GetHtmlDocument(res, catalog.Url, null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' h3 ')]");
                    var artic = GetItemInnerText(doc2, "//td/b").Replace("Артикул:", "").Replace(".", "").Trim();
                    var price = GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'price_')]").Replace("руб.", "");
                    desc = GetItemsInnerHtml(doc2, "//div[contains(concat(' ', @class, ' '), ' browseProductDescription ')]", "", null, " ");
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
                    phs = GetPhoto(doc2, "//td/a[contains(concat(' ', @onclick, ' '), 'optovik-centr')]");

                    cat = GetItemsInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/a", "", new List<string>() { "Главная", "Каталог" }, "/");

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
            SaveToFile(products, "OptovikCentr");
        }

        private void GetNpopt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://npopt.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://npopt.ru", "//div/h3/a", "//div[contains(concat(' ', @class, ' '), ' bottom ')]/ul/li/a", null, catalog.Url);

                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' center_side ')]/h2");
                    var artic = title.Substring(title.Length - 6);
                    var price = GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'price_')]").Replace("руб.", "").Trim();
                    desc = GetItemsInnerHtml(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", null, " ");
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
                        desc = GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", null, "\r\n");
                    }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), ' img-popup ')]", "//div[contains(concat(' ', @class, ' '), 'description')]/p/img", "http://npopt.ru", "http://npopt.ru");
                    if (desc.Length == 0)
                        phs.AddRange(GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), 'description')]/p/p/img", "", "http://npopt.ru", "", "", "src"));
                    cat = GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/ul/li/a", "", new List<string>() { "Главная", "Каталог товаров" }, "/");

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
            SaveToFile(products, "Npopt");
        }

        private void GetNoski(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://noski-a42.ru/", new NameValueCollection());
            var folder = path.Text + @"\" + "Noski";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url + "?page=all", cook, "http://noski-a42.ru/", "//div[contains(concat(' ', @class, ' '), ' product_info ')]/h3/a", null);
                if (prod.Count == 0)
                    continue;

                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = GetHtmlDocument(res, catalog.Url + "?page=all", null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/h1");
                    var artic = "";
                    var price = GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'price_')]");
                    desc = GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", new List<string>() { "Внимание" }, "\r\n");

                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    var dphs = GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), ' image ')]/a");
                    if (dphs.Any())
                        phs.Add(dphs[0]);
                    dphs = GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), ' images ')]/a");
                    if (dphs.Any())
                        phs.AddRange(dphs);
                    cat = GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' path ')]/a", "", new List<string>() { "Главная" }, "/");
                    size = GetItemsInnerText(doc2, "//label[contains(concat(' ', @class, ' '), ' variant_name ')]", "", null, "; ");
                    var data = GetItemsAttributt(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/h1", "", "data-product", null, ";");
                    var sizePrice = GetItemsAttributt(doc2, "//input[contains(concat(' ', @name, ' '), ' variant ')]", "chg_price(" + data + ",'", "onclick", null, "; ").Replace("')", "").Trim();
                    
                    var p=GetPhoto(doc2, "//img[contains(concat(' ', @src, ' '), '300x300')]","","","","","src");
                    var photo = SavePhoto(p, folder);
                    
                    if (sizePrice.IndexOf(";") > 0)
                    {
                        var arrPrice = Regex.Split(sizePrice, "; ");
                        var arrSize = Regex.Split(size, "; ");

                        for (int i = 0; i < arrPrice.Length; i++)
                        {
                            products.Add(new Product()
                            {
                                Url = res,
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
                            Photo=photo
                        });
                    }
                    //}
                    //catch (Exception ex) { }
                }

            }

            SaveToFile(products, "Noski");
        }

        private void GetShopNogti(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://shop-nogti.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                if (prod.Count == 0)
                {
                    var doc1 = GetHtmlDocument(catalog.Url, "http://shop-nogti.ru", null, cook);
                    var cat1 = GetPhoto(doc1, "//tr/td/a", "", "http://shop-nogti.ru", "", "categ");
                    var temp = new List<string>();
                    if (cat1.Any())
                    {
                        foreach (var c1 in cat1)
                        {
                            var tr = GetProductLinks(c1, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                            if (tr.Any())
                                temp.AddRange(tr.ToList());
                            else
                            {
                                var doc2 = GetHtmlDocument(c1, catalog.Url, null, cook);
                                var cat2 = GetPhoto(doc2, "//tr/td/a", "", "http://shop-nogti.ru", "", "categ");
                                if (cat2.Any())
                                {
                                    foreach (var c2 in cat2)
                                    {
                                        var tr2 = GetProductLinks(c2, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
                                        if (tr2.Any())
                                            temp.AddRange(tr2.ToList());
                                        else
                                        {
                                            var doc3 = GetHtmlDocument(c2, c1, null, cook);
                                            var cat3 = GetPhoto(doc3, "//tr/td/a", "", "http://shop-nogti.ru", "", "categ");
                                            if (cat3.Any())
                                            {
                                                foreach (var c3 in cat3)
                                                {
                                                    var tr3 = GetProductLinks(c3, cook, "http://shop-nogti.ru", "//div[contains(concat(' ', @class, ' '), ' browseProductContainer ')]/h2/a", null);
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

                    var doc2 = GetHtmlDocument(res, catalog.Url, null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//td/h1");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var price = GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' productPrice ')]").Replace("руб.", "").Trim();
                    desc = GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", new List<string>() { "Внимание" }, "\r\n");
                    var tre = GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/ul/li", "", null, "\r\n");

                    phs = GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]");
                    var cd = Regex.Split(res.Replace("http://shop-nogti.ru/shop/details/", ""), "/");
                    if (cd != null)
                    {
                        foreach (var c in cd)
                        {
                            var it = 0;
                            var number = Int32.TryParse(c, out it);
                            if (!number && !c.Contains(".html") && !c.Contains("'"))
                            {
                                var temp = GetItemInnerText(doc2, "//a[contains(concat(' ', @href, ' '), '" + c + "')]");
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
            SaveToFile(products, "ShopNogti");
        }
        private void GetTrikotage(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://iv-trikotage.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://iv-trikotage.ru", "//a[contains(concat(' ', @class, ' '), ' fast_pro_title ')]", null);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = GetHtmlDocument(res, "http://iv-trikotage.ru", null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//h1[contains(concat(' ', @id, ' '), ' pro_fast_name ')]");
                    var artic = GetItemsInnerText(doc2, "//tr/td/div", "Артикул:", null);
                    var price = GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), ' price ')]");
                    var table = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' detali ')]");
                    desc = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' desc ')]").Replace("Описание:", "").Trim();
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
                        phs.AddRange(photos.Select(p => "http://iv-trikotage.ru" + p.Attributes["src"].Value));
                    else
                    {
                        photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' zoomPup ')]/img");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => p.Attributes["src"].Value));
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
                                size = fd.InnerText.Trim() + "; ";
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
            SaveToFile(products, "Trikotage");
        }
        private void GetGipnozstyle(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://ru.gipnozstyle.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://ru.gipnozstyle.ru", "//div[contains(concat(' ', @class, ' '), ' sm ')]/a", null);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' naimenovanie ')]");
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
                    if (desc.Length == 0) { desc = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' opisanie ')]"); }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' sixcol ')]/div/a/img");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => p.Attributes["src"].Value));
                    //photos = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' tencol ')]/a");
                    //if (photos != null)
                    //    phs.AddRange(photos.Select(p => "http://ru.gipnozstyle.ru" + p.Attributes["href"].Value));
                    cat = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' font ')]");

                    var cols = doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @id, ' '), ' colors ')]/option");
                    var siz = doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @id, ' '), ' razmer ')]/option");
                    if (cols != null)
                    {
                        foreach (var fd in cols)
                        {
                            if (!string.IsNullOrEmpty(fd.Attributes["value"].Value.Trim()))
                                col = fd.Attributes["value"].Value.Trim() + "; ";
                        }
                        col = col.Substring(0, col.Length - 2);
                    }
                    if (siz != null)
                    {
                        foreach (var fd in siz)
                        {
                            if (!string.IsNullOrEmpty(fd.Attributes["value"].Value.Trim()))
                                size = fd.Attributes["value"].Value.Trim() + "; ";
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
            SaveToFile(products, "Gipnozstyle");
        }

        private void GetWiterra(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://witerra.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url + "&sort=products_sort_order&page=all", cook, "", "//td[contains(concat(' ', @class, ' '), ' productListing-data ')]/a", null);
                if (prod.Count == 0)
                {
                    var catal = GetProductLinks(catalog.Url, cook, "", "//td[contains(concat(' ', @class, ' '), ' smallText ')]/a", null);
                    var temp = new List<string>();
                    foreach (var c in catal)
                    {
                        temp.AddRange(GetProductLinks(c + "&sort=products_sort_order&page=all", cook, "", "//td[contains(concat(' ', @class, ' '), ' productListing-data ')]/a", null).ToList());
                    }
                    prod = new HashSet<string>(temp);
                }
                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    try
                    {
                        var doc2 = GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                        var col = "";
                        var size = "";
                        var desc = "";
                        var cat = "";
                        var phs = new List<string>();
                        var title = GetItemInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' pageHeading ')]");
                        var artic = "";
                        if (!string.IsNullOrEmpty(title))
                        {
                            artic = title.Substring(title.ToLower().IndexOf("арт") + 3, title.Length - (title.ToLower().IndexOf("арт") + 3)).Replace(".", "").Trim();
                        }
                        if (string.IsNullOrEmpty(artic))
                            artic = title;
                        var price = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')]", 1).Replace("руб.", "").Replace("Базовая цена:", "").Trim();
                        if (string.IsNullOrEmpty(price))
                            price = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')]").Replace("руб.", "").Replace(" ", "").Replace(".", ",").Trim();
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
                            desc = GetItemInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' main ')]/table", 1);
                        }
                        var desc2 = GetItemsInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' main ')]/div[contains(concat(' ', @style, ' '), 'justify')]", "", null);
                        if (desc2.Length > 0)
                            desc = (desc + " " + desc2).Trim();

                        var photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @class, ' '), ' zoom ')]");
                        if (photos != null)
                            phs.AddRange(photos.Select(p => p.Attributes["href"].Value));

                        //cat = HttpUtility.HtmlDecode(catalog.Name);
                        cat = GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), ' headerNavigation ')]", "", new List<string>() { "Главная" }, "/");
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
            SaveToFile(products, "Witerra");
        }

        private void GetPiniolo(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://www.piniolo.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url + "?page=0", cook, "http://www.piniolo.ru/", "//a[contains(concat(' ', @class, ' '), ' link-pv-name ')]", null);

                if (prod.Count == 0)
                    continue;
                foreach (var res in prod)
                {
                    //try
                    //{

                    var doc2 = GetHtmlDocument(res, catalog.Url, null, cook);
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' product-name ')]");
                    var artic = GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), ' skuValue ')]");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    var price = GetItemInnerText(doc2, "//strong/span/span").Replace("р.", "").Replace("р", "").Replace(" ", "").Replace(".", ",").Trim();
                    if (string.IsNullOrEmpty(price))
                        price = GetItemInnerText(doc2, "//p/span/span/span").Replace("р.", "").Replace("р", "").Replace(" ", "").Replace(".", ",").Trim();
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

                    cat = GetEncodingCategory(catalog.Name);

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
            SaveToFile(products, "Piniolo");
        }

        private void GetLemming(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://lemming.su/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://lemming.su", "//p/a", null);
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
                                if (!d.InnerText.Contains("Расцветки:") && !d.InnerText.Contains("уточняйте по телефону") && !string.IsNullOrEmpty(d.InnerText.Trim()) && !d.InnerText.Contains("Цвет №") && !d.InnerText.Contains("бланк заказа"))
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
            SaveToFile(products, "Lemming");
        }

        private void GetNobi(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://www.nobi54.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://www.nobi54.ru", "//div[contains(concat(' ', @class, ' '), ' prod-info ')]/h2/a",
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
            SaveToFile(products, "Nobi");
        }

        private void GetNaksa(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://naksa.ru/", new NameValueCollection());
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
            SaveToFile(products, "Haksa");
        }

        private void GetOptEconom(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://www.opt-ekonom.ru/", new NameValueCollection());
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
                    //try
                    //{
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

                    }
                    var photos = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @class, ' '), ' jshop_img_thumb ')]");
                    if (photos != null)
                        phs.AddRange(photos.Select(p => p.Attributes["src"].Value.Replace("thumb", "full")));
                    else
                    {
                        photos = doc2.DocumentNode.SelectNodes("//a[contains(concat(' ', @id, ' '), ' lightbox ')]");
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
                    //}
                    //catch (Exception ex) { }
                }

            }
            SaveToFile(products, "OptEconom");
        }

        private void GetAdel(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://td-adel.ru/", new NameValueCollection());
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
                        var data = client.OpenRead("http://td-adel.ru/catalog/tolstovki/tolstovka-uniseks-na-molnii-futer-petelchatyy1/");
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
                            price = GetItemInnerText(doc2,
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
            SaveToFile(products, "Adel");
        }

        private void GetArtvision(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://artvision-opt.ru/", new NameValueCollection());
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
            SaveToFile(products, "Artvision");
        }

        private void GetSportOpt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://xn----0tbbbddeld.xn--p1ai/", new NameValueCollection());
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
            SaveToFile(products, "SportOpt");
        }

        private void GetNashipupsi(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://nashipupsi.ru", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://nashipupsi.ru", "//p[contains(concat(' ', @class, ' '), ' product-name ')]/a",
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
                        var data = client.OpenRead("http://nashipupsi.ru/shop/product/450130-samokat-trehkolesnyy-vinni");
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
            SaveToFile(products, "Nashipupsi");
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
                            temp = doc3.DocumentNode.SelectNodes("//p/span");
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
                                else if (!s1.InnerText.Contains("if") && !string.IsNullOrEmpty(s1.InnerText.Trim()))
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

                        var cat = doc2.DocumentNode.SelectNodes("//ul/li/a/strong")[0].InnerText;
                        if (!string.IsNullOrEmpty(cat1))
                            cat = cat1 + "/" + cat;

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
            SaveToFile(products, "Sportoptovik");
        }

        private void GetRoomdecor(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost("http://roomdecor.su", new NameValueCollection());
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
                var prod = GetProductLinks(catalog.Url, cook, "", "//h2[contains(concat(' ', @class, ' '), ' prodtitle ')]/a", "//div[contains(concat(' ', @class, ' '), ' wpsc_page_numbers_bottom ')]/a", "&paged=", null, numberLastLink);
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
            SaveToFile(products, "Roomdecor");
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
                var cook = GetCookiePost(prLink[0], new NameValueCollection() { { "current_currency", "3" } });
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
                                phs.Add("http://aventum.cz" + p.Attributes["src"].Value);
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

                        var cat = GetEncodingCategory(catalog.Name);

                        products.Add(new Product() { Url = res, Article = artic, Color = col, Description = desc, Name = title, Price = price, CategoryPath = cat, Size = size, Photos = phs });
                    }
                    catch (Exception ex) { }
                }

            }
            SaveToFile(products, "Aventum");
        }

        private void GetButterfly(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = GetCookiePost(list.ToList()[0].Url, new NameValueCollection() { { "submit", "%D0%98%D0%B7%D0%BC%D0%B5%D0%BD%D0%B8%D1%82%D1%8C%20%D0%B2%D0%B0%D0%BB%D1%8E%D1%82%D1%83" }, { "virtuemart_currency_id", "131" } });
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url + "/results,1-150", cook, "http://butterfly-dress.com", "//div[contains(concat(' ', @class, ' '), ' bfvmcatimage ')]/a",
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
                        var desc = GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-description ')]").Replace("Описание", "").Trim();
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
            SaveToFile(products, "Butterfly");
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
                        var price2 = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' price-new ')]");
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
            SaveToFile(products, "Trikbel");
        }

        private void GetTrimedwedya(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = "";//GetCookiePost("http://www.trimedwedya.ru/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://www.trimedwedya.ru", "//h2/a",//"//div[contains(concat(' ', @class, ' '), ' spacer ')]/div/a",
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
            SaveToFile(products, "Trimedwedya");
        }

        private void GetTrikobakh(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            foreach (var catalog in list)
            {

                var prod = GetProductLinks(catalog.Url, "", "http://trikobakh.com", "//a[contains(concat(' ', @class, ' '), ' prod_link ')]",
                    "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);

                if (prod.Count == 0)
                    continue;
                var cook = GetCookiePost("http://trikobakh.com", new NameValueCollection() { { "Itemid", "1" }, { "option", "com_virtuemart" }, { "product_currency", "RUB" }, { "do_coupon", "yes" } });

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
            SaveToFile(products, "Trikobakh");
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
                    var desc = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @id, ' '), ' tab_description ')]")[0].InnerText.Replace("&nbsp;", "").Trim();
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
                                size = c.InnerText.Replace("  ", "").Replace("\r\n", ";").Trim();
                                size = size.Substring(1, size.Length - 2);
                            }
                            else if (n.Contains("Цвет"))
                            {
                                col = c.InnerText.Replace("  ", "").Replace("\r\n", ";").Trim();
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
            SaveToFile(products, "Leggi");
        }

        private void GetOzkan(IEnumerable<Category> list)
        {
            var products = new List<Ozkan>();
            var cook = GetCookiePost("http://www.ozkanwear.com/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog.Url, cook, "http://www.ozkanwear.com/catalog/", "//div[contains(concat(' ', @class, ' '), ' model ')]/p/a",
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
            SaveToFile(products, "Ozkan");
        }

        private void GetTvoi(IEnumerable<string> list)
        {
            var products = new List<Product>();
            //get product
            var cook = GetCookiePost("http://tvoe.ru/collection/", new NameValueCollection());
            foreach (var catalog in list)
            {
                var prod = GetProductLinks(catalog, cook, catalog, "//dt/a",
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
            SaveToFile(products, "Tvoi");

        }

        private string SavePhoto(List<string> photos, string path)
        {
            if (photos!=null&&photos.Any())
            {
                var web=new WebClient();
                var img = web.DownloadData(new Uri(photos[0]));
                var ms = new MemoryStream(img);
                Image image = Image.FromStream(ms);
                var name = Guid.NewGuid().ToString();
                var ttt =new ImageFormatConverter().ConvertToString(image.RawFormat).ToLower();
                path = path + @"\" + name + "." + ttt;
                image.Save(path);
                return path;
            }
            return string.Empty;
        }

        private HtmlAgilityPack.HtmlDocument GetHtmlDocument(string url, string refererLink, Encoding encode, string cook = "")
        {
            try
            {
                var client = new System.Net.WebClient();
                client.Headers.Add(HttpRequestHeader.Cookie, cook);
                client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client.Headers.Add(HttpRequestHeader.Referer, refererLink);
                client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data = client.OpenRead(url);
                if (encode == null)
                    encode = Encoding.UTF8;
                var reader = new StreamReader(data, encode);
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(HttpUtility.HtmlDecode(s));
                return doc;
            }
            catch (Exception ex) { }
            return null;
        }

        private List<string> GetPhoto(HtmlAgilityPack.HtmlDocument doc, string xPath1, string xPath2 = "", string host1 = "", string host2 = "", string word = "", string att1 = "href", string att2 = "src")
        {
            var phs = new List<string>();
            var photos = doc.DocumentNode.SelectNodes(xPath1);
            if (photos != null)
            {
                if (!string.IsNullOrEmpty(word))
                {
                    foreach (var p in photos)
                    {
                        if (p.Attributes[att1].Value.Contains(word))
                            phs.Add(host1 + p.Attributes[att1].Value);
                    }
                }
                else
                    phs.AddRange(photos.Select(p => host1 + p.Attributes[att1].Value));
            }
            if (!string.IsNullOrEmpty(xPath2))
            {
                photos = doc.DocumentNode.SelectNodes(xPath2);
                if (photos != null)
                {
                    if (!string.IsNullOrEmpty(word))
                    {
                        foreach (var p in photos)
                        {
                            if (p.Attributes[att2].Value.Contains(word))
                                phs.Add(host2 + p.Attributes[att2].Value);
                        }
                    }
                    else
                        phs.AddRange(photos.Select(p => host2 + p.Attributes[att2].Value));
                }
            }
            return new HashSet<string>(phs).ToList();
        }

        private string GetItemsAttributt(HtmlAgilityPack.HtmlDocument doc, string xPath, string word, string attribut, List<string> notWord, string split = "\r\n")
        {
            var temp = "";
            var obj = doc.DocumentNode.SelectNodes(xPath);
            if (obj != null)
            {
                foreach (var fd in obj)
                {
                    if (notWord != null && notWord.Count > 0)
                    {
                        var bo = notWord.Where(x => fd.Attributes[attribut].Value.ToLower().Contains(x.ToLower())).ToList();
                        if (!bo.Any() && !string.IsNullOrEmpty(fd.Attributes[attribut].Value.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                                temp += fd.Attributes[attribut].Value.Trim() + split;
                            else
                            {
                                if (fd.Attributes[attribut].Value.Contains(word))
                                    temp += fd.Attributes[attribut].Value.Trim().Replace(word, "") + split;
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(fd.Attributes[attribut].Value.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                                temp += fd.Attributes[attribut].Value.Trim() + split;
                            else
                            {
                                if (fd.Attributes[attribut].Value.Contains(word))
                                    temp += fd.Attributes[attribut].Value.Trim().Replace(word, "") + split;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(temp))
                    temp = temp.Substring(0, temp.Length - split.Length);
            }
            return temp;
        }

        private List<string> GetItemsAttributtList(HtmlAgilityPack.HtmlDocument doc, string xPath, string word, string attribut, List<string> notWord, List<string> replace)
        {
            var temp = new List<string>();
            var obj = doc.DocumentNode.SelectNodes(xPath);
            if (obj != null)
            {
                foreach (var fd in obj)
                {
                    if (notWord != null && notWord.Count > 0)
                    {
                        var bo = notWord.Where(x => fd.Attributes[attribut].Value.ToLower().Contains(x.ToLower())).ToList();
                        if (!bo.Any() && !string.IsNullOrEmpty(fd.Attributes[attribut].Value.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                            {
                                if (replace == null || !replace.Any())
                                    temp.Add(fd.Attributes[attribut].Value.Trim());
                                else
                                {
                                    var t = fd.Attributes[attribut].Value.Trim();
                                    foreach (var rep in replace)
                                    {
                                        t.Replace(rep, "");
                                    }
                                    if (!string.IsNullOrEmpty(t.Trim()))
                                        temp.Add(t.Trim());
                                }
                            }
                            else
                            {
                                if (fd.Attributes[attribut].Value.Contains(word))
                                {
                                    if (replace == null || !replace.Any())
                                        temp.Add(fd.Attributes[attribut].Value.Trim());
                                    else
                                    {
                                        var t = fd.Attributes[attribut].Value.Trim();
                                        foreach (var rep in replace)
                                        {
                                            t.Replace(rep, "");
                                        }
                                        if (!string.IsNullOrEmpty(t.Trim()))
                                            temp.Add(t.Trim());
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(fd.Attributes[attribut].Value.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                            {
                                if (replace == null || !replace.Any())
                                    temp.Add(fd.Attributes[attribut].Value.Trim());
                                else
                                {
                                    var t = fd.Attributes[attribut].Value.Trim();
                                    foreach (var rep in replace)
                                    {
                                        t.Replace(rep, "");
                                    }
                                    if (!string.IsNullOrEmpty(t.Trim()))
                                        temp.Add(t.Trim());
                                }
                            }
                            else
                            {
                                if (fd.Attributes[attribut].Value.Contains(word))
                                {
                                    if (replace == null || !replace.Any())
                                        temp.Add(fd.Attributes[attribut].Value.Trim());
                                    else
                                    {
                                        var t = fd.Attributes[attribut].Value.Trim();
                                        foreach (var rep in replace)
                                        {
                                            t.Replace(rep, "");
                                        }
                                        if (!string.IsNullOrEmpty(t.Trim()))
                                            temp.Add(t.Trim());
                                    }
                                }
                            }
                        }
                    }

                }
            }
            return temp;
        }

        /// <summary>
        /// Получает строку с данными из массива inner html element
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="xPath"></param>
        /// <param name="word">Словов которое должно обязательно присутсвовать в строке, при сохранение будет удалено из самой строки</param>
        /// <param name="notWord">С указанными словами строки сохраняться не будут</param>
        /// <param name="split">Разделитель в строке между элементами</param>
        /// <returns></returns>
        private string GetItemsInnerText(HtmlAgilityPack.HtmlDocument doc, string xPath, string word, List<string> notWord, string split = "\r\n")
        {
            var temp = "";
            var obj = doc.DocumentNode.SelectNodes(xPath);
            if (obj != null)
            {
                foreach (var fd in obj)
                {
                    if (notWord != null && notWord.Count > 0)
                    {
                        var bo = notWord.Where(x => fd.InnerText.ToLower().Contains(x.ToLower())).ToList();
                        if (!bo.Any() && !string.IsNullOrEmpty(fd.InnerText.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                                temp += fd.InnerText.Trim() + split;
                            else
                            {
                                if (fd.InnerText.Contains(word))
                                    temp += fd.InnerText.Trim().Replace(word, "") + split;
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(fd.InnerText.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                                temp += fd.InnerText.Trim() + split;
                            else
                            {
                                if (fd.InnerText.Contains(word))
                                    temp += fd.InnerText.Trim().Replace(word, "") + split;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(temp))
                    temp = temp.Substring(0, temp.Length - split.Length);
            }
            return temp;
        }

        private List<string> GetItemsInnerTextList(HtmlAgilityPack.HtmlDocument doc, string xPath, string word, List<string> notWord, List<string> replace)
        {
            var temp = new List<string>();
            var obj = doc.DocumentNode.SelectNodes(xPath);
            if (obj != null)
            {
                foreach (var fd in obj)
                {
                    if (notWord != null && notWord.Count > 0)
                    {
                        var bo = notWord.Where(x => fd.InnerText.ToLower().Contains(x.ToLower())).ToList();
                        if (!bo.Any() && !string.IsNullOrEmpty(fd.InnerText.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                            {
                                if (replace == null || !replace.Any())
                                    temp.Add(fd.InnerText.Trim());
                                else
                                {
                                    var t = fd.InnerText.Trim();
                                    foreach (var rep in replace)
                                    {
                                        t.Replace(rep, "");
                                    }
                                    if (!string.IsNullOrEmpty(t.Trim()))
                                        temp.Add(t.Trim());
                                }
                            }
                            else
                            {
                                if (fd.InnerText.Contains(word))
                                {
                                    if (replace == null || !replace.Any())
                                        temp.Add(fd.InnerText.Replace(word, "").Trim());
                                    else
                                    {
                                        var t = fd.InnerText.Replace(word, "").Trim();
                                        foreach (var rep in replace)
                                        {
                                            t.Replace(rep, "");
                                        }
                                        if (!string.IsNullOrEmpty(t.Trim()))
                                            temp.Add(t.Trim());
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(fd.InnerText.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                            {
                                if (replace == null || !replace.Any())
                                    temp.Add(fd.InnerText.Trim());
                                else
                                {
                                    var t = fd.InnerText.Trim();
                                    foreach (var rep in replace)
                                    {
                                        t.Replace(rep, "");
                                    }
                                    if (!string.IsNullOrEmpty(t.Trim()))
                                        temp.Add(t.Trim());
                                }
                            }
                            else
                            {
                                if (fd.InnerText.Contains(word))
                                {
                                    if (replace == null || !replace.Any())
                                        temp.Add(fd.InnerText.Replace(word, "").Trim());
                                    else
                                    {
                                        var t = fd.InnerText.Replace(word, "").Trim();
                                        foreach (var rep in replace)
                                        {
                                            t.Replace(rep, "");
                                        }
                                        if (!string.IsNullOrEmpty(t.Trim()))
                                            temp.Add(t.Trim());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return temp;
        }

        /// <summary>
        /// Получить содержимое тега data[numberObject].InnerText.Trim()
        /// </summary>
        /// <param name="doc">html документ</param>
        /// <param name="xPath">Путь к нужному тегу</param>
        /// <param name="numberObject"></param>
        /// <returns>Внутренний текст тега</returns>
        private string GetItemInnerText(HtmlAgilityPack.HtmlDocument doc, string xPath, int numberObject = 0)
        {
            var data = doc.DocumentNode.SelectNodes(xPath);
            var d = "";
            if (data != null)
            {
                if (data.Count > numberObject)
                    d = data[numberObject].InnerText.Trim();
            }
            return d;
        }

        private string GetItemsInnerHtml(HtmlAgilityPack.HtmlDocument doc, string xPath, string word, List<string> notWord, string split = "\r\n")
        {
            var temp = "";
            var obj = doc.DocumentNode.SelectNodes(xPath);
            if (obj != null)
            {
                foreach (var fd in obj)
                {
                    if (notWord != null && notWord.Count > 0)
                    {
                        var bo = notWord.Where(x => fd.InnerHtml.ToLower().Contains(x.ToLower())).ToList();
                        if (!bo.Any() && !string.IsNullOrEmpty(fd.InnerHtml.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                                temp += fd.InnerHtml.Trim() + split;
                            else
                            {
                                if (fd.InnerText.Contains(word))
                                    temp += fd.InnerHtml.Trim().Replace(word, "") + split;
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(fd.InnerHtml.Trim()))
                        {
                            if (string.IsNullOrEmpty(word))
                                temp += fd.InnerHtml.Trim() + split;
                            else
                            {
                                if (fd.InnerText.Contains(word))
                                    temp += fd.InnerHtml.Trim().Replace(word, "") + split;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(temp))
                    temp = temp.Substring(0, temp.Length - split.Length);
            }
            return temp;
        }

        private string GetItemInnerHtml(HtmlAgilityPack.HtmlDocument doc, string xPath, int numberObject = 0)
        {
            var data = doc.DocumentNode.SelectNodes(xPath);
            var d = "";
            if (data != null)
            {
                if (data.Count > numberObject)
                    d = data[numberObject].InnerHtml.Trim();
            }
            return d;
        }

        private string GetEncodingCategory(string str)
        {
            var win = Encoding.GetEncoding("windows-1251");
            byte[] winBytes = win.GetBytes(str);
            var cat = Encoding.UTF8.GetString(winBytes, 0, winBytes.Length);
            return cat;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="catalogLink"></param>
        /// <param name="cook"></param>
        /// <param name="site">host</param>
        /// <param name="xA">Путь к ссылкам на товар</param>
        /// <param name="xApage">Путь к ссылкам на страницы</param>
        /// <param name="type">По умолчанию Encoding.UTF8</param>
        /// <returns></returns>
        private HashSet<string> GetProductLinks(string catalogLink, string cook, string site, string xA, string xApage, Encoding type, string linkPage = "")
        {
            var prLink = new List<string>();
            try
            {
                var client = new System.Net.WebClient();
                if (!string.IsNullOrEmpty(cook))
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                client.Headers.Add(HttpRequestHeader.Accept,
                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client.Headers.Add(HttpRequestHeader.Referer, site + "/");
                client.Headers.Add(HttpRequestHeader.UserAgent,
                    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data = client.OpenRead(catalogLink);
                if (type == null)
                    type = Encoding.UTF8;
                var reader = new StreamReader(data, type);
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s);

                var a = doc.DocumentNode.SelectNodes(xA);
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add(site + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                    var pages = doc.DocumentNode.SelectNodes(xApage);
                    if (pages != null && pages.Count > 0)
                    {
                        var preLink = new List<string>();
                        foreach (var pag in pages)
                        {
                            var link2 = WebUtility.HtmlDecode(pag.Attributes["href"].Value);
                            if (!preLink.Contains(link2))
                            {
                                var web2 = new HtmlWeb();
                                if (linkPage.Length == 0)
                                    linkPage = site;
                                HtmlAgilityPack.HtmlDocument doc2 = web2.Load(site + link2);
                                var a2 = doc2.DocumentNode.SelectNodes(xA);
                                if (a2 != null)
                                {
                                    foreach (var p in a2)
                                    {
                                        prLink.Add(site + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                    }
                                }
                                preLink.Add(link2);
                            }
                        }
                    }

                }
            }
            catch (Exception ex) { }
            return new HashSet<string>(prLink);
        }

        private HashSet<string> GetProductLinks(string catalogLink, string cook, string site, string xA, string xApage, string strPage, Encoding type, int numberMaxLink = 2)
        {
            var prLink = new List<string>();
            try
            {
                var client = new System.Net.WebClient();
                client.Headers.Add(HttpRequestHeader.Cookie, cook);
                client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client.Headers.Add(HttpRequestHeader.Referer, site + "/");
                client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data = client.OpenRead(catalogLink);
                if (type == null)
                    type = Encoding.UTF8;
                var reader = new StreamReader(data, type);
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s);

                var a = doc.DocumentNode.SelectNodes(xA);
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add(site + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                    var pages = doc.DocumentNode.SelectNodes(xApage);
                    if (pages != null && pages.Count > 1)
                    {
                        var max = 0;
                        if (pages.Count > 2)
                        {
                            var tr = Int32.TryParse(pages[numberMaxLink].InnerText.Trim(), out max);
                            if (!tr)
                            {
                                var link = pages[numberMaxLink].Attributes["href"].Value;
                                var num = link.Substring(link.IndexOf(strPage) + strPage.Length, link.Length - (link.IndexOf(strPage) + strPage.Length));
                                tr = Int32.TryParse(num, out max);
                            }
                        }
                        else
                            max = Convert.ToInt32(pages[1].InnerText.Trim());
                        for (var i = 2; i <= max; i++)
                        {
                            var web2 = new HtmlWeb();
                            HtmlAgilityPack.HtmlDocument doc2 = web2.Load(catalogLink + strPage + i);
                            var a2 = doc2.DocumentNode.SelectNodes(xA);
                            if (a2 != null)
                            {
                                foreach (var p in a2)
                                {
                                    prLink.Add(site + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex) { }
            return new HashSet<string>(prLink);
        }

        private HashSet<string> GetProductLinks(string catalogLink, string cook, string site, string xA, Encoding type)
        {
            var prLink = new List<string>();
            try
            {
                var client = new System.Net.WebClient();
                client.Headers.Add(HttpRequestHeader.Cookie, cook);
                client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                client.Headers.Add(HttpRequestHeader.Referer, site + "/");
                client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                var data = client.OpenRead(catalogLink);
                if (type == null)
                    type = Encoding.UTF8;
                var reader = new StreamReader(data, type);
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(s);

                var a = doc.DocumentNode.SelectNodes(xA);
                if (a != null)
                {
                    foreach (var p in a)
                    {
                        prLink.Add(site + WebUtility.HtmlDecode(p.Attributes["href"].Value));
                    }
                }
            }
            catch (Exception ex) { }
            return new HashSet<string>(prLink);
        }

        private string GetCookiePost(string link, NameValueCollection col)
        {
            var request = (HttpWebRequest)HttpWebRequest.Create(link);
            var resp = (HttpWebResponse)request.GetResponse();
            var cooks = resp.Headers.GetValues("Set-Cookie");
            resp.Close();
            var cook = "";
            if (cooks != null)
            {
                if (cooks[0].Contains("path"))
                {
                    foreach (var s in cooks)
                    {
                        if (s.IndexOf(";") > s.IndexOf("="))
                        {
                            var temp = s.Substring(0, s.IndexOf(";") + 1);
                            if (!cook.Contains(temp))
                                cook += temp + " ";
                        }
                    }

                }
                else
                {
                    foreach (var c in cooks)
                    {
                        if (!c.Contains("path"))
                            cook += c.Substring(0, c.IndexOf(";") + 1) + " ";
                    }
                }
            }

            var client2 = new System.Net.WebClient();
            client2.Headers.Add(HttpRequestHeader.Cookie, cook);
            client2.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
            client2.Headers.Add(HttpRequestHeader.Referer, "https://google.com");
            client2.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
            client2.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
            try
            {
                var byt = client2.UploadValues(link, "POST", col);
                var reader2 = new StreamReader(new MemoryStream(byt));
                string s21 = reader2.ReadToEnd();
            }
            catch (Exception ex) { }
            return cook;
        }

        private void SaveToFile(List<Product> pr, string name)
        {
            var hash = new HashSet<string>(pr.Select(x => x.Url));
            var cL = new List<Product>();
            if (hash.Count != pr.Count)
            {
                foreach (var t in hash)
                {
                    foreach (var g in pr)
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
                cL = pr;
            }
            SaveExcel2007<Product>(cL, path.Text + @"\" + name + ".xlsx", "Каталог", cL.Max(x => x.Photos.Count));
        }

        private void SaveToFile(List<Sport> pr, string name)
        {
            var hash = new HashSet<string>(pr.Select(x => x.Url));
            var cL = new List<Sport>();
            if (hash.Count != pr.Count)
            {
                foreach (var t in hash)
                {
                    foreach (var g in pr)
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
                cL = pr;
            }
            SaveExcel2007<Sport>(cL, path.Text + @"\" + name + ".xlsx", "Каталог", cL.Max(x => x.Photos.Count));
        }

        private void SaveToFile(List<Ozkan> pr, string name)
        {
            SaveExcel2007<Ozkan>(pr, path.Text + @"\" + name + ".xlsx", "Каталог", pr.Max(x => x.Photos.Count));
        }

        private void SaveExcel2007<T>(IEnumerable<T> list, string path, string nameBook, int countPhoto)
        {
            if (list == null || !list.Any()) return;
            Type itemType = typeof(T);
            var props = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance).OrderBy(p => p.Name);


            var dt = new DataTable(itemType.Name) { Locale = System.Threading.Thread.CurrentThread.CurrentCulture };
            if (dt.Rows.Count < 1 && dt.Columns.Count < 1)
            {
                //Создаём столбцы
                foreach (var prop in props)
                {
                    //dt.Columns.Add("Dosage", typeof(int));
                    //dt.Columns.Add(prop.Name); //, prop.PropertyType);
                    if (prop.Name == "Photos")
                    {
                        for (int i = 0; i <= countPhoto; i++)
                            dt.Columns.Add(prop.Name + i);
                    }
                    //else if (prop.Name.Equals("Photo"))
                    //{
                    //    DataColumn column = new DataColumn("Photo"); //Create the column.
                    //    column.DataType = System.Type.GetType("System.Byte[]"); //Type byte[] to store image bytes.
                    //    column.AllowDBNull = true;
                    //    dt.Columns.Add(column);
                    //}
                    else
                    {
                        dt.Columns.Add(prop.Name);
                    }

                }
            }

            foreach (var x in list)
            {
                DataRow newRow = dt.NewRow();
                //newRow["CompanyID"] = "NewCompanyID";
                foreach (var prop in props)
                {
                    var val = prop.GetValue(x);
                    if (val == null)
                        newRow[prop.Name] = "";
                    else if (val is IList)
                    {
                        //var temp = "";
                        //var type=prop.GetType().GetGenericTypeDefinition();
                        //var pr2=type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                        int i = 0;
                        foreach (var v in (List<string>)val)
                        {
                            var name = prop.Name + i;
                            newRow[name] = v;
                            i++;
                            //temp += (temp.Length > 0 ? " , " : "") + v;
                        }
                        //newRow[prop.Name] = temp;
                    }
                    else if (prop.Name.Contains("Price"))
                    {
                        double price = 0;
                        var converts = val.ToString().Replace(",", ".").Replace(" ", "");
                        var conv = Double.TryParse(val.ToString(), NumberStyles.AllowDecimalPoint, CultureInfo.CreateSpecificCulture("en-US"), out price);
                        if (!conv)
                            conv = Double.TryParse(val.ToString(), out price);
                        newRow[prop.Name] = price;
                    }
                    //else if (prop.Name.Equals("Photo"))
                    //{
                    //    var img = val as System.Drawing.Image;
                    //    MemoryStream ms = new MemoryStream();
                    //    img.Save(ms,img.RawFormat);
                    //    newRow[prop.Name] = ms.ToArray();
                    //}
                    else
                    {
                        newRow[prop.Name] = val.ToString();
                    }
                }
                dt.Rows.Add(newRow);
            }

            using (var p = new ExcelPackage(File.Exists(path) ? new FileInfo(path) : null))
            {
                //Here setting some document properties
                //p.Workbook.Properties.Author = "Zeeshan Umar";
                p.Workbook.Properties.Title = nameBook;

                //Create a sheet
                //p.Workbook.Worksheets.Add("Sample WorkSheet");
                ExcelWorksheet ws = null;
                int colIndex = 1;
                int rowIndex = 1;
                if (p.Workbook.Worksheets.Count == 0)
                {
                    ws = p.Workbook.Worksheets.Add("Sample WorkSheet");
                    ws.Name = itemType.Name; //Setting Sheet's name
                    ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
                    ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

                    //Merging cells and create a center heading for out table
                    //ws.Cells[1, 1].Value = "Sample DataTable Export";
                    //ws.Cells[1, 1, 1, dt.Columns.Count].Merge = true;
                    //ws.Cells[1, 1, 1, dt.Columns.Count].Style.Font.Bold = true;
                    //ws.Cells[1, 1, 1, dt.Columns.Count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    colIndex = 1;

                    foreach (DataColumn dc in dt.Columns) //Creating Headings
                    {
                        var cell = ws.Cells[rowIndex, colIndex];

                        //Setting the background color of header cells to Gray
                        var fill = cell.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.Gray);


                        //Setting Top/left,right/bottom borders.
                        var border = cell.Style.Border;
                        border.Bottom.Style =
                            border.Top.Style =
                            border.Left.Style =
                            border.Right.Style = ExcelBorderStyle.Thin;

                        //Setting Value in cell
                        cell.Value = dc.ColumnName;

                        colIndex++;
                    }
                }
                else
                    ws = p.Workbook.Worksheets.First();

                rowIndex = ws.Dimension.End.Row;
                foreach (DataRow dr in dt.Rows) // Adding Data into rows
                {
                    colIndex = 1;
                    rowIndex++;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];
                        
                        if (dc.ColumnName.Contains("Price"))
                        {
                            cell.Style.Numberformat.Format = "#,##0.00";
                            double res = 0;
                            var rrr = double.TryParse(dr[dc.ColumnName].ToString(), out res);
                            cell.Value = res;
                        }
                        else if (dc.ColumnName.Equals("Photo"))
                        {
                            ws.Row(rowIndex).Height = 131;
                            ws.Column(colIndex).Width = 30;
                            //add picture to cell
                            //BinaryFormatter bf = new BinaryFormatter();
                            //Stream ms = new MemoryStream();
                            //bf.Serialize(ms, dr[dc.ColumnName]);
                            var pa = dr[dc.ColumnName].ToString();
                            var name = Guid.NewGuid().ToString();
                            using (var image = Image.FromFile(pa))
                            {
                                ExcelPicture pic = ws.Drawings.AddPicture(name, image);
                                //position picture on desired column
                                pic.From.Column = colIndex - 1; //pictureCol - 1;
                                pic.From.Row = rowIndex - 1; //currentRow - 1;
                                pic.From.ColumnOff = 9525; //ExcelHelper.Pixel2MTU(1);
                                pic.From.RowOff = 9525; // ExcelHelper.Pixel2MTU(1);
                                //set picture size to fit inside the cell
                                pic.SetSize(150, 150);
                            }
                        }
                        else
                        {
                            cell.Value = dr[dc.ColumnName];
                        }
                        //Setting Value in cell

                        var border = cell.Style.Border;
                        border.Left.Style =
                            border.Right.Style = ExcelBorderStyle.Thin;

                        //Setting borders of cell

                        colIndex++;
                    }
                }

                //colIndex = 0;
                //foreach (DataColumn dc in dt.Columns) //Creating Headings
                //{
                //    colIndex++;
                //    var cell = ws.Cells[rowIndex, colIndex];

                //    //Setting Sum Formula
                //    cell.Formula = "Sum(" +
                //                    ws.Cells[3, colIndex].Address +
                //                    ":" +
                //                    ws.Cells[rowIndex - 1, colIndex].Address +
                //                    ")";

                //    //Setting Background fill color to Gray
                //    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //    cell.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                //}

                //Generate A File with Random name
                Byte[] bin = p.GetAsByteArray();
                File.WriteAllBytes(path, bin);
            }
        }
    }
}



