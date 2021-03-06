﻿using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Linq.Expressions;
using System.Reflection;
using System.Security.Policy;
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
//using org.openqa.selenium.support.ui;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using WatiN.Core;
using Form = System.Windows.Forms.Form;
//using org.openqa.selenium;
//using org.openqa.selenium.firefox;
using OpenQA.Selenium;

using OpenQA.Selenium.Support;
using OpenQA.Selenium.Firefox;
//using org.openqa.selenium.htmlunit;

namespace ParserCatalog
{
    public partial class Form1 : Form
    {
        private List<Shop> shops;
        private List<ShopBig> shopBigs;
        private BackgroundWorker bwMain;
        private List<IWebDriver> listBrowsers;
        private CancellationTokenSource cancel;

        public Form1()
        {
            InitializeComponent();
        }
        //a[0].ParentNode.ParentNode.SelectNodes("//h3/following-sibling::p[text() and not(text()='&nbsp;') and not(text()='\r\n&nbsp;')]");
        private void Form1_Load(object sender, EventArgs e)
        {
            shops = new List<Shop>();
            shopBigs = new List<ShopBig>();
            listBrowsers = new List<IWebDriver>();
            cancel = new CancellationTokenSource();

            treeView1.Nodes.Clear();
            path.Text = Environment.CurrentDirectory;
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
            shops.Add(new Shop() { Name = "Japan-cosmetic", Url = "http://japan-cosmetic.biz" });
            shops.Add(new Shop() { Name = "Ubki-valentina", Url = "http://ubki-valentina.ru" });
            shops.Add(new Shop() { Name = "Lavira", Url = "http://lavira.ru" });
            shops.Add(new Shop() { Name = "Ekb-opt", Url = "http://ekb-opt.ru" });
            shops.Add(new Shop() { Name = "Милые платья", Url = "http://xn--80ajimbcy0a0fp7a.xn--p1ai/osen/" });
            shops.Add(new Shop() { Name = "Aimico-kids", Url = "http://aimico-kids.ru" });
            shops.Add(new Shop() { Name = "Texxit", Url = "http://texxit.ru" });
            shops.Add(new Shop() { Name = "Liora-shop", Url = "http://www.liora-shop.ru" });
            shops.Add(new Shop() { Name = "Indialove.my1", Url = "http://indialove.my1.ru" });
            //shops.Add(new Shop() { Name = "Sima-land", Url = "http://www.sima-land.ru" });
            shops.Add(new Shop() { Name = "Vsspb", Url = "http://vsspb.com/shop" });
            shops.Add(new Shop() { Name = "Stilgi", Url = "http://www.stilgi.ru" });
            shops.Add(new Shop() { Name = "Live-toys", Url = "http://live-toys.com" });
            shops.Add(new Shop() { Name = "Picoletto", Url = "http://www.picoletto.ru/" });
            shops.Add(new Shop() { Name = "Wildberries: Lucky Child", Url = "http://www.wildberries.ru/1.2149.Lucky%20Child?sle=2426" });
            shops.Add(new Shop() { Name = "Stefanika", Url = "http://www.stefanika.ru/" });
            shops.Add(new Shop() { Name = "Bus-i-nka", Url = "http://www.bus-i-nka.com/" });
            shops.Add(new Shop() { Name = "Opttextil", Url = "http://www.opttextil.ru" });
            shops.Add(new Shop() { Name = "Donnasara", Url = "http://donnasara.com.ua" });
            //shops.Add(new Shop() { Name = "Lefik", Url = "http://lefik.ru" });
            shops.Add(new Shop() { Name = "Voolya", Url = "http://voolya.com.ua" });
            shops.Add(new Shop() { Name = "Topopt", Url = "http://www.topopt.ru" });
            shops.Add(new Shop() { Name = "Besthat", Url = "http://besthat.ru" });
            shops.Add(new Shop() { Name = "Colgotki", Url = "http://www.colgotki.com" });
            shops.Add(new Shop() { Name = "I-teks-moskva", Url = "http://l-teks-moskva.ru" });
            shops.Add(new Shop() { Name = "Ivselena", Url = "http://www.ivselena.ru" });
            shops.Add(new Shop() { Name = "Amway", Url = "http://www.amway.ru" });
            shops.Add(new Shop() { Name = "Arcofam", Url = "http://arcofam.com.ua" });
            shops.Add(new Shop() { Name = "Alltextile", Url = "http://alltextile.info" });
            shops.Add(new Shop() { Name = "Limoni", Url = "http://www.limoni.ru/kupit-v-roznicu/store/" });
            shops.Add(new Shop() { Name = "Sklep.nife.pl", Url = "http://www.sklep.nife.pl/" });
            shops.Add(new Shop() { Name = "Firanka", Url = "http://firanka.ru" });
            shops.Add(new Shop() { Name = "Kupper-sport", Url = "http://www.kupper-sport.ru/catalog.php" });
            shops.Add(new Shop() { Name = "Dobroe-utro", Url = "http://www.dobroe-utro.com/katalog" });
            //shops.Add(new Shop() { Name = "Donna-saggia", Url = "http://donna-saggia.ru" });
            //shops.Add(new Shop() { Name = "Argoclassic", Url = "http://www.argoclassic.ru" });
            shops.Add(new Shop() { Name = "Tkelf", Url = "http://www.tkelf.ru/catalog/index.php" });
            //shops.Add(new Shop() { Name = "In-stylefashion.de", Url = "http://ru.in-stylefashion.de" });
            //shops.Add(new Shop() { Name = "Noch-sorochki", Url = "http://noch-sorochki.ru" });
            //shops.Add(new Shop() { Name = "Elite-cosmetics", Url = "http://www.elite-cosmetics.ru" });
            //shops.Add(new Shop() { Name = "Ora-tm", Url = "http://ora-tm.com/" });
            shops.Add(new Shop() { Name = "Danda", Url = "http://www.danda.ru/goods.php" });

            shops = shops.OrderBy(x => x.Name).ToList();

            Start.Enabled = false;
            Open.Enabled = false;
            Start.Text = "Подождите...";
            var bw = new BackgroundWorker();
            bw.DoWork += bw_DoWork;
            bw.RunWorkerAsync();
        }

        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            shopBigs.Add(new ShopBig()
            {
                Name = "Maximum",
                Url = "http://maximumufa.ru/represent/_represent_catalog/index.php",
                XPath = "//section/article/a",
                Host = "http://maximumufa.ru/represent/_represent_catalog/"
            });
            shopBigs.Add(new ShopBig()
            {
                Name = "Otoys",
                Url = "http://www.otoys.ru/",
                XPath = "//div[contains(concat(' ', @id, ' '), ' toy_types ')]/ul/li/a"
            });
            shopBigs.Add(new ShopBig()
            {
                Name = "Simaland",
                Url = "http://www.sima-land.ru/",
                XPath = "//ul[contains(concat(' ', @class, ' '), 'nav')]/li/a"
            });
            shopBigs.Add(new ShopBig()
            {
                Name = "Gap",
                Url = "http://www.gap.com",
                XPath = "//ul[contains(concat(' ', @class, ' '), 'gap_navigation')]/li/a"
            });
            shopBigs.Add(new ShopBig()
            {
                Name = "Carters",
                Url = "http://www.carters.com",
                XPath = "//li[contains(concat(' ', @class, ' '), 'topCat')]/a"
            });
            shopBigs.Add(new ShopBig()
            {
                Name = "Oshkosh",
                Url = "http://www.oshkosh.com",
                XPath = "//li[contains(concat(' ', @class, ' '), 'topCat')]/a"
            });
            shopBigs.Add(new ShopBig()
            {
                Name = "Oldnavy",
                Url = "http://oldnavy.gap.com",
                XPath = "//li[contains(concat(' ', @class, ' '), 'division')]/a"
            });
            Helpers.GetCatalog(ref shopBigs);

            //Загрузка магазина в TreeView
            int i = 0;
            foreach (var big in shopBigs)
            {
                treeView1.Invoke(() => treeView1.Nodes.Add(big.Name));
                foreach (var cat in big.CatalogList)
                {
                    treeView1.Invoke(() => treeView1.Nodes[i].Nodes.Add(cat.Name));
                }
                i++;
            }

            foreach (var sh in shops)
            {
                treeView1.Invoke(() => treeView1.Nodes.Add(sh.Name));
            }
            Start.SetPropertyThreadSafe(() => Start.Enabled, true);
            Open.SetPropertyThreadSafe(() => Open.Enabled, true);
            Start.SetPropertyThreadSafe(() => Start.Text, "Начать парсинг");
        }
        void bw_Parsing(object sender, DoWorkEventArgs e)
        {
            var pars = new List<Site>();
            var leggi = new List<string>();
            treeView1.SuspendLayout();
            cancel = new CancellationTokenSource();
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
            treeView1.ResumeLayout();
            var stL = new List<string>();
            var st = new Stopwatch();
            st.Start();
            timeStripStatus.Text = "";
            countStripStatus.Text = "Скачено 0 из " + pars.Count;
            var errors = new List<string>();
            if (File.Exists(pathSimaland.Text))
            {
                var thread = new Thread(GetSimalandFile);
                thread.Start();
            }
            Parallel.ForEach(pars, new ParallelOptions() { CancellationToken = cancel.Token }, site =>
            {
							try
							{
                var cL = new List<Category>();
                var shopUrl = site.Categories[0].Url;
                if (!site.Catalog)
                {
                    var catList = new List<Category>();
                    var cook = "";
										if (shopUrl.Contains("wildberries") || shopUrl.Contains("tvoe"))
                    {
                        cook = Helpers.GetCookiePost(shopUrl, new NameValueCollection());
                    }
                    var page = Helpers.GetHtmlDocument(shopUrl, "https://google.com",
                                                    Encoding.GetEncoding("windows-1251"), cook);

                    var query = Helpers.GetShopCatLink(shopUrl);

                    var cats = page.DocumentNode.SelectNodes(query);
                    if (shopUrl.Contains("stilgi"))
                    {
                        var tem = cats[0].InnerHtml;
                        var t = Regex.Split(tem, "dtree.add");
                        catList.AddRange(from s1 in t
                                         where s1.Contains("catalog")
                                         let beg = s1.IndexOf("/catalog")
                                         let link = s1.Substring(beg, s1.IndexOf("\"", beg + 2) - beg)
                                         let name =
                                                                         Helpers.GetEncodingCategory(s1.Substring(s1.IndexOf(",", s1.IndexOf(",") + 3) + 2,
                                                                                                         s1.LastIndexOf("\"", beg - 2) - 2 - s1.IndexOf(",", s1.IndexOf(",") + 3)))
                                         select new Category() { Url = "http://www.stilgi.ru" + link, Name = name });
                    }
                    else
                    {
                        catList = Helpers.GetListCategory(page, query, shopUrl);
                    }
                    if (shopUrl.Contains("limoni"))
                    {
                        var driver = new FirefoxDriver();
                        driver.Navigate().GoToUrl("http://www.limoni.ru/kupit-v-roznicu/store/");
                        FindDynamicElement(driver, By.XPath("//td/a[contains(concat(' ', @href, ' '), '/category/')]"), 10);
                        var doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(driver.PageSource);
                        driver.Close();
                        var catLinks = Helpers.GetItemsAttributtList(doc, "//td/a[contains(concat(' ', @href, ' '), '/category/')]", "", "href", null, null);
                        foreach (var c in catLinks)
                        {
                            catList.Add(new Category() { Url = "http://www.limoni.ru/kupit-v-roznicu/store/" + HttpUtility.HtmlDecode(c) });
                        }
                    }

                    var temp = new HashSet<string>(catList.Select(x => x.Url));
                    if (shopUrl.Contains("lavira"))
                    {
                        var rt = temp.ToList();
                        var tr = new List<string>() { rt[1], rt[2], rt[4], rt[5], rt[6], rt[7] };
                        foreach (var t in tr)
                        {
                            cL.Add(catList.FirstOrDefault(x => x.Url == t));
                        }
                    }
                    else
                    {
                        if (temp.Count != catList.Count && !shopUrl.Contains("lavira"))
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
                }
                else
                    cL = site.Categories;
                if (shopUrl.Contains("tvoe"))
                    GetTvoi(cL.Select(x => x.Url));
                else if (shopUrl.Contains("ozkan"))
                    GetOzkan(cL);
                else if (shopUrl.Contains("leggi"))
                {
                    leggi = cL.Select(x => x.Url).ToList();
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
                else if (shopUrl.Contains("npopt"))
                    GetNpopt(cL);
                else if (shopUrl.Contains("optovik-centr"))
                    GetOptovikCentr(cL);
                else if (shopUrl.Contains("japan-cosmetic"))
                    GetJapanCosmetic(cL);
                else if (shopUrl.Contains("maximum"))
                    GetMaximum(cL);
                else if (shopUrl.Contains("otoys"))
                    GetOtoys(cL);
                else if (shopUrl.Contains("ubki-valentina"))
                    GetUbkiValentina(cL);
                else if (shopUrl.Contains("lavira"))
                    GetLavira(cL);
                else if (shopUrl.Contains("ekb-opt"))
                    GetEkbOpt(cL);
                else if (shopUrl.Contains("xn--80ajimbcy0a0fp7a.xn--p1ai"))
                    GetMiliePlatia();
                else if (shopUrl.Contains("aimico-kids"))
                    GetAimicoKids(cL);
                else if (shopUrl.Contains("texxit"))
                    GetTexxit(cL);
                else if (shopUrl.Contains("liora-shop"))
                    GetLioraShopt(cL);
                else if (shopUrl.Contains("indialove"))
                    GetIndialove();
                else if (shopUrl.Contains("vsspb"))
                    GetVsspb(cL);
                else if (shopUrl.Contains("stilgi"))
                    GetStilge(cL);
                else if (shopUrl.Contains("live-toys"))
                    GetLiveToys();
                else if (shopUrl.Contains("wildberries"))
                    GetWildberries();
                else if (shopUrl.Contains("picoletto"))
                    GetPicoletto(cL);
                else if (shopUrl.Contains("stefanika"))
                    GetStefanika(cL);
                else if (shopUrl.Contains("opttextil"))
                    GetOpttextil(cL);
                else if (shopUrl.Contains("bus-i-nka"))
                    GetBusinka(cL);
                else if (shopUrl.Contains("donnasara"))
                    GetDonnasara(cL);
                else if (shopUrl.Contains("lefik"))
                    GetLefik(cL);
                else if (shopUrl.Contains("topopt"))
                    GetTopopt(cL);
                else if (shopUrl.Contains("besthat"))
                    GetBesthat(cL);
                else if (shopUrl.Contains("colgotki"))
                    GetColgotki(cL);
                else if (shopUrl.Contains("voolya"))
                    GetVoolya(cL);
                else if (shopUrl.Contains("l-teks-moskva"))
                    GetLTexsMoskva();
                else if (shopUrl.Contains("ivselena"))
                {
                    cL.AddRange(new List<Category>() { new Category() { Url = "http://www.ivselena.ru/catalog/podushki_igrushki/" }, new Category() { Url = "http://www.ivselena.ru/catalog/fartuki/" } });
                    GetIvselena(cL);
                }
                else if (shopUrl.Contains("amway"))
                    GetAmway(cL);
                else if (shopUrl.Contains("arcofam.com.ua"))
                    GetArcofam(cL);
                else if (shopUrl.Contains("alltextile"))
                    GetAlltextile(cL);
                else if (shopUrl.Contains("www.gap.com"))
                    GetGap(cL);
                else if (shopUrl.Contains("carters"))
                {
                    var shos = cL.Where(x => x.Url.Contains("shoes"));
                    if (shos.Any())
                    {
                        cL.Add(new Category() { Url = "http://www.oshkosh.com/oshkosh-girl-shoes-oshkosh-and-carters", Name = "Girl Shoes" });
                        cL.Add(new Category() { Url = "http://www.oshkosh.com/oshkosh-boy-shoes-oshkosh-and-carters", Name = "Boy Shoes" });
                    }
                    Thread.Sleep(15000);
                    GetCarters(cL);
                }
                else if (shopUrl.Contains("oshkosh"))
                {
                    var shos = cL.Where(x => x.Url.Contains("shoes"));
                    if (shos.Any())
                    {
                        cL.Add(new Category() { Url = "http://www.oshkosh.com/oshkosh-girl-shoes-oshkosh-and-carters", Name = "Girl Shoes" });
                        cL.Add(new Category() { Url = "http://www.oshkosh.com/oshkosh-boy-shoes-oshkosh-and-carters", Name = "Boy Shoes" });
                    }
                    Thread.Sleep(40000);
                    GetOshkosh(cL);
                }
                else if (shopUrl.Contains("oldnavy"))
                {
                    Thread.Sleep(60000);
                    GetOldnavy(cL);
                }
                else if (shopUrl.Contains("limoni"))
                    GetLimoni(cL);
                else if (shopUrl.Contains("sklep.nife"))
                    GetSklep(cL);
                else if (shopUrl.Contains("firanka"))
                    GetFiranka(cL);
                else if (shopUrl.Contains("kupper-sport"))
                    GetKupper(cL);
                else if (shopUrl.Contains("dobroe-utro"))
                    GetDobroeUtro(cL);
                else if (shopUrl.Contains("dobroe-utro"))
                    GetDobroeUtro(cL);
                else if (shopUrl.Contains("sima-land"))
                    GetSimaland(cL);

                stL.Add(st.Elapsed.ToString());
							}
							catch (Exception ex)
							{
							    errors.Add(site.Name + " - Ошибка: " + ex.Message);
							}
            });
            this.Invoke(new Action(() => { timeStripStatus.Text = "Время парсинга " + st.Elapsed; countStripStatus.Text = "Загружено " + (pars.Count - errors.Count - 1) + " из " + pars.Count; }));
            //timeStripStatus.Text = "Время парсинга " + st.Elapsed;
            //countStripStatus.Text = "Загружено " + (pars.Count - errors.Count - 1) + " из " + pars.Count;

            if (leggi.Any())
            {
                GetLeggi(leggi);
            }
            st.Stop();
            stL.Add(st.Elapsed.ToString());
            Start.SetPropertyThreadSafe(() => Start.Enabled, true);
            btnCancel.SetPropertyThreadSafe(() => btnCancel.Enabled, false);
            btnCancel.SetPropertyThreadSafe(() => btnCancel.Visible, false);
						
            if (errors.Count > 0)
            {
                var ss = errors.Aggregate("", (current, s) => current + (s + "\r\n"));
                this.Invoke(new Action(() => MessageBox.Show(this, ss.Trim())));
            }
            Start.SetPropertyThreadSafe(() => Start.Text, "Начать парсинг");
        }
        private void Start_Click(object sender, EventArgs e)
        {
            Start.Enabled = false;
            Start.Text = "Подождите...";
            btnCancel.Enabled = true;
            btnCancel.Visible = true;
            bwMain = new BackgroundWorker { WorkerSupportsCancellation = true };
            bwMain.DoWork += bw_Parsing;
            bwMain.RunWorkerAsync();
        }

        private void GetSimalandFile()
        {

            var cook = "";
            var links = new List<string>();
            string pathFile = pathSimaland.Text;
						if (File.Exists(pathFile))
						{
							links = File.ReadAllLines(pathFile).ToList();
							links = links.Where(x => !string.IsNullOrEmpty(x.Trim())).ToList();
						}
            var driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("http://www.sima-land.ru");
            var first = true;
            foreach (var catalog in links)
            {

                var products = new List<Simaland>();
                var docCat = Helpers.GetHtmlDocument(catalog, "http://www.sima-land.ru/", null, cook);
                var nameCat = Helpers.GetItemInnerText(docCat, "//ul[contains(concat(' ', @class, ' '), 'breadcrumbs')]/li[last()]/a");
                var cats = Helpers.GetItemsAttributtList(docCat, "//div[contains(concat(' ', @class, ' '), 'catalog-index-dd')]/ul/li/a", "", "data-qurl", null, null);
                var prod = new HashSet<string>();
                if (cats.Any())
                {
                    var temp = new List<string>();
                    if (first)
                    {
                        try
                        {
                            driver.Navigate().GoToUrl(cats[1]);
                            var limit =
                                driver.FindElement(
                                    By.XPath("//div[contains(concat(' ', @class, ' '), 'filter-limit')]/select"));
                            var selectElement = new SelectElement(limit);
                            selectElement.SelectByValue("500");
                        }
                        catch (Exception ex) { }
                        Thread.Sleep(500);
                        cook = string.Join("; ", driver.Manage().Cookies.AllCookies.Select(x => x.Name + "=" + x.Value));
                        driver.Close();
                        first = false;
                    }
                    foreach (var cat in cats)
                    {
											var l = cat;
											if (cat.Contains("//") && !cat.Contains("http"))
												l = "http:"+cat;
                        var tr = Helpers.GetProductLinks2(l, cook, catalog,
                            "//a[contains(concat(' ', @class, ' '), 'th140')]",
                            "//div[contains(concat(' ', @class, ' '), 'pagination')]/a[last()-1]", "p", null,
                            "http://www.sima-land.ru");
                        if (tr.Any())
                            temp.AddRange(tr);
                    }
                    prod = new HashSet<string>(temp);
                }
                else
                {
                    var temp = new List<string>();
                    if (first)
                    {
                        try
                        {
                            driver.Navigate().GoToUrl(catalog);
                            var limit =
                                driver.FindElement(
                                    By.XPath("//div[contains(concat(' ', @class, ' '), 'filter-limit')]/select"));
                            var selectElement = new SelectElement(limit);
                            selectElement.SelectByValue("500");
                        }
                        catch (Exception ex)
                        { }
                        Thread.Sleep(500);
                        cook = string.Join("; ", driver.Manage().Cookies.AllCookies.Select(x => x.Name + "=" + x.Value));
                        driver.Close();
                        first = false;
                    }
                    var tr = Helpers.GetProductLinks2(catalog, cook, "http://www.sima-land.ru",
                            "//a[contains(concat(' ', @class, ' '), 'th140')]",
                            "//div[contains(concat(' ', @class, ' '), 'pagination')]/a[last()-1]", "p", null,
                            "http://www.sima-land.ru");
                    if (tr.Any())
                        temp.AddRange(tr);
                    prod = new HashSet<string>(temp);
                }
                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 9 == 0)
                    {
                        Thread.Sleep(4000);
                    }
                    var doc2 = Helpers.GetHtmlDocument(res, catalog, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'main')]/div/h1");
                    artic = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), 'attributes')]/tr", "Артикул:", null).Trim();

                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var price = Helpers.GetItemsAttributt(doc2, "//div[contains(concat(' ', @class, ' '), 'actual')]", "", "data-price-value", null);

                    phs = Helpers.GetPhoto(doc2, "//img[contains(concat(' ', @class, ' '), 'card-img')]", "//div[contains(concat(' ', @class, ' '), 'thumbs')]/div/a/img", "", "", "", "src");
                    desc = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), 'attributes')]/tr | //div[contains(concat(' ', @class, ' '), 'description')]/p ", "", new List<string>() { "Артикул" }, ";;").Replace("\r", "").Replace("\n", "").Replace("\t", " ");
                    desc = Helpers.ReplaceWhiteSpace(desc);
                    desc = desc.Replace(";;", "\r\n");
                    cat = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), 'breadcrumbs')]/li/a", "", null, "/");
									
                    if (phs.Any())
                    {
                        var t = new HashSet<string>(phs.Select(x => x.Replace("/400", "/140")));
                        phs = t.Select(x => x.Replace("/140", "/1600")).ToList();
                    }
									var minCount=Helpers.GetItemInnerText(doc2,"//span[contains(concat(' ', @class, ' '), 'kMin')]");
									if (string.IsNullOrEmpty(minCount))
										minCount = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'min-qty')]/span");
                    products.Add(new Simaland()
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
                        MinCount=minCount
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }
                Helpers.SaveToFile(products, path.Text + @"\Simaland-" + nameCat + ".xlsx");
            }

            this.Invoke(new Action(() => { label3.Text = "SimaLand Загружен из файла"; }));
        }

        private void GetSimaland(IEnumerable<Category> list)
        {
            
            var cook = Helpers.GetCookiePost("http://www.sima-land.ru/suveniry/eksklyuzivnye-tovary/", new NameValueCollection() { { "rowlimit", "500" }, { "url", HttpUtility.UrlEncode("http://www.sima-land.ru/suveniry/eksklyuzivnye-tovary/") } });
            var tt = HttpUtility.UrlDecode(cook);
            tt = tt.Replace("i:0", "i:500");
            cook = HttpUtility.UrlEncode(tt);
            var driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("http://www.sima-land.ru");
						var first = true;
            foreach (var catalog in list)
            {
							var products = new List<Simaland>();
                var docCat = Helpers.GetHtmlDocument(catalog.Url, "http://www.sima-land.ru/", null, cook);
                var cats = Helpers.GetItemsAttributtList(docCat, "//div[contains(concat(' ', @class, ' '), 'catalog-index-dd')]/ul/li/a", "", "data-qurl", null, null);
                var prod = new HashSet<string>();
                if (cats.Any())
                {
									if (first)
									{
										try
										{
											driver.Navigate().GoToUrl(cats[1]);
											var limit =
													driver.FindElement(
															By.XPath("//div[contains(concat(' ', @class, ' '), 'filter-limit')]/select"));
											var selectElement = new SelectElement(limit);
											selectElement.SelectByValue("500");
										}
										catch (Exception ex) { }
										Thread.Sleep(500);
										cook = string.Join("; ", driver.Manage().Cookies.AllCookies.Select(x => x.Name + "=" + x.Value));
										driver.Close();
										first = false;
									}
                    var temp = new List<string>();
                    foreach (var cat in cats)
                    {
											var l = cat;
											if (cat.Contains("//") && !cat.Contains("http"))
												l = "http:" + cat;
											
                        var tr = Helpers.GetProductLinks2(l, cook, catalog.Url,
                                                "//a[contains(concat(' ', @class, ' '), 'th140')]", "//div[contains(concat(' ', @class, ' '), 'pagination')]/a[last()-1]", "p", null, "http://www.sima-land.ru");
                        if (tr.Any())
                            temp.AddRange(tr);
                    }
                    prod = new HashSet<string>(temp);
                }
                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 9 == 0)
                    {
                        Thread.Sleep(4000);
                    }
                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'main')]/div/h1");
                    artic = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), 'attributes')]/tr", "Артикул:", null).Trim();

                    if (string.IsNullOrEmpty(artic))
                        artic = title;

                    var price = Helpers.GetItemsAttributt(doc2, "//div[contains(concat(' ', @class, ' '), 'actual')]", "", "data-price-value", null);

                    phs = Helpers.GetPhoto(doc2, "//img[contains(concat(' ', @class, ' '), 'card-img')]", "//div[contains(concat(' ', @class, ' '), 'thumbs')]/div/a/img", "", "", "", "src");
                    desc = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), 'attributes')]/tr | //div[contains(concat(' ', @class, ' '), 'description')]/p ", "", new List<string>() { "Артикул" }, ";;").Replace("\r", "").Replace("\n", "").Replace("\t", " ");
                    desc = Helpers.ReplaceWhiteSpace(desc);
                    desc = desc.Replace(";;", "\r\n");
                    cat = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), 'breadcrumbs')]/li/a", "", null, "/");
                    if (phs.Any())
                    {
                        var t = new HashSet<string>(phs.Select(x => x.Replace("/400", "/140")));
                        phs = t.Select(x => x.Replace("/140", "/1600")).ToList();
                    }

										var minCount = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'kMin')]");
									if(string.IsNullOrEmpty(minCount))
										minCount=Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'min-qty')]/span");
										products.Add(new Simaland()
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
											MinCount = minCount
										});

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
										
                }
								Helpers.SaveToFile(products, path.Text + @"\Simaland-"+Helpers.GetEncodingCategory(catalog.Name)+".xlsx");
            }
            
            StatusStrip("Simaland");
        }

        private void GetDobroeUtro(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.dobroe-utro.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.dobroe-utro.com",
                                                "//p[contains(concat(' ', @class, ' '), 'more')]/a", "//p[contains(concat(' ', @class, ' '), 'pager')]/a[last()]", "?page=1&PAGEN_2=", null, "", true);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 4 == 0)
                    //{
                    //	Thread.Sleep(7000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'fc')]/h1");
                    artic = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'description')]/p", "Артикул", null).Replace(":", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'description')]/p", "", new List<string>() { "Артикул" });
                    var prices = Helpers.GetItemsInnerTextList(doc2, "//table[contains(concat(' ', @class, ' '), 'product_price')]/tr/td[3]", "", null, null);
										prices = prices.Select(x => x.Replace("Р", "").Trim()).ToList();
                    var siz = Helpers.GetItemsInnerTextList(doc2, "//table[contains(concat(' ', @class, ' '), 'product_price')]/tr/td[1]", "", null, null);
                    if (siz.Any())
                    {
                        size = string.Join("; ", siz.Select(x => x.Contains("-") ? x.Remove(x.IndexOf("-")).Trim() : x));
                    }
                    if (desc.Length > 0 && desc.Contains("Оптовикам по заказу"))
                    {
                        desc = desc.Replace("Оптовикам по заказу от 50.000 т.р. для упрощения работы, высылаем наличие товара на складе (обращаться по тел.: +7 4932 591196)", "").Trim();
                        desc = Helpers.ReplaceWhiteSpace(desc);
                    }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @id, ' '), 'zoom')]", "//div[contains(concat(' ', @class, ' '), 'prod_preview')]/ul/li/em/img");
										phs = phs.Select(x => !x.Contains("http") ? "http:" + x : x).ToList();
                    cat = Helpers.GetItemsInnerText(doc2, "//p[contains(concat(' ', @class, ' '), 'navigator')]/a", "", new List<string>() { "Главная" }, "/");
                    col = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'prod_preview')]/ul/li", "", null, "; ");
                    var temo = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), 'prod_preview')]/ul/li");
                    if (siz.Count > 1)
                    {
                        for (var i = 0; i < siz.Count; i++)
                        {
                            products.Add(new Product()
                            {
                                Url = i == 0 ? res : "",
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = prices[i],
                                CategoryPath = cat,
                                Size = siz[i].Contains("-") ? siz[i].Remove(siz[i].IndexOf("-")).Trim() : siz[i],
                                Photos = phs,
                            });
                            if (i == 0)
                            {
                                col = "";
                                desc = "";
                                cat = "";
                                phs = new List<string>();
                                title = "";
                            }
                        }
                    }
                    else
                    {
                        products.Add(new Product()
                        {
                            Url = res,
                            Article = artic,
                            Color = col,
                            Description = desc,
                            Name = title,
                            Price = prices.Any()?prices[0]:"",
                            CategoryPath = cat,
                            Size = size,
                            Photos = phs,
                        });
                    }
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\DobroeUtro.xlsx",false,false,false);
            StatusStrip("DobroeUtro");
        }
        private void GetKupper(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.kupper-sport.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.kupper-sport.ru",
                                                "//a[contains(concat(' ', @class, ' '), 'model')]", "//a[contains(concat(' ', @class, ' '), 'nump')][last()]", "&page=", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//td/h1");
                    artic = Helpers.GetItemInnerText(doc2, "//tr/td[2]/font[contains(concat(' ', @class, ' '), 'orange')]");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    var price = ""; //Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'productPrice')]").Trim();
                    size = Helpers.GetItemsAttributt(doc2, "//img[contains(concat(' ', @src, ' '), '/sizes')]", "", "alt", null, "; ");
                    var pr = doc2.DocumentNode.SelectNodes("//img[contains(concat(' ', @src, ' '), '/sizes')]");
                    if (pr != null)
                    {
                        var f = pr[0].ParentNode.SelectNodes("font");
                        if (f != null)
                            price = f[0].InnerText;
                    }
                    var d = doc2.DocumentNode.SelectNodes("//tr/td[2]/font[contains(concat(' ', @class, ' '), 'orange')]");
                    if (d != null)
                    {
                        var de = d[0].ParentNode.InnerText;
                        de = de.Replace(artic, "").Trim();
                        //if(de.Contains("Информация")){
                        //	desc=de.Remove(de.IndexOf("Информация")).Trim();
                        //}else
                        //	desc=de;
                        desc = Helpers.ReplaceWhiteSpace(de);
                    }
                    phs = Helpers.GetPhoto(doc2, "", "//div[contains(concat(' ', @id, ' '), 'big')]/img", "", "http://www.kupper-sport.ru");
                    col = Helpers.GetItemsAttributt(doc2, "//div[contains(concat(' ', @id, ' '), 'big')]/img", "", "alt", null, "; ");
                    cat = Helpers.GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), 'navigator')]", "", null, "/");
                    if (cat.Length > 0)
                    {
                        cat = Helpers.ReplaceWhiteSpace(cat.Replace("//", "/"));
                        cat = cat.Replace("карта сайта/спортивные костюмы kupper/", "");
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
                        state = desc.ToLower().Contains("нет в наличие") ? "no_sale" : ""
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Kupper.xlsx");
            StatusStrip("Kupper");
        }
        private void GetFiranka(IEnumerable<Category> list)
        {
            var products = new List<ProductPrices>();
            var cook = Helpers.GetCookiePost("http://firanka.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var link = catalog.Url.Replace("group", "groupphoto");
                var prod = Helpers.GetProductLinks(link, cook, "http://firanka.ru",
                                                "//a[contains(concat(' ', @class, ' '), 'more')]", null);
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

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var title = "";
                    var phs = new List<string>();
                    var str = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'prod_name')]");
                    var ars = Regex.Split(str, ",").ToList();

                    artic = ars[0];
                    ars.RemoveAt(0);
                    if (ars.Any())
                    {
                        title = ars[0];
                        ars.RemoveAt(0);
                        foreach (var a in ars)
                        {
                            if (a.Contains("цвет"))
                                col = a;
                        }
                    }
                    if (col.Length > 0 && col.Contains(":"))
                        col = col.Substring(col.IndexOf(":") + 1);
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'prod_prop-block')]/table/tr[2]/td[2]").Trim();
                    var price2 = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'prod_prop-block')]/table/tr[3]/td[2]").Replace(".", ",").Replace(" ", "").Trim();
                    desc += Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'prod_prop-block')]/div", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), 'prod_pict-big')]/a", "", "http://firanka.ru");
                    cat = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), 'partner')]/tr/td/a", "", new List<string>() { "Главная" }, "/");

                    products.Add(new ProductPrices()
                    {
                        Url = res,
                        Article = artic,
                        Color = col.Trim(),
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs,
                        Prices = new List<string>() { price2 }
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Firanka.xlsx");
            StatusStrip("Firanka");
        }
        private void GetSklep(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.sklep.nife.pl/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var session = cook.Substring(cook.IndexOf("PHPS"), cook.IndexOf(";", cook.IndexOf("PHPS") + 1) - cook.IndexOf("PHPS"));
                var link = catalog.Url.Remove(catalog.Url.IndexOf("?") + 1) + session;
                var doc = Helpers.GetHtmlDocument(HttpUtility.HtmlEncode(link), "http://www.sklep.nife.pl/", null, cook);
                var pages = Helpers.GetItemsAttributt(doc, "//a[contains(concat(' ', @href, ' '), 'show_all=1')]", "", "href", null);
                if (pages.Length > 0)
                {
                    pages = pages.Remove(pages.IndexOf("PHPS"));
                    pages = "http://www.sklep.nife.pl/" + pages + session;
                }
                else
                    pages = link;
                var prod = Helpers.GetProductLinks(pages, cook, "http://www.sklep.nife.pl/",
                                                "//span[contains(concat(' ', @class, ' '), 'lista_produktow_nazwa')]/a", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 4 == 0)
                    //{
                    //	Thread.Sleep(7000);
                    //}
                    var link2 = res.Remove(res.IndexOf("PHPS"));
                    //link2 = link2; //+ session;
                    var doc2 = Helpers.GetHtmlDocument(link2, catalog.Url, Encoding.GetEncoding("iso-8859-2"), cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'txt_nazwa')]");
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'txt_cena_cyfry')]").Trim();

                    col = Helpers.GetTextReplaceTags(doc2, "//select[contains(concat(' ', @name, ' '), 'kolory')]", new List<string>() { "wybierz kolor" }, "", "; ");
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'txt_opis')]", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "", "//img[contains(concat(' ', @id, ' '), 'obrazek')]", "", "http://www.sklep.nife.pl");
                    var win = Encoding.GetEncoding("windows-1251");
                    byte[] winBytes = win.GetBytes(catalog.Name);
                    cat = Encoding.GetEncoding("iso-8859-2").GetString(winBytes, 0, winBytes.Length);
                    cat = HttpUtility.HtmlDecode(cat);
                    if (col.Length > 0)
                    {
                        var reg = Regex.Matches(doc2.DocumentNode.InnerHtml, @"\|\d{2,3}\|")
                                                                                            .Cast<Match>()
                                                                                            .Select(m => m.Value.Replace("|", ""))
                                                                                            .ToList();
                        size = string.Join("; ", reg);
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
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Sklep.xlsx");
            StatusStrip("Sklep");
        }
        private void GetLimoni(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.limoni.ru/", new NameValueCollection());
            FirefoxDriver driver = null;
            bool table = false;
            var frPr = new FirefoxProfile();
            foreach (var catalog in list)
            {
                driver = new FirefoxDriver(frPr);
                driver.Navigate().GoToUrl(catalog.Url);
                FindDynamicElement(driver, By.XPath("//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a"), 10);
                Thread.Sleep(2500);
                //if (!table)
                //{
                try
                {
                    driver.FindElement(By.XPath("//div[contains(concat(' ', @class, ' '), 'ecwid-results-topPanel-viewAsPanel')]/div[last()]")).Click();
                    table = true;
                }
                catch (Exception ex) { }
                //}
                Thread.Sleep(2500);
                var doc = new HtmlAgilityPack.HtmlDocument();
                var host = "http://www.limoni.ru/kupit-v-roznicu/store/";
                doc.LoadHtml(driver.PageSource);
                var catLinks = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a", "", "href", null, null);
                var prod = new HashSet<string>();
                var temp2 = new List<string>();
                temp2.AddRange(catLinks.Select(x => host + HttpUtility.HtmlDecode(x)));
                var catal = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @style, ' '), 'display: inline-block;')]/a", "", "href", null, null);
                if (catal.Any())
                {
                    catal = catal.Select(x => host + HttpUtility.HtmlDecode(x)).ToList();
                    var lastCount2 = catal.Count;
                    foreach (var c in catal)
                    {

                        try
                        {
                            driver.Navigate().GoToUrl(c);
                            Thread.Sleep(2500);

                            doc.LoadHtml(driver.PageSource);
                        }
                        catch (Exception ex)
                        {
                            driver = new FirefoxDriver(frPr);
                            driver.Navigate().GoToUrl(c);
                            doc.LoadHtml(driver.PageSource);
                        }
                        var catal2 = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @style, ' '), 'display: inline-block;')]/a", "", "href", null, null);
                        var tr = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a", "", "href", null, null);
                        if (tr.Any())
                        {
                            temp2.AddRange(tr.Select(x => host + HttpUtility.HtmlDecode(x)));
                            var next = Helpers.GetItemsInnerText(doc,
                                    "//span[contains(concat(' ', @class, ' '), 'ecwid-pager-link ecwid-pager-link-enabled')]", "", null, ";;");
                            if (next.Length > 0)
                            {
                                var countPage = Regex.Split(next, ";;").Count();
                                for (var i = 0; i < countPage - 1; i++)
                                {
                                    driver.FindElements(By.XPath("//span[contains(concat(' ', @class, ' '), 'ecwid-pager-link ecwid-pager-link-enabled')][last()]"))[0].Click();
                                    Thread.Sleep(2500);
                                    doc.LoadHtml(driver.PageSource);
                                    var tr3 = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a", "", "href", null, null);
                                    if (tr3.Any())
                                        temp2.AddRange(tr3.Select(x => host + HttpUtility.HtmlDecode(x)));
                                }
                            }
                        }
                        if (catal2.Any())
                        {
                            catal2 = catal2.Select(x => host + HttpUtility.HtmlDecode(x)).ToList();
                            foreach (var c2 in catal2)
                            {
                                try
                                {
                                    driver.Navigate().GoToUrl(c2);
                                    Thread.Sleep(2500);
                                    doc.LoadHtml(driver.PageSource);
                                }
                                catch (Exception ex)
                                {
                                    driver = new FirefoxDriver(frPr);
                                    driver.Navigate().GoToUrl(c2);
                                    doc.LoadHtml(driver.PageSource);
                                }
                                var tr2 = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a", "", "href", null, null);
                                if (tr2.Any())
                                {
                                    temp2.AddRange(tr2.Select(x => host + HttpUtility.HtmlDecode(x)));
                                    var next = Helpers.GetItemsInnerText(doc,
                            "//span[contains(concat(' ', @class, ' '), 'ecwid-pager-link ecwid-pager-link-enabled')]", "", null, ";;");
                                    if (next.Length > 0)
                                    {
                                        var countPage = Regex.Split(next, ";;").Count();
                                        for (var i = 0; i < countPage - 1; i++)
                                        {
                                            driver.FindElements(By.XPath("//span[contains(concat(' ', @class, ' '), 'ecwid-pager-link ecwid-pager-link-enabled')][last()]"))[0].Click();
                                            Thread.Sleep(2500);
                                            doc.LoadHtml(driver.PageSource);
                                            var tr3 = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a", "", "href", null, null);
                                            if (tr3.Any())
                                                temp2.AddRange(tr3.Select(x => host + HttpUtility.HtmlDecode(x)));
                                        }
                                    }
                                }
                            }
                        }
                    }


                }
                else if (temp2.Count == 20)
                {
                    var next = Helpers.GetItemsInnerText(doc,
                                            "//span[contains(concat(' ', @class, ' '), 'ecwid-pager-link ecwid-pager-link-enabled')]", "", null, ";;");
                    if (next.Length > 0)
                    {
                        var countPage = Regex.Split(next, ";;").Count();
                        for (var i = 0; i < countPage - 1; i++)
                        {
                            driver.FindElements(By.XPath("//span[contains(concat(' ', @class, ' '), 'ecwid-pager-link ecwid-pager-link-enabled')][last()]"))[0].Click();
                            Thread.Sleep(2500);
                            doc.LoadHtml(driver.PageSource);
                            var tr3 = Helpers.GetItemsAttributtList(doc, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-productNameLink')]/a", "", "href", null, null);
                            if (tr3.Any())
                                temp2.AddRange(tr3.Select(x => host + HttpUtility.HtmlDecode(x)));
                        }
                    }
                }
                prod = new HashSet<string>(temp2);

                if (prod.Count == 0)
                    continue;
                int countRequest = 0;
                foreach (var res in prod)
                {
                    var doc2 = new HtmlAgilityPack.HtmlDocument();

                    try
                    {
                        driver.Navigate().GoToUrl(res);
                        Thread.Sleep(2500);
                        //driver.FindElement(By.ClassName("ecwid-productBrowser-head")).Click();
                        FindDynamicElement(driver, By.XPath("//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-price')]"), 10);
                        doc2.LoadHtml(driver.PageSource);
                    }
                    catch (Exception ex)
                    {
                        driver = new FirefoxDriver(frPr);
                        driver.Navigate().GoToUrl(res);
                        Thread.Sleep(2000);
                        doc2.LoadHtml(driver.PageSource);
                    }

                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-head')]");
                    artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-sku')]").Replace("Артикул", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-price')]").Replace("p.", "").Replace(" ", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'gwt-HTML')]/p", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "", "//img[contains(concat(' ', @class, ' '), 'gwt-Image')]");
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-categoryPath-categoryLink')]/a", "", new List<string>() { "Магазин" }, "/");
                    var st = Helpers.GetItemInnerText(doc2,
                                    "//div[contains(concat(' ', @class, ' '), 'ecwid-productBrowser-details-inStockLabel')]");

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
                        state = !st.Contains("В наличии") ? "no_sale" : ""
                    });
                    driver.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(2000);
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }
                driver.Close();

            }
            Helpers.SaveToFile(products, path.Text + @"\Limoni.xlsx");
            StatusStrip("Limoni");
        }
        private void GetOldnavy(IEnumerable<Category> list)
        {

            var cook = Helpers.GetCookiePost("http://oldnavy.gap.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var products = new List<Product>();
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://oldnavy.gap.com",
                                                                                                                                                "//a[contains(concat(' ', @class, ' '), 'category')]", "", "#pageId=", null);
                var driver = new FirefoxDriver();
                listBrowsers.Add(driver);
                if (prod.Count > 0)
                {
                    var temp = new List<string>();
                    var gg = 0;
                    foreach (var cat in prod)
                    {

                        try
                        {
                            //driver.Manage().Timeouts().SetScriptTimeout(TimeSpan.FromSeconds(10));
                            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(15));
                            // Thread.Sleep(2000);
                            driver.Navigate().GoToUrl(cat);
                            // Thread.Sleep(10000);
                            FindDynamicElement(driver, By.ClassName("productItemName"), 10);
                            Thread.Sleep(1500);
                            driver.FindElement(By.Id("totalItemCountDivSB")).Click();
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                driver.Close();
                            }
                            catch (Exception e) { }
                            driver = new FirefoxDriver();
                            driver.Navigate().GoToUrl(cat);
                            //FindDynamicElement(driver, By.Id("totalItemCountDivSB"), 10);
                            //driver.FindElement(By.Id("totalItemCountDivSB")).Click();
                        }
                        var source = driver.PageSource;
                        var doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(source);
                        var pagesLink = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'productItemName')]", "", "href", null, null);
                        if (pagesLink.Count == 0)
                            continue;
                        pagesLink = pagesLink.Select(x => "http://oldnavy.gap.com" + HttpUtility.HtmlDecode(x)).ToList();
                        var oidLink = pagesLink.Where(x => x.Contains("oid=")).ToList();
                        foreach (var o in oidLink)
                        {
                            pagesLink.Remove(o);
                        }
                        temp.AddRange(pagesLink);
                        try
                        {
                            foreach (var t1 in oidLink)
                            {
                                Thread.Sleep(700);
                                driver.Navigate().GoToUrl(t1);
                                FindDynamicElement(driver,
                                                By.XPath("//div[contains(concat(' ', @id, ' '), 'outfitImages')]"), 5);
                                var html = driver.PageSource;
                                var reg =
                                                Regex.Matches(html, "strProductId: \"\\w*\"")
                                                                .Cast<Match>()
                                                                .Select(m => m.Value.Substring(m.Value.IndexOf("\"") + 1).Replace("\"", ""))
                                                                .ToList();
                                var link = HttpUtility.HtmlDecode(driver.Url);
                                link = link.Replace("outfit.do", "product.do");
                                link = link.Remove(link.IndexOf("&oid"));
                                temp.AddRange(reg.Select(r => link + "&pid=" + r));
                            }

                            Thread.Sleep(5000);
                            if (driver.Url.Contains("oid"))
                                driver.Navigate().GoToUrl(cat);
                            var page = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), 'pagePaginator hideMe')]");
                            if (page == null)
                            {
                                try
                                {
                                    driver.FindElement(By.XPath("//a[contains(concat(' ', @title, ' '), 'Next page')]")).Click();
                                }
                                catch (Exception ex)
                                {
                                    Thread.Sleep(1000);
                                    driver.Navigate().Back();
                                    continue;
                                }
                                Thread.Sleep(1000);
                                source = driver.PageSource;
                                doc = new HtmlAgilityPack.HtmlDocument();
                                doc.LoadHtml(source);//pagePaginator hideMe
                                pagesLink = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'productItemName')]", "", "href", null, null);
                                pagesLink = pagesLink.Select(x => "http://oldnavy.gap.com" + HttpUtility.HtmlDecode(x)).ToList();
                                oidLink = pagesLink.Where(x => x.Contains("oid=")).ToList();
                                foreach (var o in oidLink)
                                {
                                    pagesLink.Remove(o);
                                }
                                temp.AddRange(pagesLink);
                                var linkpag1 = driver.Url;
                                foreach (var t1 in oidLink)
                                {
                                    Thread.Sleep(700);
                                    driver.Navigate().GoToUrl(t1);
                                    FindDynamicElement(driver,
                                                    By.XPath("//div[contains(concat(' ', @id, ' '), 'outfitImages')]"), 5);
                                    var html = driver.PageSource;
                                    var reg =
                                                    Regex.Matches(html, "strProductId: \"\\w*\"")
                                                                    .Cast<Match>()
                                                                    .Select(m => m.Value.Substring(m.Value.IndexOf("\"") + 1).Replace("\"", ""))
                                                                    .ToList();
                                    var link = HttpUtility.HtmlDecode(driver.Url);
                                    link = link.Replace("outfit.do", "product.do");
                                    link = link.Remove(link.IndexOf("&oid"));
                                    temp.AddRange(reg.Select(r => link + "&pid=" + r));
                                }
                                Thread.Sleep(1000);
                                if (driver.Url.Contains("oid"))
                                    driver.Navigate().GoToUrl(linkpag1);
                                FindDynamicElement(driver, By.XPath("//a[contains(concat(' ', @title, ' '), 'Next page')]"), 5);

                                bool p2 = false;
                                try
                                {
                                    var page2 = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), 'pagePaginator hideMe')]");
                                    if (page2 == null)
                                    {
                                        driver.FindElement(By.XPath("//a[contains(concat(' ', @title, ' '), 'Next page')]")).Click();
                                        p2 = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Thread.Sleep(1000);
                                    driver.Navigate().Back();
                                    continue;
                                }
                                if (p2)
                                {
                                    Thread.Sleep(1000);
                                    source = driver.PageSource;
                                    doc = new HtmlAgilityPack.HtmlDocument();
                                    doc.LoadHtml(source);
                                    pagesLink = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'productItemName')]", "", "href", null, null);
                                    pagesLink = pagesLink.Select(x => "http://oldnavy.gap.com" + HttpUtility.HtmlDecode(x)).ToList();
                                    oidLink = pagesLink.Where(x => x.Contains("oid=")).ToList();
                                    foreach (var o in oidLink)
                                    {
                                        pagesLink.Remove(o);
                                    }
                                    temp.AddRange(pagesLink);

                                    foreach (var t1 in oidLink)
                                    {
                                        Thread.Sleep(700);
                                        driver.Navigate().GoToUrl(t1);
                                        FindDynamicElement(driver,
                                                        By.XPath("//div[contains(concat(' ', @id, ' '), 'outfitImages')]"), 5);
                                        var html = driver.PageSource;
                                        var reg =
                                                        Regex.Matches(html, "strProductId: \"\\w*\"")
                                                                        .Cast<Match>()
                                                                        .Select(m => m.Value.Substring(m.Value.IndexOf("\"") + 1).Replace("\"", ""))
                                                                        .ToList();
                                        var link = HttpUtility.HtmlDecode(driver.Url);
                                        link = link.Replace("outfit.do", "product.do");
                                        link = link.Remove(link.IndexOf("&oid"));
                                        temp.AddRange(reg.Select(r => link + "&pid=" + r));
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                    prod = new HashSet<string>(temp);
                }
                else
                    continue;
                driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(1));
                driver.Manage().Timeouts().SetScriptTimeout(TimeSpan.FromSeconds(10));
                int countRequest = 0;
                foreach (var res in prod)
                {
                    try
                    {
                        driver.Navigate().GoToUrl(res);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            driver.Close();
                        }
                        catch (Exception ee) { }
                        Thread.Sleep(80000);
                        driver = new FirefoxDriver();
                    }
                    var col = "";
                    FindDynamicElement(driver, By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]"), 5);
                    var images = 0;
                    try
                    {
                        images = driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]")).Count;
                    }
                    catch (Exception ex) { }
                    var phs = new List<string>();
                    for (var i = 0; i < images; i++)
                    {
                        try
                        {
                            FindDynamicElement(driver,
                                            By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]"), 5);
                            //var images2 =
                            driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]"))[i].Click();
                            //var img = images2[i];
                            //img.Click();
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();

                            var pathImg =
                                            driver.FindElement(By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"))
                                                            .GetAttribute("src");
                            phs.Add(pathImg);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    var imagCol = 0;
                    try
                    {
                        imagCol = driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')]")).Count;
                    }
                    catch (Exception ex)
                    {

                    }
                    for (var i = 1; i < imagCol && imagCol > 1; i++)
                    {
                        try
                        {
                            FindDynamicElement(driver, By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')]"), 5);
                            //var images2 =
                            driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')]"))[i].Click();
                            //var img = images2[i];
                            //img.Click();
                            try
                            {
                                driver.FindElement(By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]")).Click();
                            }
                            catch (Exception ex)
                            {
                            }
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                            FindDynamicElement(driver,
                                                        By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"), 5);
                            Thread.Sleep(300);
                            var text =
                                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'textColor')]")).Text;
                            if (text.Length > 0)
                                col += text.Trim() + "; ";
                            var pathImg =
                                            driver.FindElement(By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"))
                                                            .GetAttribute("src");
                            phs.Add(pathImg);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    try
                    {
                        driver.FindElement(By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')][1]")).Click();
                        FindDynamicElement(driver,
                                                            By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"), 5);
                        Thread.Sleep(300);
                        driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                        driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                        var img =
                                        driver.FindElement(By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"))
                                                        .GetAttribute("src");
                        phs.Add(img);
                    }
                    catch (Exception ex) { }
                    var doc2 = new HtmlAgilityPack.HtmlDocument();

                    try
                    {
                        var ttt = driver.PageSource;
                        doc2.LoadHtml(ttt);
                    }
                    catch (Exception ex)
                    {
                        try { driver.Close(); }
                        catch (Exception e) { }
                        driver = new FirefoxDriver();
                        driver.Navigate().GoToUrl(res);
                        Thread.Sleep(1500);
                        var ttt = driver.PageSource;
                        doc2.LoadHtml(ttt);
                    }
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";

                    var title = Helpers.GetItemInnerText(doc2,
                                    "//span[contains(concat(' ', @class, ' '), 'productName')]");
                    artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), 'productNumber')]").ToLower().Replace("prices may vary", "").Trim();
                    col +=
                                    Helpers.GetItemsInnerText(doc2, "//label[contains(concat(' ', @for, ' '), 'colorSwatch')]",
                                                    "", null, "; ").Replace("product image", "").Trim();
                    var price =
                                    Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'priceText')][1]")
                                                    .Replace("$", "")
                                                    .Trim();
                    if (Regex.Matches(price, ".").Count > 1)
                    {
                        price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'priceText')][1]/span")
                                                    .Replace("$", "")
                                                    .Trim();
                    }
                    if (price.Trim().Length == 0)
                    {
                        price =
                                        Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'salePrice')][1]")
                                                        .Replace("$", "")
                                                        .Trim();
                    }

                    size =
                                    Helpers.GetItemsInnerText(doc2, "//label[contains(concat(' ', @for, ' '), 'size1Swatch')]",
                                                    "", null, "; ").Replace("Select size", "");
                    if (size.Length > 0)
                        size = Helpers.ReplaceWhiteSpace(size.Replace("Size", "").Replace("currently not available", "").Replace("on order. Estimated ship", ""));
                    if (size.Contains("date"))
                    {
                        var beg = size.IndexOf("date");
                        var end = size.IndexOf(";", beg);
                        var temp = "";
                        if (end <= beg)
                            temp = size.Remove(beg);
                        else
                            temp = size.Remove(beg, size.IndexOf(";", beg) - beg);
                        size = Helpers.ReplaceWhiteSpace(size.Replace(temp, ""));
                    }
                    if (col.Length > 0)
                        col = Helpers.ReplaceWhiteSpace(col.ToLower().Replace("select", "").Replace("color:", "")).Replace(" ;", ";");
                    desc = Helpers.GetItemsInnerText(doc2,
                                    "//div[contains(concat(' ', @id, ' '), 'tabWindow')]/ul/li", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    cat = Helpers.GetEncodingCategory(catalog.Name) + "/" +
                                            Helpers.GetItemInnerText(doc2,
                                                            "//a[contains(concat(' ', @class, ' '), 'categorySelected')]");
                    var st = doc2.DocumentNode.InnerText.Contains("Sorry, this item is currently out of stock");

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = HttpUtility.HtmlDecode(cat),
                        Size = size,
                        Photos = phs,
                        state = st ? "no_sale" : ""
                    });

                    countRequest++;
                    driver.Manage().Cookies.DeleteAllCookies();


                }
                driver.Close();
                listBrowsers.Remove(driver);
                Helpers.SaveToFile(products, path.Text + @"\Oldnavy-" + Helpers.ReplaceWhiteSpace(Helpers.GetEncodingCategory(catalog.Name).Replace("\n", " ")) + ".xlsx", false, false);
            }

            StatusStrip("Oldnavy");
        }
        private void GetOshkosh(IEnumerable<Category> list)
        {
            var cook = Helpers.GetCookiePost("http://www.oshkosh.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var products = new List<Product>();
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.oshkosh.com",
                                                                                "//ul[contains(concat(' ', @class, ' '), 'category-group')]/li/a", null);

                if (prod.Count > 0)
                {
                    var html = new FirefoxDriver();
                    listBrowsers.Add(html);
                    var temp = new List<string>();
                    foreach (var cat in prod)
                    {
                        try
                        {
                            var c = cat.Replace("http://www.oshkosh.com/", "");
                            c = c.Remove(c.IndexOf("?"));
                            var doc = Helpers.GetHtmlDocument(cat, "http://www.oshkosh.com/", null, cook);

                            html.Navigate().GoToUrl(cat);
                            FindDynamicElement(html, By.XPath("//div[contains(concat(' ', @class, ' '), 'breadcrumb')]/ul/li[last()]"), 5);
                            //var link2 =
                            html.FindElement(
                                            By.XPath("//div[contains(concat(' ', @class, ' '), 'breadcrumb')]/ul/li[last()]")).Click();
                            try
                            {
                                //var link =
                                html.FindElement(
                                                By.XPath("//a[contains(concat(' ', @href, ' '), '&sz=')]")).Click();
                            }
                            catch (Exception ex) { }
                            FindDynamicElement(html, By.XPath("//a[contains(concat(' ', @class, ' '), 'name-link')][last()]"), 5);
                            Thread.Sleep(1000);
                            var source = html.PageSource;
                            doc.LoadHtml(source);
                            var tr = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'name-link')]", "", "href", null, null);
                            if (!tr.Contains("http://www.oshkosh.com"))
                                tr = tr.Select(x => "http://www.oshkosh.com" + x).ToList();
                            if (tr.Any())
                                temp.AddRange(tr);
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                html.Close();
                            }
                            catch (Exception z) { }
                            Thread.Sleep(50000);
                            html = new FirefoxDriver();
                        }
                    }
                    html.Close();
                    listBrowsers.Remove(html);
                    prod = new HashSet<string>(temp);
                }
                else if (catalog.Url.Contains("shoes"))
                {
                    var html = new FirefoxDriver();
                    listBrowsers.Add(html);
                    var temp = new List<string>();
                    try
                    {
                        var doc = new HtmlAgilityPack.HtmlDocument();

                        html.Navigate().GoToUrl(catalog.Url);
                        FindDynamicElement(html, By.XPath("//div[contains(concat(' ', @class, ' '), 'results-hits')]/span/a"), 5);
                        try
                        {
                            var link =
                                            html.FindElement(
                                                            By.XPath("//div[contains(concat(' ', @class, ' '), 'results-hits')]/span/a"));
                            link.Click();
                        }
                        catch (Exception ex)
                        {
                        }
                        FindDynamicElement(html, By.XPath("//a[contains(concat(' ', @class, ' '), 'name-link')][last()]"), 5);
                        Thread.Sleep(1000);
                        var source = html.PageSource;
                        doc.LoadHtml(source);
                        var tr = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'name-link')]", "",
                                        "href", null, null);
                        if (!tr.Contains("http://www.oshkosh.com"))
                            tr = tr.Select(x => "http://www.oshkosh.com" + x).ToList();
                        if (tr.Any())
                            temp.AddRange(tr);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            html.Close();
                        }
                        catch (Exception z) { }
                        Thread.Sleep(50000);
                        html = new FirefoxDriver();
                    }

                    html.Close();
                    listBrowsers.Remove(html);
                    prod = new HashSet<string>(temp);
                }
                if (!prod.Any())
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 5 == 0)
                    {
                        Thread.Sleep(7000);
                    }

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var lis = Helpers.GetItemInnerText(doc2,
                                    "//div[contains(concat(' ', @id, ' '), 'product-set-list')]");
                    if (lis.Length > 0)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), 'product-name')]");
                    artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), 'productID')]");
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), 'price')]").Replace("$", "").Trim();
                    size = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), 'swatches size')]/li/a", "", null, "; ");
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @itemprop, ' '), 'description')]|//div[contains(concat(' ', @class, ' '), 'additional')]/ul/li", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'product-image')]");
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'breadcrumb')]/ul/li/a", "", new List<string>() { "Home" }, "/");
                    col = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), 'swatches color')]/li/a", "", null, "; ");
                    if (cat.Length == 0)
                        cat = Helpers.GetEncodingCategory(catalog.Name);
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
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }
                Helpers.SaveToFile(products, path.Text + @"\Oshkosh-" + Helpers.ReplaceWhiteSpace(Helpers.GetEncodingCategory(catalog.Name).Replace("\n", " ")) + ".xlsx");
            }

            StatusStrip("Oshkosh");
        }
        private void GetCarters(IEnumerable<Category> list)
        {

            var cook = Helpers.GetCookiePost("http://www.carters.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var products = new List<Product>();
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.carters.com",
                                                                                "//ul[contains(concat(' ', @class, ' '), 'category-group')]/li/a", null);

                if (prod.Count > 0)
                {
                    var html = new FirefoxDriver();
                    listBrowsers.Add(html);
                    var temp = new List<string>();
                    foreach (var cat in prod)
                    {
                        try
                        {
                            var c = cat.Replace("http://www.carters.com/", "");
                            c = c.Remove(c.IndexOf("?"));
                            var doc = Helpers.GetHtmlDocument(cat, "http://www.carters.com/", null, cook);

                            html.Navigate().GoToUrl(cat);
                            FindDynamicElement(html, By.XPath("//div[contains(concat(' ', @class, ' '), 'results-hits')]/span/a"), 5);
                            try
                            {
                                var link =
                                                html.FindElement(
                                                                By.XPath("//div[contains(concat(' ', @class, ' '), 'results-hits')]/span/a"));
                                link.Click();
                            }
                            catch (Exception ex)
                            {
                            }
                            FindDynamicElement(html, By.XPath("//a[contains(concat(' ', @class, ' '), 'name-link')][last()]"), 5);
                            Thread.Sleep(1000);
                            var source = html.PageSource;
                            doc.LoadHtml(source);
                            var tr = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'name-link')]", "",
                                            "href", null, null);
                            if (!tr.Contains("http://www.carters.com"))
                                tr = tr.Select(x => "http://www.carters.com" + x).ToList();
                            //var tr=Helpers.GetProductLinks2(cat, cook, "http://www.carters.com",
                            //        "//a[contains(concat(' ', @class, ' '), 'name-link')]", "//div[contains(concat(' ', @class, ' '), 'pagination-temp')]/a[last()]", null);
                            if (tr.Any())
                                temp.AddRange(tr);
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                html.Close();
                            }
                            catch (Exception z) { }
                            Thread.Sleep(50000);
                            html = new FirefoxDriver();
                        }
                    }
                    html.Close();
                    listBrowsers.Remove(html);
                    prod = new HashSet<string>(temp);
                }
                else if (catalog.Url.Contains("shoes"))
                {
                    var html = new FirefoxDriver();
                    listBrowsers.Add(html);
                    var temp = new List<string>();
                    try
                    {
                        var doc = new HtmlAgilityPack.HtmlDocument();

                        html.Navigate().GoToUrl(catalog.Url);
                        FindDynamicElement(html, By.XPath("//div[contains(concat(' ', @class, ' '), 'results-hits')]/span/a"), 5);
                        try
                        {
                            var link =
                                            html.FindElement(
                                                            By.XPath("//div[contains(concat(' ', @class, ' '), 'results-hits')]/span/a"));
                            link.Click();
                        }
                        catch (Exception ex)
                        {
                        }
                        FindDynamicElement(html, By.XPath("//a[contains(concat(' ', @class, ' '), 'name-link')][last()]"), 5);
                        Thread.Sleep(1000);
                        var source = html.PageSource;
                        doc.LoadHtml(source);
                        var tr = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'name-link')]", "",
                                        "href", null, null);
                        if (!tr.Contains("http://www.carters.com"))
                            tr = tr.Select(x => "http://www.carters.com" + x).ToList();
                        if (tr.Any())
                            temp.AddRange(tr);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            html.Close();
                        }
                        catch (Exception z) { }
                        Thread.Sleep(50000);
                        html = new FirefoxDriver();
                    }

                    html.Close();
                    listBrowsers.Remove(html);
                    prod = new HashSet<string>(temp);
                }
                if (!prod.Any())
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 5 == 0)
                    {
                        Thread.Sleep(7000);
                    }

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var lis = Helpers.GetItemInnerText(doc2,
                                    "//div[contains(concat(' ', @id, ' '), 'product-set-list')]");
                    if (lis.Length > 0)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), 'product-name')]");
                    artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), 'productID')]");
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @itemprop, ' '), 'price')]").Replace("$", "").Trim();
                    size = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), 'swatches size')]/li/a", "", null, "; ");
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @itemprop, ' '), 'description')]|//div[contains(concat(' ', @class, ' '), 'additional')]/ul/li", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'product-image')]");
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'breadcrumb')]/ul/li/a", "", new List<string>() { "Home" }, "/");
                    col = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), 'swatches color')]/li/a", "", null, "; ");
                    if (cat.Length == 0)
                        cat = Helpers.GetEncodingCategory(catalog.Name);
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
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }
                Helpers.SaveToFile(products, path.Text + @"\Carters-" + Helpers.ReplaceWhiteSpace(Helpers.GetEncodingCategory(catalog.Name).Replace("\n", " ")) + ".xlsx");
            }

            StatusStrip("Carters");
        }
        private void GetGap(IEnumerable<Category> list)
        {

            var cook = Helpers.GetCookiePost("http://www.gap.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var products = new List<Product>();
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.gap.com",
                                                                                "//a[contains(concat(' ', @class, ' '), 'category')]", "", "#pageId=", null);
                var driver = new FirefoxDriver();
                listBrowsers.Add(driver);
                if (prod.Count > 0)
                {
                    var temp = new List<string>();
                    var gg = 0;
                    foreach (var cat in prod)
                    {

                        try
                        {
                            //driver.Manage().Timeouts().SetScriptTimeout(TimeSpan.FromSeconds(10));
                            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(15));
                            // Thread.Sleep(2000);
                            driver.Navigate().GoToUrl(cat);
                            // Thread.Sleep(10000);
                            FindDynamicElement(driver, By.ClassName("productItemName"), 10);
                            Thread.Sleep(1500);
                            driver.FindElement(By.Id("totalItemCountDivSB")).Click();
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                driver.Close();
                            }
                            catch (Exception e) { }
                            driver = new FirefoxDriver();
                            driver.Navigate().GoToUrl(cat);
                            //FindDynamicElement(driver, By.Id("totalItemCountDivSB"), 10);
                            //driver.FindElement(By.Id("totalItemCountDivSB")).Click();
                        }
                        var source = driver.PageSource;
                        var doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(source);
                        var pagesLink = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'productItemName')]", "", "href", null, null);
                        if (pagesLink.Count == 0)
                            continue;
                        pagesLink = pagesLink.Select(x => "http://www.gap.com" + HttpUtility.HtmlDecode(x)).ToList();
                        var oidLink = pagesLink.Where(x => x.Contains("oid=")).ToList();
                        foreach (var o in oidLink)
                        {
                            pagesLink.Remove(o);
                        }
                        temp.AddRange(pagesLink);
                        try
                        {
                            foreach (var t1 in oidLink)
                            {
                                Thread.Sleep(700);
                                driver.Navigate().GoToUrl(t1);
                                FindDynamicElement(driver,
                                                By.XPath("//div[contains(concat(' ', @id, ' '), 'outfitImages')]"), 5);
                                var html = driver.PageSource;
                                var reg =
                                                Regex.Matches(html, "strProductId: \"\\w*\"")
                                                                .Cast<Match>()
                                                                .Select(m => m.Value.Substring(m.Value.IndexOf("\"") + 1).Replace("\"", ""))
                                                                .ToList();
                                var link = HttpUtility.HtmlDecode(driver.Url);
                                link = link.Replace("outfit.do", "product.do");
                                link = link.Remove(link.IndexOf("&oid"));
                                temp.AddRange(reg.Select(r => link + "&pid=" + r));
                            }

                            Thread.Sleep(5000);
                            if (driver.Url.Contains("oid"))
                                driver.Navigate().GoToUrl(cat);
                            var page = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), 'pagePaginator hideMe')]");
                            if (page == null)
                            {
                                try
                                {
                                    driver.FindElement(By.XPath("//a[contains(concat(' ', @title, ' '), 'Next page')]")).Click();
                                }
                                catch (Exception ex)
                                {
                                    Thread.Sleep(1000);
                                    driver.Navigate().Back();
                                    continue;
                                }
                                Thread.Sleep(1000);
                                source = driver.PageSource;
                                doc = new HtmlAgilityPack.HtmlDocument();
                                doc.LoadHtml(source);//pagePaginator hideMe
                                pagesLink = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'productItemName')]", "", "href", null, null);
                                pagesLink = pagesLink.Select(x => "http://www.gap.com" + HttpUtility.HtmlDecode(x)).ToList();
                                oidLink = pagesLink.Where(x => x.Contains("oid=")).ToList();
                                foreach (var o in oidLink)
                                {
                                    pagesLink.Remove(o);
                                }
                                temp.AddRange(pagesLink);
                                var linkpag1 = driver.Url;
                                foreach (var t1 in oidLink)
                                {
                                    Thread.Sleep(700);
                                    driver.Navigate().GoToUrl(t1);
                                    FindDynamicElement(driver,
                                                    By.XPath("//div[contains(concat(' ', @id, ' '), 'outfitImages')]"), 5);
                                    var html = driver.PageSource;
                                    var reg =
                                                    Regex.Matches(html, "strProductId: \"\\w*\"")
                                                                    .Cast<Match>()
                                                                    .Select(m => m.Value.Substring(m.Value.IndexOf("\"") + 1).Replace("\"", ""))
                                                                    .ToList();
                                    var link = HttpUtility.HtmlDecode(driver.Url);
                                    link = link.Replace("outfit.do", "product.do");
                                    link = link.Remove(link.IndexOf("&oid"));
                                    temp.AddRange(reg.Select(r => link + "&pid=" + r));
                                }
                                Thread.Sleep(1000);
                                if (driver.Url.Contains("oid"))
                                    driver.Navigate().GoToUrl(linkpag1);
                                FindDynamicElement(driver, By.XPath("//a[contains(concat(' ', @title, ' '), 'Next page')]"), 5);

                                bool p2 = false;
                                try
                                {
                                    var page2 = doc.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), 'pagePaginator hideMe')]");
                                    if (page2 == null)
                                    {
                                        driver.FindElement(By.XPath("//a[contains(concat(' ', @title, ' '), 'Next page')]")).Click();
                                        p2 = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Thread.Sleep(1000);
                                    driver.Navigate().Back();
                                    continue;
                                }
                                if (p2)
                                {
                                    Thread.Sleep(1000);
                                    source = driver.PageSource;
                                    doc = new HtmlAgilityPack.HtmlDocument();
                                    doc.LoadHtml(source);
                                    pagesLink = Helpers.GetItemsAttributtList(doc, "//a[contains(concat(' ', @class, ' '), 'productItemName')]", "", "href", null, null);
                                    pagesLink = pagesLink.Select(x => "http://www.gap.com" + HttpUtility.HtmlDecode(x)).ToList();
                                    oidLink = pagesLink.Where(x => x.Contains("oid=")).ToList();
                                    foreach (var o in oidLink)
                                    {
                                        pagesLink.Remove(o);
                                    }
                                    temp.AddRange(pagesLink);

                                    foreach (var t1 in oidLink)
                                    {
                                        Thread.Sleep(700);
                                        driver.Navigate().GoToUrl(t1);
                                        FindDynamicElement(driver,
                                                        By.XPath("//div[contains(concat(' ', @id, ' '), 'outfitImages')]"), 5);
                                        var html = driver.PageSource;
                                        var reg =
                                                        Regex.Matches(html, "strProductId: \"\\w*\"")
                                                                        .Cast<Match>()
                                                                        .Select(m => m.Value.Substring(m.Value.IndexOf("\"") + 1).Replace("\"", ""))
                                                                        .ToList();
                                        var link = HttpUtility.HtmlDecode(driver.Url);
                                        link = link.Replace("outfit.do", "product.do");
                                        link = link.Remove(link.IndexOf("&oid"));
                                        temp.AddRange(reg.Select(r => link + "&pid=" + r));
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                    prod = new HashSet<string>(temp);
                }
                else
                    continue;
                driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(1));
                driver.Manage().Timeouts().SetScriptTimeout(TimeSpan.FromSeconds(10));
                int countRequest = 0;
                //var prod = new List<string>() { "http://www.gap.com/browse/product.do?cid=65618&vid=1&pid=535576072" };
                foreach (var res in prod)
                {
                    try
                    {
                        driver.Navigate().GoToUrl(res);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            driver.Close();
                        }
                        catch (Exception ee) { }
                        Thread.Sleep(80000);
                        driver = new FirefoxDriver();
                    }
                    var col = "";
                    FindDynamicElement(driver, By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]"), 5);
                    var images = 0;
                    try
                    {
                        images = driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]")).Count;
                    }
                    catch (Exception ex) { }
                    var phs = new List<string>();
                    for (var i = 0; i < images; i++)
                    {
                        try
                        {
                            FindDynamicElement(driver,
                                            By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]"), 5);
                            //var images2 =
                            driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]"))[i].Click();
                            //var img = images2[i];
                            //img.Click();
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();

                            var pathImg =
                                            driver.FindElement(By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"))
                                                            .GetAttribute("src");
                            phs.Add(pathImg);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    var imagCol = 0;
                    try
                    {
                        imagCol = driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')]")).Count;
                    }
                    catch (Exception ex)
                    {

                    }
                    for (var i = 1; i < imagCol && imagCol > 1; i++)
                    {
                        try
                        {
                            FindDynamicElement(driver, By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')]"), 5);
                            //var images2 =
                            driver.FindElements(By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')]"))[i].Click();
                            //var img = images2[i];
                            //img.Click();
                            try
                            {
                                driver.FindElement(By.XPath("//input[contains(concat(' ', @id, ' '), 'thumbImage')]")).Click();
                            }
                            catch (Exception ex)
                            {
                            }

                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                            FindDynamicElement(driver,
                                                    By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"), 5);
                            Thread.Sleep(300);
                            var text =
                                            driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'textColor')]")).Text;
                            if (text.Length > 0)
                                col += text.Trim() + "; ";
                            var pathImg =
                                            driver.FindElement(By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"))
                                                            .GetAttribute("src");
                            phs.Add(pathImg);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    try
                    {
                        driver.FindElement(By.XPath("//input[contains(concat(' ', @id, ' '), 'colorSwatch')][1]")).Click();
                        FindDynamicElement(driver,
                                                            By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"), 5);
                        Thread.Sleep(300);
                        driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                        driver.FindElement(By.XPath("//div[contains(concat(' ', @id, ' '), 'dragLayer')]")).Click();
                        var img =
                                        driver.FindElement(By.XPath("//img[contains(concat(' ', @id, ' '), 'zoomImg')]"))
                                                        .GetAttribute("src");
                        phs.Add(img);
                    }
                    catch (Exception ex) { }
                    var doc2 = new HtmlAgilityPack.HtmlDocument();

                    try
                    {
                        var ttt = driver.PageSource;
                        doc2.LoadHtml(ttt);
                    }
                    catch (Exception ex)
                    {
                        try { driver.Close(); }
                        catch (Exception e) { }
                        driver = new FirefoxDriver();
                        driver.Navigate().GoToUrl(res);
                        Thread.Sleep(1500);
                        var ttt = driver.PageSource;
                        doc2.LoadHtml(ttt);
                    }
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";

                    var title = Helpers.GetItemInnerText(doc2,
                                    "//span[contains(concat(' ', @class, ' '), 'productName')]");
                    artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), 'productNumber')]").ToLower().Replace("prices may vary", "").Trim();
                    col +=
                                    Helpers.GetItemsInnerText(doc2, "//label[contains(concat(' ', @for, ' '), 'colorSwatch')]",
                                                    "", null, "; ").Replace("product image", "").Trim();
                    var price =
                                    Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'priceText')][1]")
                                                    .Replace("$", "")
                                                    .Trim();
                    if (Regex.Matches(price, ".").Count > 1)
                    {
                        price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), 'priceText')][1]/span")
                                                    .Replace("$", "")
                                                    .Trim();
                    }
                    if (price.Trim().Length == 0)
                    {
                        price =
                                        Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'salePrice')][1]")
                                                        .Replace("$", "")
                                                        .Trim();
                    }

                    size =
                                    Helpers.GetItemsInnerText(doc2, "//label[contains(concat(' ', @for, ' '), 'size1Swatch')]",
                                                    "", null, "; ").Replace("Select size", "");
                    if (size.Length > 0)
                        size = Helpers.ReplaceWhiteSpace(size.Replace("Size", "").Replace("currently not available", "").Replace("on order. Estimated ship", ""));
                    if (size.Contains("date"))
                    {
                        var beg = size.IndexOf("date");
                        var end = size.IndexOf(";", beg);
                        var temp = "";
                        if (end <= beg)
                            temp = size.Remove(beg);
                        else
                            temp = size.Remove(beg, size.IndexOf(";", beg) - beg);
                        size = Helpers.ReplaceWhiteSpace(size.Replace(temp, ""));
                    }
                    if (col.Length > 0)
                        col = Helpers.ReplaceWhiteSpace(col.ToLower().Replace("select", "").Replace("color:", "")).Replace(" ;", ";");
                    desc = Helpers.GetItemsInnerText(doc2,
                                    "//div[contains(concat(' ', @id, ' '), 'tabWindow')]/ul/li", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    cat = Helpers.GetEncodingCategory(catalog.Name) + "/" +
                                            Helpers.GetItemInnerText(doc2,
                                                            "//a[contains(concat(' ', @class, ' '), 'categorySelected')]");
                    var st = doc2.DocumentNode.InnerText.Contains("Sorry, this item is currently out of stock");

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = title,
                        Price = price,
                        CategoryPath = HttpUtility.HtmlDecode(cat),
                        Size = size,
                        Photos = phs,
                        state = st ? "no_sale" : ""
                    });

                    countRequest++;
                    driver.Manage().Cookies.DeleteAllCookies();


                }
                driver.Close();
                listBrowsers.Remove(driver);
                Helpers.SaveToFile(products, path.Text + @"\Gap-" + Helpers.ReplaceWhiteSpace(Helpers.GetEncodingCategory(catalog.Name).Replace("\n", " ")) + ".xlsx", false, false);
            }

            StatusStrip("Gap");
        }
        private void GetAlltextile(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://alltextile.info/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://alltextile.info",
                                                "//div[contains(concat(' ', @class, ' '), 'browseProductContainer')]/div/a/a", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 4 == 0)
                    //{
                    //	Thread.Sleep(7000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var artic = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), 'product_desc')]");
                    if (title.ToLower().Contains("арт"))
                    {
                        artic = title.ToLower().Remove(0, title.ToLower().IndexOf("арт"));
                        artic = artic.Remove(0, 4);
                        artic = artic.Substring(0, artic.IndexOf(".")).Trim();
                        artic = artic.ToUpper();
                    }

                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'productPrice')]").Trim();
                    size = Helpers.GetItemsInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'size')]", "", null);

                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'highslide')]");
                    cat = Helpers.GetEncodingCategory(catalog.Name);


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
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Alltextile.xlsx", false, false, false);
            StatusStrip("Alltextile");
        }
        private void GetArcofam(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = "";//Helpers.GetCookiePost("http://arcofam.com.ua/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "/page/all", cook, "http://arcofam.com.ua",
                                                "//div[contains(concat(' ', @class, ' '), 'wrapper')]/a", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 4 == 0)
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
                    var title = Helpers.GetItemInnerText(doc2, "//article/header/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'product-serial')]").Replace("Артикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'product-price-data')]").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), 'product-full-desc')]/p", "", null);
                    var desc2 = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @class, ' '), 'user-inner')]/p", null, ";;");
                    if (desc2.Length > 0)
                        desc = desc + "\r\n" + desc2.Replace("\r", "").Replace("\n", "").Replace(";;", "\r\n").Trim();
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'fancy-img')]");
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'section-bread-crumbs')]/a", "", new List<string>() { "Каталог посуды" }, "/");


                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc.Trim(),
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs,
                    });

                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Arcofam.xlsx");
            StatusStrip("Arcofam");
        }
        private void GetAmway(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.amway.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.amway.ru",
                                                "//div[contains(concat(' ', @class, ' '), 'title')]/a", null);
                if (prod.Any())
                {
                    var temp = new List<string>();
                    foreach (var l in prod)
                    {
                        var doc = Helpers.GetProductLinks(l + "?table=catalog_category_products&size=100", cook, "http://www.amway.ru",
                        "//div[contains(concat(' ', @class, ' '), 'product_name')]/a", null);
                        if (doc.Any())
                            temp.AddRange(doc.ToList());
                    }
                    prod = new HashSet<string>(temp);
                }
                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 4 == 0)
                    //{
                    //		Thread.Sleep(7000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//td[contains(concat(' ', @class, ' '), 'product_details_content')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'sku')]").Replace("Артикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'nowrap')]").Replace("RUR", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'content_area')]/p", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'MagicZoomPlus')]", "//div[contains(concat(' ', @class, ' '), 'content_carousel')]/ul/li/a/img", "http://www.amway.ru", "http://www.amway.ru");
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'breadcrumb')]/a", "", new List<string>() { "На главную" }, "/");
                    var status = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'text_big')]/a", "Временно отсутствующая", null);
                    var count = Helpers.GetPhoto(doc2, "", "//span[contains(concat(' ', @class, ' '), 'img_wrapper')]/img", "", "http://www.amway.ru");
                    phs = new HashSet<string>(phs.Select(x => x.Replace("_mini_", "_max_"))).ToList();
                    if (count.Any())
                    {
                        var stt = doc2.DocumentNode.InnerHtml.Substring(doc2.DocumentNode.InnerHtml.IndexOf("variantNames"));
                        var reg = Regex.Matches(stt, @"variantSkus\[(\d*)\]='\w*'").Cast<Match>().Select(m => m.Value.Substring(m.Value.IndexOf("=") + 1).Replace("'", "")).ToList();
                        var title2 = Regex.Matches(stt, @"variantNames\[(\d*)\]='\w*|[А-ЯЁ][а-яё]*'").Cast<Match>().Select(m => m.Value.Substring(m.Value.IndexOf("=") + 1).Replace("'", "")).ToList();

                        var price2 = Helpers.GetItemsInnerTextList(doc2, "//span[contains(concat(' ', @class, ' '), 'nowrap')]", "", null, null);
                        price2 = price2.Select(x => x.Replace("RUR", "").Trim()).ToList();
                        phs.AddRange(count);
                        title += "-";
                        for (int i = 0; i < reg.Count; i++)
                        {
                            products.Add(new Product()
                            {
                                Url = i == 0 ? res : "",
                                Article = reg[i],
                                Color = col,
                                Description = desc,
                                Name = title + title2[i],
                                Price = price2[i],
                                CategoryPath = cat,
                                Size = size,
                                Photos = phs,
                                state = status.Length > 0 ? "no_sale" : ""
                            });
                            phs = new List<string>();
                            title = "";
                            desc = "";
                            cat = "";
                        }
                    }
                    else
                    {
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
                            state = status.Length > 0 ? "no_sale" : ""
                        });
                    }
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Amway.xlsx", false, false, false);
            StatusStrip("Amway");
        }
        private void GetIvselena(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.ivselena.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                //var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.ivselena.ru",
                //		"//a[contains(concat(' ', @class, ' '), ' over ')]",
                //		"//div[contains(concat(' ', @id, ' '), ' paginationControl ')]/a[last()-1] ", "/0/page/", null);

                //if (prod.Count == 0)
                //	continue;

                int countRequest = 0;
                //foreach (var res in prod)
                //{
                //try
                //{
                //if (countRequest % 4 == 0)
                //{
                //	Thread.Sleep(7000);
                //}

                var doc2 = Helpers.GetHtmlDocument(catalog.Url, "http://www.ivselena.ru", Encoding.GetEncoding("windows-1251"), cook);
                if (doc2 == null)
                    continue;
                var col = "";
                var size = "";
                var desc = "";
                var cat = "";
                var phs = new List<string>();
                var title = Helpers.GetItemInnerText(doc2, "//h2[contains(concat(' ', @id, ' '), ' pagetitle ')]");
                var artic = title;
                phs = Helpers.GetPhoto(doc2, "", "//div[contains(concat(' ', @class, ' '), 'catalogue-description')]//img", "", "http://www.ivselena.ru");
                desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'catalogue-description')]/p|//div[contains(concat(' ', @class, ' '), 'catalogue-description')]/div", "", new List<string>());
                cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), 'breadcrumb')]/a", "", new List<string>() { "Главная" }, "/");
                var cont = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), 'content-table')]/tbody/tr");
                var price = "";
                if (cont != null)
                {
                    foreach (var c in cont)
                    {
                        var siz = c.SelectNodes("td[1]");
                        if (siz != null)
                            size = siz[0].InnerText.Replace(title, "").Trim();
                        var pr = c.SelectNodes("td[2]");
                        if (pr != null)
                            price = pr[0].InnerText.Replace("руб.", "").Trim();
                        var colLink = c.SelectNodes("td[3]/a");
                        if (colLink != null)
                        {
                            var link = colLink[0].Attributes["href"].Value;
                            if (!link.Contains("ivselena"))
                                link = "http://www.ivselena.ru" + link;
                            var doc3 = Helpers.GetHtmlDocument(link, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                            var phs2 = Helpers.GetPhoto(doc3, "//a[contains(concat(' ', @rel, ' '), 'gallery-images')]", "", "http://www.ivselena.ru");
                            if (phs2.Any())
                                phs.AddRange(phs2);
                            var desc2 = Helpers.GetItemsInnerText(doc3, "//td/div[font or span]", "", null);
                            if (desc2.Length > 0)
                                desc += "\r\n" + desc2;
                            if (phs2.Count == 36)
                            {
                                link = link + "?PAGEN_1=2";
                                var doc4 = Helpers.GetHtmlDocument(link, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                                phs2 = Helpers.GetPhoto(doc4, "//a[contains(concat(' ', @rel, ' '), 'gallery-images')]", "", "http://www.ivselena.ru");
                                col = Helpers.GetItemsInnerText(doc4, "//div[contains(concat(' ', @class, ' '), 'name')]/a", "", null, "; ");
                                if (phs2.Any())
                                    phs.AddRange(phs2);
                                if (phs2.Count == 36)
                                {
                                    link = link.Replace("?PAGEN_1=2", "?PAGEN_1=3");
                                    var doc5 = Helpers.GetHtmlDocument(link, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                                    phs2 = Helpers.GetPhoto(doc5, "//a[contains(concat(' ', @rel, ' '), 'gallery-images')]", "", "http://www.ivselena.ru");
                                    col += " " + Helpers.GetItemsInnerText(doc5, "//div[contains(concat(' ', @class, ' '), 'name')]/a", "", null, "; ");
                                    if (phs2.Any())
                                        phs.AddRange(phs2);
                                }
                            }
                            col += " " + Helpers.GetItemsInnerText(doc3, "//div[contains(concat(' ', @class, ' '), 'name')]/a", "", null, "; ");
                            if (col.Length > 0)
                                col = col.Trim();
                            products.Add(new Product()
                            {
                                Url = catalog.Url,
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = price,
                                CategoryPath = cat,
                                Size = size,
                                Photos = phs,
                            });
                        }
                        else
                        {
                            products.Add(new Product()
                            {
                                Url = catalog.Url,
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = price,
                                CategoryPath = cat,
                                Size = size,
                                Photos = phs,
                            });
                        }
                        price = "";
                        size = "";
                        phs = new List<string>();
                    }
                    //}


                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Ivselena.xlsx", false, false, false);
            StatusStrip("Ivselena");
        }
        private void GetLTexsMoskva()
        {
            var products = new List<ProductPrices>();
            var cook = Helpers.GetCookiePost("http://l-teks-moskva.ru/", new NameValueCollection());


            var prod = Helpers.GetProductLinks2("http://l-teks-moskva.ru/product_list", cook, "http://l-teks-moskva.ru",
                                            "//a[contains(concat(' ', @class, ' '), 'b-product-line__image-wrapper')]",
                                            "//div[contains(concat(' ', @class, ' '), 'b-pager')]/a[last()-1] ", "/page_", null);

            int countRequest = 0;
            foreach (var res in prod)
            {
                //try
                //{
                if (countRequest % 3 == 0)
                {
                    Thread.Sleep(7000);
                }

                var doc2 = Helpers.GetHtmlDocument(res, "http://l-teks-moskva.ru/product_list", null, cook);
                if (doc2 == null)
                    continue;
                var col = "";
                var size = "";
                var desc = "";
                var cat = "";
                var phs = new List<string>();
                var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), 'b-product__name')]");
                var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), 'b-product__sku')]").Replace("Код:", "").Trim();
                var price = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @class, ' '), 'b-product__price')]").Replace("руб.", "").Replace("/", "").Replace(" ", "").Replace("упаковка", "").Trim();
                desc = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), 'b-product-info')]/tr", "", null, ";;");
                if (string.IsNullOrEmpty(artic))
                    artic = title;
                phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'imagebox')]");
                size = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'b-user-content')]/table/tr[2]/td[3]").Replace(",", ";").Trim();
                cat = Helpers.GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), 'b-breadcrumb__link')]", "", new List<string>() { "Л-ТЕКС" }, "/");
                col = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/p", "Цвет", null).Replace(":", "").Trim();
                var price2 = Helpers.GetItemInnerText(doc2,
                                "//div[contains(concat(' ', @class, ' '), 'b-user-content')]/table/tr[2]/td[5]");
                desc = Helpers.ReplaceWhiteSpace(Helpers.ReplaceWhiteSpace(desc)).Replace("\r", "").Replace("\n", "").Replace(";;", "\r\n");
                products.Add(new ProductPrices()
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
                    Prices = new List<string>() { price2 }
                });
                countRequest++;
                //}
                //catch (Exception ex) { }
            }


            Helpers.SaveToFile(products, path.Text + @"\LTexsMoskva.xlsx");
            StatusStrip("LTexsMoskva");
        }
        private void GetVoolya(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://voolya.com.ua/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://voolya.com.ua",
                                "//a[contains(concat(' ', @class, ' '), ' over ')]",
                                "//div[contains(concat(' ', @id, ' '), ' paginationControl ')]/a[last()-1] ", "/0/page/", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 4 == 0)
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
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' title ')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' art ')]").Replace("Арт.", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' price red ')]").Replace("грн. (оптовая цена)", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", new List<string>() { "<!-" });
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), 'otherImages')]/a", "", "http://voolya.com.ua");
                    size = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' size ')]", "", null, "; ").Replace("Размеры:", "").Replace("есть", "").Trim();
                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    col = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/p", "Цвет", null).Replace(":", "").Trim();
                    if (size.Length == 0)
                    {
                        size = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/p", "Размер", null).Replace(":", "").Trim();
                    }
                    if (size.Contains(","))
                        size = size.Replace(",", "; ").Replace("  ", " ").Replace(" ;", ";");

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
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Voolya.xlsx");
            StatusStrip("Voolya");
        }
        private void GetColgotki(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.colgotki.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "&limit=400", cook, "http://www.colgotki.com",
                                "//a[text()='[Подробнее...]']", null);
                if (prod.Count == 0)
                {
                    var cat = Helpers.GetProductLinks(catalog.Url, cook, "http://www.colgotki.com", "//td[contains(concat(' ', @align, ' '), ' center ')]/a", null);
                    if (cat.Any())
                    {
                        var temp = new List<string>();
                        foreach (var c in cat)
                        {
                            var tr = Helpers.GetProductLinks(c + "&limit=400", cook, "http://www.colgotki.com",
                                                                            "//a[text()='[Подробнее...]']", null);
                            if (tr.Any())
                                temp.AddRange(tr.ToList());
                            else
                            {
                                var cat2 = Helpers.GetProductLinks(c, cook, "http://www.colgotki.com", "//td[contains(concat(' ', @align, ' '), ' center ')]/a", null);
                                foreach (var c1 in cat2)
                                {
                                    var tr2 = Helpers.GetProductLinks(c1 + "&limit=400", cook, "http://www.colgotki.com",
                                                                                    "//a[text()='[Подробнее...]']", null);
                                    if (tr2.Any())
                                        temp.AddRange(tr2.ToList());
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
                    //if (countRequest % 3 == 0)
                    //{
                    //    Thread.Sleep(7000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//td/h1");
                    var artic = title;
                    var price = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @id, ' '), ' old-field ')]").Replace("руб.", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//td/p", "", new List<string>() { "Размер" });
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]");
                    cat = Helpers.GetItemsInnerText(doc2, "//a[contains(concat(' ', @id, ' '), 'active')]", "", new List<string>() { "Главная" }, "/");
                    cat = catalog.Name + "/" + cat;
                    if (desc.Contains("Цвет"))
                    {
                        try
                        {
                            //    var begin = desc.IndexOf("Цвет");
                            //    col = desc.Substring(begin, desc.IndexOf("\r\n", begin + 2)).Replace("Цвет:", "");
                            //    desc = desc.Replace(desc.Substring(begin, desc.IndexOf("\r\n", begin + 2)), "").Trim();
                            //}

                            //if (desc.Contains("Цвет"))
                            //{
                            var temp = desc.Substring(desc.IndexOf("Цвет:")).Trim();
                            if (col.Length > 0 && temp.Length > 6)
                                col += col.Trim() + "; " + temp.Replace("Цвет:", "").Trim().Replace(",", ";");
                            else if (temp.Length > 6)
                                col = temp.Replace("Цвет:", "").Trim().Replace(",", ";");
                            if (col.Length > 0)
                                desc = desc.Replace(col, "").Trim();
                            desc = desc.Replace("Цвет:", "").Trim();
                        }
                        catch (Exception ex)
                        {
                        }

                    }
                    if (col.Length == 0)
                    {
                        col = Helpers.GetItemsInnerText(doc2, "//td/li|//td/ul/li", "", null, "; ");
                        if (col.Length == 0)
                        {
                            col = Helpers.GetItemInnerText(doc2, "//dt").Replace(",", ";");
                        }
                    }
                    size = Helpers.GetItemsInnerText(doc2, "//td/p", "Размер", null).Trim().Replace("ы:", "").Replace(":", "").Trim().Replace("\t", "; ");
                    if (!size.Contains(";"))
                    {
                        if (size.Contains(", "))
                            size = size.Replace(",", ";");
                        else if (size.Contains(","))
                            size = size.Replace(",", "; ");
                        else
                            size = size.Replace(" ", "; ");
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
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Colgotki.xlsx");
            StatusStrip("Colgotki");
        }
        private void GetBesthat(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://besthat.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://besthat.ru",
                                "//a[contains(concat(' ', @class, ' '), ' product-title ')]",
                                "//div[contains(concat(' ', @class, ' '), 'pagination')][last()]/a[last()]",
                                 "page-", null);//"//div[contains(concat(' ', @class, ' '), ' pagination ')]/a[last()]",

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 6 == 0)
                    {
                        Thread.Sleep(6000);
                    }

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' mainbox-title ')]");
                    var artic = title;
                    if (artic.Contains("арт."))
                    {
                        artic = artic.Substring(artic.IndexOf("арт") + 4).Trim();
                    }
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' price ')]/span").Replace("руб", "").Replace(",", "").Trim();
                    desc = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @id, ' '), ' content_block_description ')]", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), 'image-border')]/a/img[contains(concat(' ', @id, ' '), 'det_img')]", "//img[contains(concat(' ', @class, ' '), 'object-image')] ", "http://besthat.ru", "http://besthat.ru", "", "src");
                    var content = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @class, ' '), ' product-list-field ')]", null);
                    //size = Helpers.GetTextReplaceTags(doc2, "//select[contains(concat(' ', @name, ' '), 'product_data')]", null);
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/a", "", new List<string>() { "Главная" }, "/");
                    if (content.Contains("Цвет"))
                    {
                        var begin = content.IndexOf("Цвет");
                        try
                        {
                            col =
                                            content.Substring(begin,
                                                            content.LastIndexOf("\r\n", content.IndexOf(":", begin + 6)) - begin)
                                                            .ToLower()
                                                            .Replace("цвет:", "")
                                                            .Trim()
                                                            .Replace("\r\n", "; ");
                        }
                        catch (Exception ex)
                        {
                            col = content.ToLower().Replace("цвет:", "").Trim().Replace("\r\n", "; ");
                        }
                        col = col.Replace("; количество:", "");
                    }
                    if (content.Contains("Размер"))
                    {
                        var begin = content.IndexOf("Размер");
                        var end = content.IndexOf("Цвет");
                        if (end == -1 || end < begin)
                            end = content.Length;
                        try
                        {
                            size =
                                            content.Substring(begin,
                                                            end - begin)
                                                            .ToLower()
                                                            .Replace("размер:", "")
                                                            .Trim()
                                                            .Replace("\r\n", "; ");
                        }
                        catch (Exception ex)
                        {
                            size = content.ToLower().Replace("размер:", "").Trim().Replace("\r\n", "; ");
                        }
                        size = size.Replace("; количество:", "");
                    }
                    if (phs.Any())
                    {
                        phs = phs.Select(x => x.Replace("30/30", "1000").Replace("150", "1000")).ToList();
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
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }
            }
            Helpers.SaveToFile(products, path.Text + @"\Besthat.xlsx");
            StatusStrip("Besthat");
        }
        private void GetLefik(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://lefik.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://lefik.ru",
                                "//div[contains(concat(' ', @class, ' '), ' title ')]/a",
                                "//span[contains(concat(' ', @class, ' '), ' part-link ')][last()]/a ", "?page=", null);

                if (prod.Count == 0)
                    continue;

                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' product-title ')]");
                    if (title.Length == 0)
                        continue;
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-sku ')]/div[contains(concat(' ', @class, ' '), ' value ')]").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' price-value ')]").Replace("руб", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' b-description ')]/p", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'fancybox')]");
                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    size = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @class, ' '), ' b-product-properties ')]",
                                    null, "", "; ").Replace("/", ",");
                    col = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @class, ' '), 'selector-wrapper')]",
                                    null, "Цвет", "; ");
                    var trr = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), 'l-content')]/script[last()]");
                    //option_names
                    //[{"id":144214,"title":"\u0420\u0430\u0437\u043c\u0435\u0440","position":1},
                    //{"id":144320,"title":"\u041f\u0440\u043e\u0438\u0437\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044c","position":3},
                    //{"id":144321,"title":"\u0426\u0432\u0435\u0442","position":4},
                    //{"id":145593,"title":"\u0421\u043e\u0441\u0442\u0430\u0432","position":5}]

                    //Размер \u0420\u0430\u0437\u043c\u0435\u0440  144214
                    //Цвет 
                    //Ткань \u0422\u043a\u0430\u043d\u044c
                    //Производитель \u041f\u0440\u043e\u0438\u0437\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044c  170884 
                    //if (desc.Length > 0)
                    //{
                    //    desc = Helpers.ReplaceWhiteSpace(desc.Replace("\n", " ").Replace("\t", " ")).Replace(";;", "\r\n");
                    //}
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
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Lefik.xlsx");
            StatusStrip("Lefik");
        }
        private void GetTopopt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.topopt.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.topopt.ru",
                                "//a[contains(concat(' ', @class, ' '), ' product-title ')]",
                                "//a[contains(concat(' ', @class, ' '), ' pagenav ')][text()='В конец »»'] ", "?page=", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 3 == 0)
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
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' product-title ')]");
                    var artic = title;
                    var price = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @id, ' '), ' old-field ')]").Replace("руб.", "").Trim();
                    //desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/p|//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/div/p", "", new List<string>() { "Размер", "Цвет", "<!--" });
                    var page = new HtmlAgilityPack.HtmlDocument();
                    var descHtml =
                                    doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' product-text ')]");
                    if (descHtml != null)
                    {
                        page.LoadHtml(descHtml[0].InnerHtml);
                        desc = Helpers.GetItemsInnerText(page, "//p", "", new List<string>() { "Цвет", "<!--" });
                        size = Helpers.GetItemsInnerText(page, "//p", "Размер", null).Replace(":", "").Replace("ы", "").Trim().Replace(" ", "; ");
                    }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'zoom-id')]|//a[contains(concat(' ', @id, ' '), 'Zoomer')]");
                    //size = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' variant-size-radio ')]", "", null, "; ");
                    cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' breadcrumb_main ')]/a", "", new List<string>() { "Главная" }, "/");
                    col = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/p|//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/div/p", "Цвет", null).Replace(":", "").Trim();
                    if (size.Length == 0)
                    {
                        size = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/p|//div[contains(concat(' ', @class, ' '), ' product-text ')]/div/div/p", "Размер", null).Replace(":", "").Replace("ы", "").Trim();
                        if (!size.Contains(";"))
                        {
                            size = size.Replace(" ", "; ");
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
                        Photos = phs,
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Topopt.xlsx");
            StatusStrip("Topopt");
        }
        private void GetDonnasara(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://donnasara.com.ua/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://donnasara.com.ua",
                                "//a[contains(concat(' ', @class, ' '), ' b-product-line__product-name-link ')]",
                                "//div[contains(concat(' ', @class, ' '), ' b-pager ')]/a ", "/page_", null, 1);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 3 == 0)
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
                    var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' b-product__name ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' b-product__sku ')]").Replace("Код:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' b-data-list__text-name ')]").Replace("грн.", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' b-content__body ')]/p | //table[contains(concat(' ', @class, ' '), ' b-product-info ')]/tr ", "", new List<string>() { "Основные" }, ";;");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'imagebox')]");

                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    if (desc.Length > 0)
                    {
                        desc = Helpers.ReplaceWhiteSpace(desc.Replace("\n", " ").Replace("\t", " ")).Replace(";;", "\r\n");
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
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Donnasara.xlsx");
            StatusStrip("Donnasara");
        }
        private void GetBusinka(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.bus-i-nka.com/user/login", new NameValueCollection() { { "email", "kamalij%40mail.ru" }, { "password", "mail" }, { "login", "%D0%92%D0%BE%D0%B9%D1%82%D0%B8" } });

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "?page=all", cook, "http://www.bus-i-nka.com/", "//h3/a", null);

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

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product ')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' articul ')]").Replace("Артикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' price ')]/span").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), ' features ')]/li | //div[contains(concat(' ', @id, ' '), ' catdescription ')]/p", "", null, ";;");
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'zoom')]");

                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    if (desc.Length > 0)
                    {
                        desc = Helpers.ReplaceWhiteSpace(desc.Replace("\n", " ").Replace("\t", " ")).Replace(";;", "\r\n");
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
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Businka.xlsx");
            StatusStrip("Businka");
        }
        private void GetOpttextil(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.opttextil.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://www.opttextil.ru/",
                                "//div[contains(concat(' ', @class, ' '), ' catList ')]/div/span[contains(concat(' ', @class, ' '), ' title ')]/a",
                                null);


                if (prod.Count == 0)
                    prod.Add(catalog.Url);
                else
                {
                    var temp = new List<string>();
                    foreach (var s in prod)
                    {
                        var prod2 = Helpers.GetProductLinks(s, cook, s + "/",
                        "//div[contains(concat(' ', @class, ' '), ' catList ')]/div/span[contains(concat(' ', @class, ' '), ' title ')]/a", null);
                        if (prod2.Any())
                        {
                            //bool pro
                            //foreach (var s2 in prod2)
                            //{
                            //    var prod3 = Helpers.GetProductLinks(s2, cook, s2,
                            //    "//div[contains(concat(' ', @class, ' '), ' catList ')]/div/span[contains(concat(' ', @class, ' '), ' title ')]/a", null);
                            //    if (prod3.Any())
                            //    {
                            //        temp.AddRange(prod3);
                            //    }
                            //    else
                            //    {
                            temp.AddRange(prod2);
                            //}
                            //}
                        }
                    }
                    if (temp.Any())
                        prod = new HashSet<string>(temp);
                }
                var temp2 = new List<string>();

                foreach (var p in prod)
                {
                    var page = Helpers.GetHtmlDocument(p, "http://www.opttextil.ru", null, cook);
                    var links = page.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' navigationLink ')]/ul/li[last()]/a");
                    if (links != null && links.Any())
                    {
                        var max = 0;
                        var l = HttpUtility.HtmlDecode(links[0].Attributes["href"].Value);
                        var num = Regex.Replace(l.Substring(l.LastIndexOf("_p"), l.Length - l.LastIndexOf("_p")), @"[^\d]", "");
                        var tr = Int32.TryParse(num, out max);

                        for (var i = 1; i <= max; i++)
                        {
                            temp2.Add(p + "/_p" + i);
                        }
                    }
                }
                if (temp2.Any())
                {
                    temp2.AddRange(prod);
                    prod = new HashSet<string>(temp2);
                }
                int countRequest = 0;
                foreach (var res in prod)
                {

                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    if (doc2 == null)
                        continue;
                    var cont = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' catList ')]");
                    if (cont != null)
                    {
                        foreach (var c in cont)
                        {
                            var page1 = new HtmlAgilityPack.HtmlDocument();
                            page1.LoadHtml(c.InnerHtml);
                            var artic = "";
                            var img = Helpers.GetPhoto(page1, "",
                                            "/div[contains(concat(' ', @class, ' '), ' lblock ')]/span[contains(concat(' ', @class, ' '), ' img ')]/img",
                                            "", "http://www.opttextil.ru");
                            var title = Helpers.GetItemInnerText(page1,
                                            "/div[contains(concat(' ', @class, ' '), ' lblock ')]/span[contains(concat(' ', @class, ' '), ' title ')]");
                            var desc1 = Helpers.GetItemInnerText(page1,
                                            "/div[contains(concat(' ', @class, ' '), ' lblock ')][text()]");
                            var desc2 = Helpers.GetItemsInnerText(page1,
                                            "/div[contains(concat(' ', @class, ' '), ' lblock ')]/div[contains(concat(' ', @class, ' '), ' details ')]/div/table/tr/td/p | /div[contains(concat(' ', @class, ' '), ' lblock ')]/div[contains(concat(' ', @class, ' '), ' details ')]/div/table/tr/td/span",
                                            "", null, ";;");
                            if (desc1.Contains("<!--"))
                            {
                                var remove = desc1.Substring(desc1.IndexOf("<!--"),
                                                desc1.LastIndexOf("-->") + 3 - desc1.IndexOf("<!--"));
                                desc1 = desc1.Replace(remove, "").Trim();
                            }
                            if (desc2.Length > 4 && desc1.Contains(desc2.Substring(0, 4)))
                                desc1 = desc1.Remove(desc1.IndexOf(desc2.Substring(0, 4))).Trim();
                            if (desc1.Contains(title) && title.Length > 0)
                                desc1 = desc1.Replace(title, ";;").Trim();
                            var desc = (Helpers.ReplaceWhiteSpace(desc1) + ";;" + Helpers.ReplaceWhiteSpace(desc2)).Trim().Replace("\n", "").Replace("\t", "").Replace("\r", "");
                            //вырезать <!-- --> брать только 1 desc убрать пробелы
                            desc = desc.Replace(";;", "\r\n");
                            if (string.IsNullOrEmpty(artic))
                                artic = title;
                            var cont2 = page1.DocumentNode.SelectNodes("/div[contains(concat(' ', @class, ' '), ' itemOne ')]");
                            if (cont2 != null)
                            {
                                int z = 0;
                                foreach (var c2 in cont2)
                                {
                                    var page2 = new HtmlAgilityPack.HtmlDocument();
                                    page2.LoadHtml(c2.InnerHtml);
                                    var phs = Helpers.GetPhoto(page2,
                                                    "//a[contains(concat(' ', @rel, ' '), ' lightbox1 ')]",
                                                    "", "http://www.opttextil.ru");
                                    if (phs.Any())
                                        phs = phs.Where(x => !x.Contains("empty.png")).ToList();
                                    var prices = Helpers.GetItemsInnerTextList(page2,
                                                    "/div/div/div[contains(concat(' ', @class, ' '), ' price ')]/strong", "", null,
                                                    new List<string>() { "р.", " " });
                                    prices = prices.Select(x => x.Replace("р.", "").Replace(" ", "")).ToList();
                                    var sizes = Helpers.GetItemsInnerTextList(page2,
                                                    "/div/div/div[contains(concat(' ', @class, ' '), ' num ')]/div", "", null,
                                                    null);
                                    var col = Helpers.GetItemInnerText(page2,
                                                    "/div/span[contains(concat(' ', @class, ' '), ' color ')]");
                                    var cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' navstring ')]/a", "", new List<string>() { "Главная страница", " Оптовый интернет-магазин" }, "/");
                                    artic = title + "-" + col;
                                    var desc3 = Helpers.GetItemsInnerText(page2, "/div/div/table/tr/td/p", "", null);
                                    var price = "";
                                    var size = "";
                                    if (desc3.Length > 0 && z == 0)
                                        desc += "\r\n" + desc3;
                                    z++;
                                    if (prices.Any())
                                    {

                                        for (int i = 0; i < prices.Count(); i++)
                                        {
                                            price = prices[i];
                                            size = sizes[i];
                                            products.Add(new Product()
                                            {
                                                Url = z == 0 ? res : "",
                                                Article = artic,
                                                Color = col,
                                                Description = desc,
                                                Name = title,
                                                Price = price,
                                                CategoryPath = cat,
                                                Size = size,
                                                Photos = phs,
                                            });
                                            desc = "";
                                            title = "";
                                            phs = new List<string>();
                                            col = "";
                                            cat = "";
                                        }
                                    }
                                    else
                                    {
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
                                        });
                                    }
                                }
                            }
                            else
                            {

                            }
                        }
                    }


                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Opttextil.xlsx", false, false);
            StatusStrip("Opttextil");
        }
        private void GetStefanika(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.stefanika.ru/login.html", new NameValueCollection() { { "email", "usa@niksstore.ru" }, { "password", "1234567" } });

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "", "//div[contains(concat(' ', @class, ' '), ' good-title ')]/a", "//div[contains(concat(' ', @class, ' '), ' links ')]/a[last()]", "&page=", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 3 == 0)
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
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' center ')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/div[contains(concat(' ', @class, ' '), ' middle ')]/div[1]").Replace("Код товара:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' price ')]").Replace("руб.", "").Replace("Цена:", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div/font/div", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'thickbox')]");

                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    var prices = Helpers.GetItemsInnerTextList(doc2,
                                    "//label[contains(concat(' ', @for, ' '), 'option')]", "", null, null);
                    if (prices.Count > 1)
                    {
                        int i = 0;
                        foreach (var price1 in prices)
                        {
                            var t1 = price1.Substring(price1.LastIndexOf("-") + 1)
                                                            .Replace("р", "")
                                                            .Replace(" ", "")
                                                            .Trim();
                            var t2 = price1.Substring(0, price1.LastIndexOf("-")).Trim();
                            products.Add(new Product()
                            {
                                Url = i == 0 ? res : "",
                                Article = artic,
                                Color = col,
                                Description = desc,
                                Name = title,
                                Price = t1,
                                CategoryPath = cat,
                                Size = t2,
                                Photos = phs
                            });
                            if (i == 0)
                            {
                                title = "";
                                desc = "";
                                phs = new List<string>();
                                col = "";
                                cat = "";
                            }
                            i++;
                        }
                    }
                    else
                    {
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
                        });
                    }
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Stefanika.xlsx", false, false, false);
            StatusStrip("Stefanika");
        }
        private void GetPicoletto(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.picoletto.ru/index.php?route=account/login", new NameValueCollection() { { "email", "usa@niksstore.ru" }, { "password", "1234567" } });

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks2(catalog.Url, cook, "http://www.picoletto.ru", "//div[contains(concat(' ', @class, ' '), ' good-title ')]/a", "//div[contains(concat(' ', @class, ' '), ' links ')]/a[last()]", "&page=", null);

                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' center ')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/div[contains(concat(' ', @class, ' '), ' middle ')]/div[1]").Replace("Код товара:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//td/span[contains(concat(' ', @style, ' '), 'font-size:20px;')]").Replace("р", "").Replace(",", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div/font/div", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'thickbox')]");

                    cat = Helpers.GetEncodingCategory(catalog.Name);

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
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Picoletto.xlsx");
            StatusStrip("Picoletto");
        }
        private void GetWildberries()
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.wildberries.ru/1.2149.Lucky%20Child?sle=2426 ", new NameValueCollection());

            var prod = Helpers.GetProductLinks("http://www.wildberries.ru/1.2149.Lucky%20Child", cook, "",
                            "//a[contains(concat(' ', @class, ' '), ' ref_goods_n_p ')]", "//div[contains(concat(' ', @class, ' '), ' pageToInsert ')]/a", null, "http://www.wildberries.ru");
            int countRequest = 0;

            foreach (var res in prod)
            {
                //try
                //{
                if (countRequest % 3 == 0)
                {
                    Thread.Sleep(6000);
                }

                var doc2 = Helpers.GetHtmlDocument(res, "http://www.wildberries.ru/1.2149.Lucky%20Child", Encoding.GetEncoding("windows-1251"), cook);
                if (doc2 == null)
                    continue;
                var col = "";
                var size = "";
                var desc = "";
                var cat = "";
                var phs = new List<string>();
                var title = Helpers.GetItemInnerText(doc2, "//h3[contains(concat(' ', @itemprop, ' '), ' name ')]");
                var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' article ')]");
                col = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' color ')]");
                size = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' GoodCharacteristic ')]/label ", "", null, "; ");
                var price = Helpers.GetItemInnerText(doc2, "//ins[contains(concat(' ', @itemprop, ' '), ' price ')]").Replace("руб.", "").Replace(" ", "").Trim();
                desc = Helpers.GetItemsInnerText(doc2, "//p[contains(concat(' ', @itemprop, ' '), ' description ')] | //p[contains(concat(' ', @class, ' '), ' pp composition ')]", "", new List<string>(), ";;").Replace("\r\n", " ").Trim();
                var desc2 = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), ' pp-additional ')]/tr", "", new List<string>(), ";;").Replace("\r\n", " ").Trim();
                phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), ' enabledZoom ')]");
                if (desc2.Length > 0)
                {
                    desc = desc + " " + desc2;
                }
                if (desc.Length > 0)
                    desc = Helpers.ReplaceWhiteSpace(desc).Replace(";;", "\r\n");
                cat = "Lucky Child";
                if (phs.Any())
                {
                    phs = phs.Select(x => "http:" + x).ToList();
                }
                if (col.Length > 0)
                {
                    col = col.Replace(",", ";");
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
                });
                countRequest++;
                //}
                //catch (Exception exs) { }
            }


            Helpers.SaveToFile(products, path.Text + @"\WildberriesLuckyChild.xlsx");
            StatusStrip("WildberriesLuckyChild");
        }
        private void GetLiveToys()
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://live-toys.com", new NameValueCollection());

            var prod = Helpers.GetProductLinks("http://live-toys.com", cook, "http://live-toys.com",
                            "//a[contains(concat(' ', @class, ' '), ' prod ')]", null);
            int countRequest = 0;

            foreach (var res in prod)
            {
                //try
                //{
                //if (countRequest % 6 == 0)
                //{
                //    Thread.Sleep(6000);
                //}

                var doc2 = Helpers.GetHtmlDocument(res, "http://live-toys.com", null, cook);
                if (doc2 == null)
                    continue;
                var col = "";
                var size = "";
                var desc = "";
                var cat = "";
                var phs = new List<string>();
                var title = Helpers.GetItemInnerText(doc2, "//td/h1");
                var artic = title;
                var video = "";
                video = Helpers.GetItemsAttributt(doc2, "//param[contains(concat(' ', @name, ' '), ' data ')]", "", "value", null);
                desc += "Посмотрите наглядное видео с товаром: " + video + "\r\n";
                var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' phone ')]").Replace("руб.", "").Replace(" ", "").Trim();
                //var content = Helpers.GetItemsInnerText(doc2, "//ul/li/span/span | //td/p/span/span", "", new List<string>() { "Нажимайте на картинки справа", "ОБ УСЛОВИЯХ СОТРУДНИЧЕСТВА", "О ЗАКУПКЕ ТОВАРА НА САЙТЕ", "О ДОСТАВКЕ В РЕГИОНЫ", "ПРОЧИТАЙТЕ РАЗДЕЛ" });
                var content = Helpers.GetTextReplaceTags(doc2, "//ul/li/span/span | //td/p/span/span",
                                new List<string>()
                    {
                        "Нажимайте на картинки справа",
                        "ОБ УСЛОВИЯХ СОТРУДНИЧЕСТВА",
                        "О ЗАКУПКЕ ТОВАРА НА САЙТЕ",
                        "О ДОСТАВКЕ В РЕГИОНЫ",
                        "ПРОЧИТАЙТЕ РАЗДЕЛ"
                    });
                if (content.Length > 0)
                {
                    content = content.Replace("•", "").Replace("↑", "").Trim();
                    var temp = Regex.Split(content, "\r\n");
                    foreach (var s in temp)
                    {
                        //if (s.Contains("Размер"))
                        //    size = s.Replace("Размеры:", "").Replace("Размер:", "").Trim();
                        //else if (s.ToLower().Contains("цвета"))
                        //    col = s.ToLower().Replace("цвета", "").Replace(",", "; ").Trim();
                        //else 
                        if (!string.IsNullOrEmpty(s))
                            desc += s + "\r\n";
                    }
                    if (desc.Length > 0)
                        desc = desc.Trim();
                    if (size.Length > 0)
                        size = size.Substring(size.LastIndexOf(":") + 1).Replace(".", "").Trim();
                    if (col.Length > 0)
                        col = col.Substring(col.LastIndexOf(":") + 1).Replace(".", "").Trim();
                }

                phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), ' gallery-plants ')]", "", "http://live-toys.com");
                if (phs.Count > 0)
                {
                    phs = phs.Where(x => !x.Contains("no_photo")).ToList();
                }
                cat = "Каталог";
                var status = "";
                if (desc.Contains("в наличии"))
                {
                    if (desc.Contains("Есть в наличии."))
                    {
                        status = "";
                        desc = desc.Replace("Есть в наличии.", "");
                    }
                    else if (desc.Contains("ет в наличии."))
                    {
                        status = "no_sale";
                        desc = desc.Replace("Нет в наличии.", "").Replace("Этого товара временно нет в наличии.", "");
                    }
                    desc = desc.Replace("\r\n\r\n", "\r\n");
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
                    state = status
                });
                countRequest++;
                //}
                //catch (Exception exs) { }
            }


            Helpers.SaveToFile(products, path.Text + @"\LiveToys.xlsx");
            StatusStrip("LiveToys");
        }
        private void GetStilge(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.stilgi.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url + "?SHOWALL_1=1", cook, "http://www.stilgi.ru", "//span[contains(concat(' ', @class, ' '), ' name ')]/a", null);

                if (prod.Count == 0)
                    continue;

                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' art-article ')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' productPrice ')]");
                    var price = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), ' detail ')]/tr/td/div", "Мелкий опт:", null).Replace("руб.", "").Replace("Цена:", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), ' detail ')]/tr/td[contains(concat(' ', @colspan, ' '), ' 2 ')]", "", new List<string>() { "Размер" });
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]", "", "http://www.stilgi.ru");
                    cat = Helpers.GetItemsInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/a", "", new List<string>() { "Главная", "Каталог модной детской одежды" }, "/");
                    size = Helpers.GetTextReplaceTags(doc2, "//select[contains(concat(' ', @id, ' '), 'product')]",
                                    new List<string>() { "Выберите размер" }, "", "; ");
                    col = Helpers.GetTextReplaceTags(doc2, "//select[contains(concat(' ', @id, ' '), 'color')]",
                                    new List<string>() { "----" }, "", "; ");
                    if (price.Length == 0)
                    {
                        price = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), ' detail ')]/tr/td/div", "Цена:", null).Replace("руб.", "").Replace("Цена:", "").Trim();
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
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Stilge.xlsx");
            StatusStrip("Stilge");
        }
        private void GetVsspb(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://vsspb.com/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://vsspb.com", "//div[contains(concat(' ', @class, ' '), ' name ')]/a", null);

                if (prod.Count == 0)
                    continue;

                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' anonce ')]/h1");
                    var artic = title;
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' price_in ')]/span").Replace("руб.", "").Replace("Цена:", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' anonce ')]/p|//div[contains(concat(' ', @class, ' '), ' anonce ')]/div |//tr[contains(concat(' ', @height, ' '), '63')]/td/p", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'highslide')]", "//div[contains(concat(' ', @class, ' '), ' anonce ')]/p/img", "http://vsspb.com", "http://vsspb.com/");

                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    if (desc.Length == 0)
                    {
                        //var desc2 = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @class, ' '), 'anonce')]", null).Replace(title,"").Trim();
                    }
                    var desc2 = Helpers.GetTextReplaceTags(doc2, "//div[contains(concat(' ', @class, ' '), 'anonce')]", null, "", ";;").Replace(title, "").Replace("\r\n", "").Replace(";;", "\r\n").Trim();

                    products.Add(new Product()
                    {
                        Url = res,
                        Article = artic,
                        Color = col,
                        Description = desc2,
                        Name = title,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs,
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Vsspb.xlsx");
            StatusStrip("Vsspb");
        }
        private void GetIndialove()
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://indialove.my1.ru/", new NameValueCollection());

            var prod = Helpers.GetProductLinks("http://indialove.my1.ru/shop/all", cook, "http://indialove.my1.ru",
                            "//a[contains(concat(' ', @class, ' '), ' i-title ')]", "//a[contains(concat(' ', @class, ' '), ' pgSwch ')]", "/", null, 3);
            int countRequest = 0;
            var tt = new List<string>();
            //int countRequest = 0;
            foreach (var res in prod)
            {
                //try
                //{
                if (countRequest % 4 == 0)
                {
                    Thread.Sleep(6000);
                }

                var doc2 = Helpers.GetHtmlDocument(res, "http://indialove.my1.ru/shop/all", null, cook);
                if (doc2 == null)
                    continue;
                var col = "";
                var size = "";
                var desc = "";
                var cat = "";
                var phs = new List<string>();
                var title = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' ip-title ')]");
                var artic = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), ' shop-options ')]/li", "Артикул", null);
                if (string.IsNullOrEmpty(artic))
                    artic = title;
                else
                {
                    artic = artic.Substring(artic.LastIndexOf(":") + 1).Trim();
                }
                var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' ip-price ')]/span").Replace("руб.", "").Replace(" ", "").Trim();
                desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' ip-content ')]/div|//div[contains(concat(' ', @class, ' '), ' ip-content ')]/p", "", null);
                var desc2 = Helpers.GetItemsInnerText(doc2, "//ul[contains(concat(' ', @class, ' '), ' shop-options ')]/li", "", new List<string>() { "Артикул" });
                if (desc2.Length > 0)
                {
                    desc = desc2 + "\r\n" + desc;
                }
                //size = Helpers.GetItemsAttributt(doc2, "//select[contains(concat(' ', @class, ' '), ' inputboxattrib ')]/option", "есть", "value", null, "; ").Replace("_", " ").Trim().Replace(" ", "; ");

                phs = Helpers.GetPhoto(doc2, "", "//div[contains(concat(' ', @class, ' '), ' ip-photos ')]/img", "", "http://indialove.my1.ru");
                if (phs.Any())
                {
                    phs = phs.Select(x => x.Replace("m.", ".")).ToList();
                }
                cat = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' breadcrumbs ')]/a", "", new List<string>() { "Главная" }, "/");

                if (desc.Length > 0)
                {
                    var tr = Regex.Split(desc, "<!--<s");
                    desc = "";
                    foreach (var t in tr)
                    {
                        if (t.IndexOf("-->") > -1)
                        {
                            desc += t.Substring(t.IndexOf("-->") + 3).Replace("<!--</s>-->", "").Trim() + "\r\n";
                        }
                        else
                        {
                            desc += t + "\r\n";
                        }
                    }
                    desc = desc.Trim();
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
                });
                countRequest++;
                //}
                //catch (Exception exs) { }
            }


            Helpers.SaveToFile(products, path.Text + @"\Indialove.xlsx");
            StatusStrip("Indialove");
        }
        private void GetLioraShopt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://www.liora-shop.ru", new NameValueCollection());
            var list2 = new List<string>();
            foreach (var catalog in list)
            {
                var url = Helpers.GetEncodingCategory(catalog.Url);
                var prod = Helpers.GetProductLinks(url, cook, "http://www.liora-shop.ru", "//td/table/tr/td/a", "//ul[contains(concat(' ', @class, ' '), ' pagination ')]/li/a", null);
                //*[@id="vmMainPage"]/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/a
                if (prod.Count == 0)
                    continue;

                int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    if (countRequest % 3 == 0)
                    {
                        Thread.Sleep(6000);
                    }

                    var doc2 = Helpers.GetHtmlDocument(res, Helpers.GetEncodingCategory(catalog.Url), null, cook + "; _ym_visorc=w");
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//td/h1");
                    var artic = title;
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' productPrice ')]").Replace("руб.", "").Replace(" ", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//td[contains(concat(' ', @colspan, ' '), ' 4 ')]/p", "", new List<string>() { "Размер" });
                    //if (desc.Length < 5)
                    //{
                    //    desc = Helpers.GetItemsInnerText(doc2,
                    //        "//div[contains(concat(' ', @class, ' '), ' wsw_cat ')]/p", "",
                    //        new List<string>() {"Арт", "Цена", "руб"});
                    //}
                    size = Helpers.GetItemsAttributt(doc2, "//select[contains(concat(' ', @class, ' '), ' inputboxattrib ')]/option", "", "value", new List<string>() { "есть" }, "; ").Replace("_", " ").Trim();
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]");

                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    if (desc.Contains("Цвет"))
                    {
                        col =
                                        desc.Substring(desc.IndexOf("Цвет"), desc.Length - desc.IndexOf("Цвет"))
                                                        .Replace("Цвет", "")
                                                        .Replace(":", "")
                                                        .Trim();
                        desc = desc.Remove(desc.IndexOf("Цвет")).Trim();
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
                    });
                    countRequest++;
                    //}
                    //catch (Exception exs) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\LioraShopt.xlsx");
            StatusStrip("LioraShopt");
        }
        private void GetTexxit(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://texxit.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                //var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://texxit.ru", "//h3/a", null);
                var prLink = new List<Category>();
                try
                {
                    var client = new System.Net.WebClient();
                    client.Headers.Add(HttpRequestHeader.Cookie, cook);
                    client.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                    client.Headers.Add(HttpRequestHeader.Referer, "http://texxit.ru/");
                    client.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
                    client.Headers.Add(HttpRequestHeader.AcceptLanguage, "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                    var data = client.OpenRead(catalog.Url);
                    var reader = new StreamReader(data, Encoding.UTF8);
                    string s = reader.ReadToEnd();
                    data.Close();
                    reader.Close();
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(s);
                    var a = doc.DocumentNode.SelectNodes("//h3/a");

                    if (a != null)
                    {
                        var par = a[0].ParentNode.ParentNode.SelectNodes("//h3/following-sibling::p[text() and not(text()='&nbsp;') and not(text()='\r\n&nbsp;')]");
                        int i = 0;
                        foreach (var p in a)
                        {
                            var temp = WebUtility.HtmlDecode(p.Attributes["href"].Value);
                            if (!temp.Contains("http://texxit.ru"))
                                temp = "http://texxit.ru" + temp;
                            prLink.Add(new Category() { Url = temp, Name = par[i].InnerText.Trim() });
                            i++;
                        }
                    }
                }
                catch (Exception ex) { }
                //var prod = new HashSet<string>(prLink);
                if (prLink.Count == 0)
                    continue;

                //int countRequest = 0;
                foreach (var res in prLink)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res.Url, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h1/span").Replace("Арт.", "");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' node ')]/table/tr/td/b").Replace("Артикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' price ')]").Replace("руб.", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//span/span/strong", "", new List<string>() { "Арт", "Цена", "руб" });
                    if (desc.Length < 5)
                    {
                        desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' wsw_cat ')]/p", "", new List<string>() { "Арт", "Цена", "руб" });
                    }
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @class, ' '), 'openImg')]", "//div[contains(concat(' ', @class, ' '), 'prod_photo')]/img", "http://texxit.ru", "http://texxit.ru");
                    if (string.IsNullOrEmpty(price))
                    {
                        price = Helpers.GetItemsInnerText(doc2, "//span/span/strong", "Цена со скидкой", null).Replace("руб.", "").Trim();
                    }
                    if (phs.Any())
                    {
                        phs = phs.Select(x => x.Replace(" ", "")).ToList();
                    }
                    cat = Helpers.GetItemsInnerText(doc2, "//h1/a", "", new List<string>() { "Каталог" }, "/");

                    products.Add(new Product()
                    {
                        Url = res.Url,
                        Article = artic,
                        Color = col,
                        Description = desc,
                        Name = res.Name,
                        Price = price,
                        CategoryPath = cat,
                        Size = size,
                        Photos = phs,
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\Texxit.xlsx");
            StatusStrip("Texxit");
        }
        private void GetAimicoKids(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://aimico-kids.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://aimico-kids.ru", "//span[contains(concat(' ', @class, ' '), ' field-content ')]/a", "//ul[contains(concat(' ', @class, ' '), ' pager ')]/li/a", null);

                if (prod.Count == 0)
                    continue;

                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' node ')]/h1");
                    var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' node ')]/table/tr/td/b").Replace("Артикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' priceonpage ')]").Replace("руб.", "").Replace(".", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-body ')]", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'gallery')]", "", "http://aimico-kids.ru");
                    if (phs.Any())
                    {
                        phs =
                                        phs.Select(
                                                        x => x.Replace("http://aimico-kids.ruhttp://aimico-kids.ru", "http://aimico-kids.ru")).ToList();
                        if (phs.Count > 1)
                            phs.RemoveAt(0);
                    }
                    var labels = doc2.DocumentNode.SelectNodes("//label[contains(concat(' ', @for, ' '), 'edit-attributes')]");
                    var select = doc2.DocumentNode.SelectNodes("//select[contains(concat(' ', @name, ' '), 'attributes')]");
                    if (labels != null)
                    {
                        for (var i = 0; i < labels.Count(); i++)
                        {
                            var t = labels[i].InnerText;
                            if (t.ToLower().Contains("цвет"))
                            {
                                var s = select[i].InnerHtml;
                                var br = Regex.Split(s, ">");
                                foreach (var b in br)
                                {
                                    var temp = b;
                                    if (b.Contains("<"))
                                        temp = b.Remove(b.IndexOf("<"));
                                    if (!temp.Contains("Пожалуйста, выберите") && !string.IsNullOrEmpty(temp))
                                        col += temp + "; ";
                                }

                            }

                        }
                        if (col.Length > 0)
                            col = col.Remove(col.Length - 2);
                    }
                    var cats = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), 'breadcrumb')]/a", "", new List<string>() { "Главная", title }, "/");
                    cat = Helpers.GetEncodingCategory(catalog.Name).Trim();
                    if (!cats.Trim().Contains(cat) && !string.IsNullOrEmpty(cats.Trim()))
                        cat += "/" + cats;

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
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\AimicoKids.xlsx");
            StatusStrip("AimicoKids");
        }
        private void GetMiliePlatia()
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://xn--80ajimbcy0a0fp7a.xn--p1ai/", new NameValueCollection());

            var prod = Helpers.GetProductLinks("http://xn--80ajimbcy0a0fp7a.xn--p1ai/osen/", cook, "", "//div[contains(concat(' ', @class, ' '), ' subpagescontent ')]/p/p/a", null);
            //int countRequest = 0;
            var tt = new List<string>();
            foreach (var pr in prod)
            {
                tt.Add(pr.Replace("милыеплатья.рф", "xn--80ajimbcy0a0fp7a.xn--p1ai"));
            }
            prod = new HashSet<string>(tt);
            foreach (var res in prod)
            {
                //try
                //{
                //if (countRequest % 8 == 0)
                //{
                //    Thread.Sleep(5000);
                //}

                var doc2 = Helpers.GetHtmlDocument(res, "http://xn--80ajimbcy0a0fp7a.xn--p1ai/osen/", null/*Encoding.GetEncoding("windows-1251")*/, cook);
                if (doc2 == null)
                    continue;
                var col = "";
                var size = "";
                var desc = "";
                var cat = "";
                var price = "";
                var title = "";
                var phs = new List<string>();
                var content = Helpers.GetItemsInnerTextList(doc2, "//tr/td/div | //tr/td/p/span", "", null, null);
                var ccc = Helpers.GetTextReplaceTags(doc2, "//tr/td/div | //tr/td/p", new List<string>() { "*" });
                if (ccc.Length > 0)
                {
                    var ttt = Regex.Split(ccc, "\r\n");
                    title = ttt[0];
                    ttt[0] = "";
                    foreach (var s in ttt)
                    {
                        if (s.ToLower().Contains("цена"))
                        {
                            price = s.ToLower().Replace("цена", "").Replace("р.", "").Replace(":", "").Trim();
                        }
                        else if (s.ToLower().Contains("цвет") && !s.ToLower().Contains("аблица"))
                        {
                            col += s.ToLower().Replace("цвета:", "").Replace("цвет:", "").Trim() + ", ";
                        }
                        else if (s.ToLower().Contains("размеры"))
                        {
                            size += s.ToLower().Replace("размеры:", "").Replace("размеры", "").Trim() + ", ";
                        }
                        else if (!string.IsNullOrEmpty(s.Trim()) && !s.Contains("*") && !s.ToLower().Contains("аблица") && !s.ToLower().Contains("т."))
                        {
                            desc += s + ". ";
                        }
                    }
                    if (desc.Contains(".."))
                        desc = desc.Replace("..", ".").Replace("Т. АБЛИЦА ЦВЕТОВ.", "").Replace("Т.", "").Trim();
                    if (col.Length > 0)
                        col = col.Remove(col.Length - 2).Replace("ткань:", "").Replace(",", "; ").Trim();
                    if (size.Length > 0)
                    {
                        size = size.Remove(size.Length - 2).Replace(",", "; ").Trim();
                        if (size.Contains("ткань"))
                        {
                            size = size.Remove(size.IndexOf("ткань")).Trim();
                        }
                    }
                }

                var artic = "";
                if (string.IsNullOrEmpty(artic))
                    artic = title;
                phs = Helpers.GetPhoto(doc2, "", "//td/p/img|//td/img");
                if (string.IsNullOrEmpty(price))
                {
                    price = Helpers.GetItemInnerText(doc2, "//td/p/strong/span").ToLower().Replace("цена", "").Replace("р.", "").Replace(":", "").Trim();
                }
                cat = "Каталог";
                if (phs.Any())
                {
                    var temp = phs.Where(x => x.Contains("../")).ToList();
                    if (temp.Any())
                    {
                        foreach (var t in temp)
                        {
                            phs.Remove(t);
                        }
                        phs.AddRange(temp.Select(x => x.Replace("../", "http://xn--80ajimbcy0a0fp7a.xn--p1ai/")));
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
                    Photos = phs,
                });
                //countRequest++;
                //}
                //catch (Exception ex) { }
            }

            Helpers.SaveToFile(products, path.Text + @"\MiliePlatia.xlsx");
            StatusStrip("MiliePlatia");
        }
        private void GetEkbOpt(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://ekb-opt.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var tre = catalog.Url.Replace(".htm", "-pall.htm");
                var prod = Helpers.GetProductLinks(tre, cook, "http://ekb-opt.ru/", "//p[contains(concat(' ', @class, ' '), ' detal ')]/a", null);

                if (prod.Count == 0)
                    continue;

                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 8 == 0)
                    //{
                    //    Thread.Sleep(5000);
                    //}

                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, Encoding.GetEncoding("windows-1251"), cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//h2[contains(concat(' ', @class, ' '), ' h ')]");
                    var artic = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' articul ')]").Replace("Артикул:", "").Trim();
                    var price = Helpers.GetItemInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' price ')]").Replace("руб.", "").Replace("Цена:", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//p[contains(concat(' ', @class, ' '), ' descr ')] | //p[contains(concat(' ', @class, ' '), ' articul ')][2]", "", null);
                    if (string.IsNullOrEmpty(artic))
                        artic = title;
                    phs = Helpers.GetPhoto(doc2, "", "//div[contains(concat(' ', @class, ' '), 'image')]/img", "", "http://ekb-opt.ru");

                    cat = Helpers.GetItemInnerText(doc2, "//h1[contains(concat(' ', @class, ' '), ' razdel_title ')]").Replace("/" + title, "");

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
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\EkbOpt.xlsx");
            StatusStrip("EkbOpt");
        }
        private void GetLavira(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://lavira.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://lavira.ru",
                                "//div[contains(concat(' ', @class, ' '), ' pagination ')]/a", null);
                prod.Add(catalog.Url);
                if (prod.Count == 0)
                    continue;
                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 15 == 0)
                    //{
                    //    Thread.Sleep(7000);
                    //}
                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var cat = Helpers.GetEncodingCategory(catalog.Name);
                    var items = doc2.DocumentNode.SelectNodes("//table[contains(concat(' ', @class, ' '), ' product ')]");
                    if (items != null)
                    {
                        foreach (var item in items)
                        {
                            var doc = new HtmlAgilityPack.HtmlDocument();
                            doc.LoadHtml(HttpUtility.HtmlDecode(item.InnerHtml));
                            var content = Helpers.GetItemsInnerHtml(doc, "//div[contains(concat(' ', @class, ' '), ' description ')]/p", "", null, ";;");
                            var phs = Helpers.GetPhoto(doc, "//a[contains(concat(' ', @class, ' '), 'highslide')]");
                            if (content.Length > 0)
                            {
                                var cont = Regex.Split(content.Replace("<strong>", "").Replace("</strong>", ""), ";;");
                                foreach (var c in cont)
                                {
                                    var br = Regex.Split(c, "<br>");
                                    var price = "";
                                    var title = br[0];
                                    br[0] = string.Empty;
                                    if (!br[1].Contains("размеры") && !br[1].Contains("цвет") && br[1].Contains("\""))
                                    {
                                        title += " " + br[1];
                                        br[1] = string.Empty;
                                    }
                                    var artic = title;
                                    var size = "";
                                    var desc = "";
                                    var col = "";
                                    foreach (var b in br)
                                    {
                                        if (b.ToLower().Contains("размеры"))
                                            size = b.ToLower().Replace("размеры", "").Replace(":", "").Trim();
                                        else if (b.ToLower().Contains("цена"))
                                            price = b.ToLower().Replace("цена:", "").Replace("р.", "").Replace("р", "").Trim();
                                        else if (b.ToLower().Contains("цвет"))
                                            col = b.ToLower().Replace("цвет", "").Replace(":", "").Trim();
                                        else if (b.ToLower().Contains("цены"))
                                        {

                                        }
                                        else if (b.ToLower().Contains("р") && b.Contains("-"))
                                        {
                                            size += b.Substring(0, 5) + "; ";
                                            price += b.Replace("р", "").Replace(b.Substring(0, 5), "").Replace("р", "").Replace("-", "").Trim() + "; ";
                                        }
                                        else if (b.Contains("-"))
                                        {
                                            size += b + "; ";
                                        }
                                        else if (!string.IsNullOrEmpty(b))
                                            desc += b.Trim();
                                    }
                                    if (desc.Length == 0 && title.Length > (title.LastIndexOf("\"") + 1))
                                    {
                                        desc = title.Substring(title.LastIndexOf("\"") + 1).Trim();
                                        if (desc.Length > 0)
                                            title = artic = title.Replace(desc, "").Trim();
                                    }
                                    if (!price.Contains(";"))
                                    {
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
                                        });
                                    }
                                    else
                                    {
                                        var ss = Regex.Split(size, "; ");
                                        var pp = Regex.Split(price, "; ");
                                        if (pp.Count() > 2)
                                        {
                                            for (int i = 0; i < ss.Count(); i++)
                                            {
                                                if (!string.IsNullOrEmpty(ss[i]) && !string.IsNullOrEmpty(pp[i]))
                                                {
                                                    products.Add(new Product()
                                                    {
                                                        Url = res,
                                                        Article = artic,
                                                        Color = col,
                                                        Description = desc,
                                                        Name = title,
                                                        Price = pp[i],
                                                        CategoryPath = cat,
                                                        Size = ss[i],
                                                        Photos = phs,
                                                    });
                                                }
                                            }
                                        }
                                        else
                                        {
                                            for (int i = 0; i < ss.Count(); i++)
                                            {
                                                if (!string.IsNullOrEmpty(ss[i]) && !string.IsNullOrEmpty(pp[i]))
                                                {
                                                    products.Add(new Product()
                                                    {
                                                        Url = res,
                                                        Article = artic,
                                                        Color = col,
                                                        Description = desc,
                                                        Name = title,
                                                        Price = price,
                                                        CategoryPath = cat,
                                                        Size = ss[i],
                                                        Photos = phs,
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }

            Helpers.SaveToFile(products, path.Text + @"\Lavira.xlsx", false, false);
            StatusStrip("Lavira");

        }
        private void GetUbkiValentina(IEnumerable<Category> list)
        {
            var products = new List<Product>();
            var cook = Helpers.GetCookiePost("http://ubki-valentina.ru/", new NameValueCollection());

            foreach (var catalog in list)
            {
                var prod = Helpers.GetProductLinks(catalog.Url, cook, "http://ubki-valentina.ru",
                                "//div[contains(concat(' ', @class, ' '), ' field-content ')]/a", null);

                if (prod.Count == 0)
                    continue;
                //int countRequest = 0;
                foreach (var res in prod)
                {
                    //try
                    //{
                    //if (countRequest % 15 == 0)
                    //{
                    //    Thread.Sleep(7000);
                    //}
                    var doc2 = Helpers.GetHtmlDocument(res, catalog.Url, null, cook);
                    if (doc2 == null)
                        continue;
                    var col = "";
                    var size = "";
                    var desc = "";
                    var cat = "";
                    var phs = new List<string>();
                    var title = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @id, ' '), ' content ')]/h1");
                    var artic = title;
                    var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @class, ' '), ' uc-price ')]").Replace("р", "").Replace("p", "").Trim();
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' product-body ')]/p", "", new List<string>() { }).Trim();
                    phs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), 'main-product-image')]/a");
                    size = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' field-item odd ')]");
                    cat = Helpers.GetEncodingCategory(catalog.Name);
                    col = Helpers.GetItemsInnerText(doc2, "//table[contains(concat(' ', @class, ' '), ' table_product ')]/tbody/tr/td[1]", "", null, "; ").Trim();
                    if (size.Contains("-"))
                    {
                        var temp = Regex.Split(size, "-");
                        var result = "";
                        var look = (Int32.Parse(temp[1].Trim()) - Int32.Parse(temp[0].Trim())) / 2;
                        var begin = Int32.Parse(temp[0].Trim());
                        for (int i = 0; i <= look; i++)
                        {
                            result += begin + i * 2 + "; ";
                        }
                        size = result.Substring(0, result.Length - 2);
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
                    });
                    //countRequest++;
                    //}
                    //catch (Exception ex) { }
                }

            }
            Helpers.SaveToFile(products, path.Text + @"\UbkiValentina.xlsx");
            StatusStrip("UbkiValentina");

        }
        private void GetOtoys(IEnumerable<Category> list)
        {

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
                var products = new List<Product>();
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
                Helpers.SaveToFile(products, path.Text + @"\Otoys-" + catalog.Name + ".xlsx", photoCheck.Checked);
            }
            //Helpers.SaveToFile(products, path.Text + @"\Otoys.xlsx", photoCheck.Checked);
            StatusStrip("Otoys");

        }
        private void GetMaximum(IEnumerable<Category> list)
        {

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
                            var tr = Helpers.GetProductLinks(c, cook, "http://maximumufa.ru", "//td[2]/a", "//p[contains(concat(' ', @class, ' '), ' commonPager ')]/a[contains(text(), 'Вывести всё')]", null, "", "http://maximumufa.ru/represent/_represent_catalog/");
                            if (tr.Any())
                                temp.AddRange(tr.ToList());
                            else
                            {
                                var cat2 = Helpers.GetProductLinks(c, cook, "http://maximumufa.ru", "//section/ul/li/a", null, "http://maximumufa.ru/represent/_represent_catalog/");
                                if (cat.Any())
                                {
                                    foreach (var c2 in cat2)
                                    {
                                        var tr2 = Helpers.GetProductLinks(c2, cook, "http://maximumufa.ru",
                                                        "//td[contains(concat(' ', @class, ' '), 'second')]/a",
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
                var products = new List<Maximum>();
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
                    desc = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' info ')]/table/tr", "", new List<string>() { "Артикул" }, ";;");
                    var temp = Regex.Split(Helpers.ReplaceWhiteSpace(desc).Replace("\r\n", ""), ";;");
                    var made = "";
                    var upak = "";
                    var ufa = "";
                    var sterl = "";
                    desc = "";
                    foreach (var s in temp)
                    {
                        if (s.Contains("Производитель"))
                            made = s.Replace("Производитель", "").Trim();
                        else if (s.Contains("В упаковке"))
                            upak = s.Replace("В упаковке", "").Trim();
                        else if (s.Contains("Уфа"))
                            ufa = s.Replace("Уфа", "").Trim();
                        else if (s.Contains("Стерлитамак"))
                            sterl = s.Replace("Стерлитамак", "").Trim();
                    }
                    desc = string.Join("_", desc.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)).Replace("_", " ").Replace(";;", ", ").Replace("\r\n", "");
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

                    products.Add(new Maximum()
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
                        Photo = photo,
                        Ufa = ufa,
                        Sterlitamak = sterl,
                        Made = made,
                        Upakovka = upak
                    });
                    countRequest++;
                    //}
                    //catch (Exception ex) { }
                }
                Helpers.SaveToFile(products, path.Text + @"\Maximumufa-" + catalog.Name + ".xlsx", photoCheck.Checked);
            }
            //Helpers.SaveToFile(products, path.Text + @"\Maximum.xlsx", photoCheck.Checked);
            StatusStrip("Maximum");

        }
        private void GetJapanCosmetic(IEnumerable<Category> list)
        {
            var products = new List<JapanCosmetic>();
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
                    phs = Helpers.GetPhoto(doc2, "//a[contains(concat(' ', @rel, ' '), 'lightbox')]", "//span[contains(concat(' ', @itemprop, ' '), ' description ')]/p/img", "", "http://japan-cosmetic.biz");

                    cat = Helpers.GetItemsInnerText(doc2, "//a[contains(concat(' ', @class, ' '), ' pathway ')]", "", new List<string>() { "На главную" }, "/");
                    var attrList = Helpers.GetItemsAttributtList(doc2,
                                    "//select[contains(concat(' ', @class, ' '), 'inputboxattrib')][1]/option", "", "value", null, null);
                    var ton = "";
                    var aromat = "";
                    var option = "";
                    if (attrList.Count > 0)
                    {
                        var tt = Helpers.GetItemInnerText(doc2,
                                        "//div[contains(concat(' ', @class, ' '), ' vmCartAttributes ')]/div/label");
                        if (tt.Contains("Тон"))
                            ton = string.Join("; ", attrList);
                        else if (tt.Contains("Аромат"))
                            aromat = string.Join("; ", attrList);
                        else if (tt.Contains("Вариант"))
                            option = string.Join("; ", attrList);
                        else if (tt.Contains("Цвет"))
                            col = string.Join("; ", attrList);
                    }
                    products.Add(new JapanCosmetic()
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
                        Ton = ton,
                        Aroma = aromat,
                        Option = option
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
            if (photoCheck.Checked)
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
            }
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
                        var url = res;
                        for (int i = 0; i < arrPrice.Count(); i++)
                        {
                            if (i != 0)
                            {
                                title = desc = url = cat = photo = string.Empty;
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

            Helpers.SaveToFile(products, path.Text + @"\Noski.xlsx", photoCheck.Checked, false);
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
                        phs.AddRange(photos.Select(p => "http://iv-trikotage.ru" + p.Attributes["src"].Value.Replace("68_", "1200_")));
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
                        phs.AddRange(photos.Select(p => p.Attributes["src"].Value.Replace("_sm", "")));
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
            var products = new List<ProductPrices>();
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
                        var price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')]", 1).Replace("руб.", "").Replace("Базовая цена:", "").Replace(",", "").Replace(".", ",").Trim();
                        if (string.IsNullOrEmpty(price))
                            price = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')]").Replace("руб.", "").Replace("Базовая цена:", "").Replace(" ", "").Replace(",", "").Replace(".", ",").Trim();
                        var price2 = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')][1]").Replace("руб.", "").Replace("Роз. цена:", "").Replace(" ", "").Replace(",", "").Replace(".", ",").Trim();
                        var price3 = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' pricePr ')][3]").Replace("руб.", "").Replace("Мин. цена:", "").Replace(" ", "").Replace(",", "").Replace(".", ",").Trim();
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
                            desc = Helpers.GetItemInnerText(doc2, "//td[contains(concat(' ', @class, ' '), ' main ')]/table", 1);
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
                        products.Add(new ProductPrices()
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
                            Prices = new List<string>() { price2, price3 }
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
                    if (string.IsNullOrEmpty(price))
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
                            int i = 0;
                            foreach (var tr in table)
                            {
                                if (tr.InnerText.Trim() != string.Empty)
                                {
                                    var title2 = "";
                                    var desc3 = "";
                                    var page = new HtmlAgilityPack.HtmlDocument();
                                    page.LoadHtml(HttpUtility.HtmlDecode(tr.InnerHtml));
                                    var tit1 =
                                                    page.DocumentNode.SelectNodes(
                                                                    "//div[contains(concat(' ', @class, ' '), ' tovar-card__title ')]");
                                    if (tit1 != null)
                                    {
                                        title2 = tit1[0].InnerText.Trim();
                                        //if (title2.ToLower().Contains("размер"))
                                        size = title2.ToLower().Replace("размер", "").Trim();
                                        //desc += " "+tit1[0].InnerText.Trim();
                                    }
                                    //if (!string.IsNullOrEmpty(title.Trim()))
                                    //    title2 = title + " " + title2;
                                    //artic = title2;
                                    var des =
                                                    page.DocumentNode.SelectNodes(
                                                                    "//div[contains(concat(' ', @class, ' '), ' tovar-card__text1 ')]");
                                    //if (des != null)
                                    //	desc3 = des[0].InnerText.Trim();
                                    //if (!string.IsNullOrEmpty(desc.Trim()))
                                    //	desc3 = desc.Replace("\n", " "); //+ ". " + desc3;
                                    //desc3 =
                                    //		desc3.Replace("..", ".")
                                    //				.Replace("\n", "")
                                    //				.Replace("\t", "")
                                    //				.Replace("  ", "").Trim();
                                    var pr =
                                                    page.DocumentNode.SelectNodes(
                                                                    "//input[contains(concat(' ', @name, ' '), ' price ')]");
                                    if (pr != null)
                                        price = pr[0].Attributes["value"].Value.Replace(".", ",").Trim();
                                    products.Add(new Product()
                                    {
                                        Url = i == 0 ? res : "",
                                        Article = artic,
                                        Color = col,
                                        Description = desc,
                                        Name = title,
                                        Price = price,
                                        CategoryPath = cat,
                                        Size = size,
                                        Photos = phs
                                    });
                                    i++;
                                    if (i != 0)
                                    {
                                        title = "";
                                        cat = "";
                                        col = "";
                                        desc = "";
                                        phs = new List<string>();
                                    }
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
            Helpers.SaveToFile(products, path.Text + @"\Adel.xlsx", false, false, false);
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
                        prLink.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
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
                                    prLink.Add(WebUtility.HtmlDecode(p.Attributes["href"].Value));
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
            var products = new List<ProductPrices>();
            var cook = Helpers.GetCookiePost("http://xn----0tbbbddeld.xn--p1ai/login/?login=yes&backurl=%2F", new NameValueCollection() { { "AUTH_FORM", "Y" }, { "TYPE", "AUTH" }, { "backurl", "%2Flogin%2F%3Fbackurl%3D%252F" }, { "USER_LOGIN", "Shatlx" }, { "USER_PASSWORD", "5newinewi7" }, { "Login", "1" } });
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

                    var title = doc2.DocumentNode.SelectNodes("//h1[contains(concat(' ', @id, ' '), ' pagetitle ')]")[0].InnerText.Trim();
                    var artic = title;
                    var price = "";
                    var price1 = doc2.DocumentNode.SelectNodes("//span[contains(concat(' ', @class, ' '), ' catalog-item-price ')]");
                    if (price1 != null)
                        price = price1[0].InnerText.Replace("руб", "").Replace(" ", "").Trim();
                    var price3 = Helpers.GetItemsInnerTextList(doc2,
                                    "//span[contains(concat(' ', @class, ' '), ' catalog-item-price ')]", "", new List<string>() { price }, null);
                    if (price3.Any())
                        price3 = price3.Select(x => x.Replace("руб", "").Replace(" ", "").Trim()).ToList();
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
                    if (price3.Count > 4)
                        price3.RemoveAt(0);
                    cat = HttpUtility.HtmlDecode(cat);
                    products.Add(new ProductPrices()
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
                        Prices = price3
                    });

                    //}
                    //catch (Exception ex) { }
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
                                //if (s1.InnerText.Contains("Цвет:"))
                                //{
                                //    var t1 = s1.InnerText.Substring(s1.InnerText.IndexOf("Цвет:"), s1.InnerText.Length - s1.InnerText.IndexOf("Цвет:"));
                                //    col = t1.Replace("Цвет:", "").Trim();
                                //}
                                //else if (s1.InnerText.Contains("Размеры"))
                                //{
                                //    if (s1.InnerText.Contains("Размеры:"))
                                //        size = s1.InnerText.Replace("Размеры:", "").Trim();
                                //    else if (string.IsNullOrEmpty(size))
                                //    {
                                //        var begin = s1.InnerText.IndexOf(": ", s1.InnerText.IndexOf("Размеры")) + 1;
                                //        var t1 = s1.InnerText.Substring(begin + 1, s1.InnerText.Length - begin);
                                //        size = t1.Replace("Размеры:", "").Trim();
                                //    }
                                //}
                                //else 
                                if (!s1.InnerText.Contains("if") && !string.IsNullOrEmpty(s1.InnerText.Trim()) && !s1.InnerText.Contains("Название"))
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
                                        "", new List<string>() { "Главна" }, "/");
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
                                phs.Add("http://aventum.cz" + p.Attributes["src"].Value.Replace("_med", ""));
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
                        if (desc.Length > 1)
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
            Helpers.SaveToFile(products, path.Text + @"\Ozkan.xlsx", false, false, false);
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
                    var doc2 = Helpers.GetHtmlDocument(res,catalog,Encoding.GetEncoding("windows-1251"),cook);
                    
                    var title = Helpers.GetItemInnerText(doc2,"//dt/strong");
                    var d = doc2.DocumentNode.SelectNodes("//div[contains(concat(' ', @class, ' '), ' info-block additional ')]");

										var desc = Helpers.GetItemInnerText(doc2, "//dd/p");
										var price = Helpers.GetItemInnerText(doc2, "//span[contains(concat(' ', @id, ' '), ' share_price ')]").Replace("р.", "");
										var artic = Helpers.GetItemInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' info-block colors')]/div[contains(concat(' ', @class, ' '), ' right ')]").Replace("Артикул:", "");
										var col = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' info-block colors')]/div/ul/li","",null,"; ");
										var size = Helpers.GetItemsInnerText(doc2, "//div[contains(concat(' ', @class, ' '), ' info-block sizes')]/div/ul/li/b","",null,"; ");
                    var siz = "";
                    
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

										var phs = Helpers.GetPhoto(doc2, "//div[contains(concat(' ', @class, ' '), ' previews ')]/ul/li/a", "//img[contains(concat(' ', @id, ' '), ' main-image ')]", "http://tvoe.ru", "http://tvoe.ru");
                    
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
            //treeView1.SuspendLayout();
            //TreeNode tree = treeView1.Nodes.Cast<TreeNode>().FirstOrDefault(treeNode => treeNode.Text == nameTreeView);
            //if (tree != null)
            //{
            //    tree.ForeColor = Color.RoyalBlue;
            //    tree.Checked = false;
            //}
            //treeView1.ResumeLayout();
            var end = countStripStatus.Text.Substring(countStripStatus.Text.IndexOf("из"), countStripStatus.Text.Length - countStripStatus.Text.IndexOf("из"));
            var strip = countStripStatus.Text.Replace(end, "");
            var numb = Convert.ToInt32(Regex.Replace(strip, @"[^\d]", ""));
            this.Invoke(new Action(() => { countStripStatus.Text = "Скачено " + (numb + 1) + " " + end; }));
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
        public static bool FindDynamicElement(IWebDriver dr, By by, int timeOut)
        {
            try
            {
                var wait = new WebDriverWait(dr, new TimeSpan(timeOut * 1000));
                wait.Until(ExpectedConditions.ElementIsVisible(by));
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            foreach (var listBrowser in listBrowsers)
            {
                try
                {
                    listBrowser.Close();
                }
                catch (Exception ex) { }
            }
            listBrowsers.Clear();
            bwMain.CancelAsync();
            cancel.Cancel();
            timeStripStatus.Text = "Закачка отменена!!!";
            Start.Text = "Начать парсинг";
            Start.Enabled = true;
            Open.Enabled = false;
            Open.Visible = false;
            this.Update();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                foreach (var listBrowser in listBrowsers)
                {
                    try
                    {
                        listBrowser.Close();
                    }
                    catch (Exception ex) { }
                }
                listBrowsers.Clear();
                bwMain.CancelAsync();
                cancel.Cancel();
            }
            catch (Exception ex) { }
        }

        private void btnOpenSimaLand_Click(object sender, EventArgs e)
        {
            var fbd = new OpenFileDialog() { Filter = "txt files (*.txt)|*.txt", InitialDirectory = Environment.CurrentDirectory }; ;
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                pathSimaland.Text = fbd.FileName;
            }
        }
    }
}



