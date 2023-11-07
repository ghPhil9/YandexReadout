using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;

namespace YandexReadout
{
    internal class Core
    {
        static void Main(string[] args) => new Core();

        private string[] Requests { get; set; }

        internal Core()
        {
            Console.Title = "YandexReadout v0.1 [https://t.me/CSharpHive]";

            try
            {
                // Чтение списка запросов
                ReadList();

                // Старт
                Parse();
            }
            catch (Exception e) { LogMsg("[Ошибка] " + e.Message); }
            finally
            {
                try { browser.Quit(); }
                catch { }

                LogMsg("Парсинг выдачи завершён. Нажмите любую кнопку...");
                Console.ReadKey();

                // Сохранение логов
                File.WriteAllLines($"{Environment.CurrentDirectory}\\LastLogs.txt", logs);
            }
        }

        private List<string> logs = new List<string>();
        private IWebDriver browser;
        private Excel excel;

        private void ReadList()
        {
            string path = Environment.CurrentDirectory + "\\SearchList.txt";

            if (!File.Exists(path)) throw new Exception("Файл с поисковыми запросами не найден");

            Requests = File.ReadAllLines(path, Encoding.UTF8);
            LogMsg($"Поисковых запросов: {Requests.Count()} шт.");
        }

        private void Parse()
        {
            string xTitle = "(//span[@class='OrganicTitleContentSpan organic__title'])";
            string xLink = "(//div[contains(@class,'Path Organic-Path')]/a)";
            string xNextPage = "//a[@class='VanillaReact Pager-Item Pager-Item_type_next']";

            // Создаём браузер и лист таблицы
            CreateDriver();
            Captcha();
            excel = new Excel();

            // Отказ от дефолтного браузера
            if (browser.FindElements(By.XPath("//main//button")).Count != 0)
            {
                browser.FindElement(By.XPath("//main//button")).Click();
                Thread.Sleep(TimeSpan.FromSeconds(5));
            }

            // Логика
            foreach (var request in Requests)
            {
                try
                {
                    int count = 3;
                    bool first = true;

                    // Ввод поискового запроса
                    NewReq(request);
                    Captcha();

                    // Парсинг
                    for (int p = 0; p < count; p++)
                    {
                        // Следующая страница
                        if (p != 0)
                        {
                            browser.FindElement(By.XPath(xNextPage)).Click();
                            Captcha();
                        }

                        if (browser.FindElements(By.XPath(xTitle)).Count == browser.FindElements(By.XPath(xLink)).Count)
                        {
                            for (int i = 1; i <= browser.FindElements(By.XPath(xTitle)).Count; i++)
                            {
                                string title = browser.FindElement(By.XPath($"{xTitle}[{i}]")).Text;
                                string link = browser.FindElement(By.XPath($"{xLink}[{i}]")).GetAttribute("href");

                                LogMsg($"{i}. {title}");
                                excel.NewLine(title, link, first ? request : null);
                                first = false;
                            }
                        }
                        else
                        {
                            count++;
                            LogMsg($"На {p + 1} странице несовпадение тайтлов к ссылкам. Переход к следующей странице");
                        }
                    }
                }
                catch (Exception e) { LogMsg("[Ошибка] " + e.Message); }
            }

            // Завершение
            excel.Save();
        }

        private void NewReq(string request)
        {
            string xInput = "//input[@class='HeaderDesktopForm-Input mini-suggest__input']";
            string xSearch = "//div[@class='mini-suggest__button-text']";

            LogMsg("[Запрос] " + request);
            browser.FindElement(By.XPath(xInput)).Clear();
            browser.FindElement(By.XPath(xInput)).SendKeys(request);
            browser.FindElement(By.XPath(xSearch)).Click();
        }

        private void CreateDriver()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            browser = new ChromeDriver(service);
            browser.Manage().Window.Maximize();
            browser.Navigate().GoToUrl("https://ya.ru/search");
        }

        private void Captcha()
        {
            if (browser.Title.Contains("Ой!"))
            {
                Console.Beep();
                Console.Beep();
                LogMsg("[Капча] Решите капчу и нажмите Enter...");
                Console.ReadLine();
            }
        }

        private void LogMsg(string msg)
        {
            msg = $"[{DateTime.Now}] {msg}";
            logs.Add(msg);
            Console.WriteLine(msg);
        }
    }
}
