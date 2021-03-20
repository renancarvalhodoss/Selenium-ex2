using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using System.Threading;
using System.IO;

 namespace ex2
 {
      class Program
      {
          static void Main(string[] args)
          {
              //abrindo url
              IWebDriver driver = new ChromeDriver();
              driver.Navigate().GoToUrl("https://www.americanas.com.br/");


              // obtendo dados do html
              driver.FindElement(By.XPath("//input[@id='h_search-input']")).SendKeys("Samsung Galaxy S10+");
              driver.FindElement(By.XPath("//input[@id='h_search-input']")).SendKeys(Keys.Enter);
              driver.FindElement(By.XPath("//a[@to='/produto/134217344?pfm_carac=samsung-galaxy-s10&pfm_page=search&pfm_pos=grid&pfm_type=search_page']")).Click(); 
              string preco_sam = driver.FindElement(By.XPath("//div[@class='src__BestPrice-sc-1jvw02c-5 cBWOIB priceSales']")).Text;
              string aparelho_sam = driver.FindElement(By.XPath("//span[@class='src__Text-sc-154pg0p-0 src__Title-uexifx-0 dItrhU']")).Text;
              string cashback_sam = driver.FindElement(By.XPath("//span[@class='cashback__Green-j5qxid-0 fVeHjJ']")).Text;

              driver.FindElement(By.XPath("//input[@id='h_search-input']")).SendKeys("Xiaomi Redmi Note 9 Pro");
              driver.FindElement(By.XPath("//input[@id='h_search-input']")).SendKeys(Keys.Enter);
              driver.FindElement(By.XPath("//a[@to='/produto/2672512076?pfm_carac=xiaomi-redmi-note-9-pro&pfm_index=2&pfm_page=search&pfm_pos=grid&pfm_type=search_page']")).Click(); 
              string preco_xia = driver.FindElement(By.XPath("//div[@class='src__BestPrice-sc-1jvw02c-5 cBWOIB priceSales']")).Text;
              string aparelho_xia = driver.FindElement(By.XPath("//span[@class='src__Text-sc-154pg0p-0 src__Title-uexifx-0 dItrhU']")).Text;
              string cashback_xia = driver.FindElement(By.XPath("//span[@class='cashback__Green-j5qxid-0 fVeHjJ']")).Text;
              
              driver.FindElement(By.XPath("//input[@id='h_search-input']")).SendKeys("IPhone 11");
              driver.FindElement(By.XPath("//input[@id='h_search-input']")).SendKeys(Keys.Enter);
              Thread.Sleep(8000);
              driver.FindElement(By.XPath("//div[@class='src__Wrapper-sc-1k0ejj6-2 dGIFSc']/a[@to='/produto/338918556?pfm_carac=iphone-11&pfm_page=search&pfm_pos=grid&pfm_type=search_page']")).Click(); 
              string preco_iph = driver.FindElement(By.XPath("//div[@class='src__BestPrice-sc-1jvw02c-5 cBWOIB priceSales']")).Text;
              string aparelho_iph = driver.FindElement(By.XPath("//span[@class='src__Text-sc-154pg0p-0 src__Title-uexifx-0 dItrhU']")).Text;
              string cashback_iph = driver.FindElement(By.XPath("//body/div[@id='root']/div[1]/div[3]/div[2]/div[1]/div[3]/span[1]/p[1]/span[1]")).Text;
     
              //criando xlsx
              var package = new ExcelPackage();
              var workbook = package.Workbook;
              var planilha = workbook.Worksheets.Add("planilha");

              //criando tabela 
              planilha.Cells["A1:C1"].Style.Font.Bold = true;
              planilha.Cells["A1:C1"].Style.Font.Italic = true;
              planilha.Row(1).Height = 40;
              planilha.Column(1).Width = 90;
              planilha.Column(2).Width = 30;
              planilha.Column(3).Width = 30;
              planilha.Cells["1:4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
              planilha.Row(1).Style.VerticalAlignment =OfficeOpenXml.Style.ExcelVerticalAlignment.Center;              
              planilha.TabColor = System.Drawing.Color.Black;
              planilha.Cells["A1"].Value = "APARELHO";
              planilha.Cells["B1"].Value = "PEÇO";
              planilha.Cells["C1"].Value = "CASHBACK"; 
              
              //passando dados pra tabela
              planilha.Cells["A2"].Value = aparelho_sam;
              planilha.Cells["B2"].Value = preco_sam;
              planilha.Cells["C2"].Value = cashback_sam;
               
              planilha.Cells["A3"].Value = aparelho_xia;
              planilha.Cells["B3"].Value = preco_xia;
              planilha.Cells["C3"].Value = cashback_xia;

              planilha.Cells["A4"].Value = aparelho_iph;
              planilha.Cells["B4"].Value = preco_iph;
              planilha.Cells["C4"].Value = cashback_iph;                           

              //salvando planilha
              package.SaveAs(new FileInfo(@"cotação.xlsx"));
              
              driver.Quit();
                         
            }
        }
    }