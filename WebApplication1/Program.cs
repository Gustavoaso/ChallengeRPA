using System;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // Carrega a planilha de dados
        using (var workbook = new XLWorkbook("challenge.xlsx"))
        {
            var planilha = workbook.Worksheet(1);

            // Navega até a página inicial do formulário
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.rpachallenge.com/");

            // Clica no botão Start
            IWebElement botaoStart = driver.FindElement(By.CssSelector("button"));
            botaoStart.Click();

            // Itera sobre as linhas da planilha
            for (int row = 2; row <= planilha.LastRowUsed().RowNumber(); row++)
            {
                // Preenche os campos do formulário com os dados da planilha
                IWebElement campoNome = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelFirstName']"));
                campoNome.SendKeys(planilha.Cell(row, 1).GetString());

                IWebElement LastName = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelLastName']"));
                LastName.SendKeys(planilha.Cell(row, 2).GetString());

                IWebElement CompanyName = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelCompanyName']"));
                CompanyName.SendKeys(planilha.Cell(row, 3).GetString());

                IWebElement Role = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelRole']"));
                Role.SendKeys(planilha.Cell(row, 4).GetString());

                IWebElement Adress = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelAddress']"));
                Adress.SendKeys(planilha.Cell(row, 5).GetString());

                IWebElement campoEmail = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelEmail']"));
                campoEmail.SendKeys(planilha.Cell(row, 6).GetString());

                IWebElement campoTelefone = driver.FindElement(By.CssSelector("input[ng-reflect-name= 'labelPhone']"));
                campoTelefone.SendKeys(planilha.Cell(row, 7).GetString());

         

                // Clica no botão "Enviar"
                IWebElement botaoEnviar = driver.FindElement(By.CssSelector("input[type='submit']"));
                botaoEnviar.Click();

            }

          
          
        }
    }
}
