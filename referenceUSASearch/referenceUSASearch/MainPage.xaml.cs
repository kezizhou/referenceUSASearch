using HtmlAgilityPack;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace referenceUSASearch {
	/// <summary>
	/// Interaction logic for MainPage.xaml
	/// </summary>
	public partial class MainPage : Page {

		List<string> lstLastNames = new List<string>();
		List<string> lstZipCodes = new List<string>();
		IWebDriver driver;
		WebDriverWait wait;
		DataTable dt = new DataTable();

		public MainPage() {
			InitializeComponent();
		}

		private void btnSubmit_Click(object sender, RoutedEventArgs e) {
			try {
				startDriverLogin();

				// Loop through last names
				for (int i = 0; i < lstLastNames.Count; i++) {
					// First search, fill in all search filters
					if (i == 0) {
						initialSearchFilters();
					}

					Wait(2);

					// Last name
					IWebElement lastName;
					try {
						lastName = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='LastName']")));
					} catch (WebDriverTimeoutException) {
						// Name checkbox was cleared
						IWebElement nameCheckbox = driver.FindElement(By.XPath("//*[@id='cs-Name']"));
						if (!nameCheckbox.Selected) {
							nameCheckbox.Click();
						}

						lastName = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='LastName']")));
					}
					lastName.Clear();
					lastName.SendKeys(lstLastNames[i]);
					Wait(2);

					// One contact per household
					IWebElement oneContact;
					try {
						oneContact = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='primaryContacts']")));
					} catch (WebDriverTimeoutException) {
						// Contacts per household checkbox was cleared
						IWebElement contactsCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='cs-ContactsPerHousehold']")));
						oneContact = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='primaryContacts']")));
					}
					if (!oneContact.Selected) {
						oneContact.Click();
					}
					Wait(2);

					// Search button
					IWebElement search = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/div[3]/div/a[1]")));
					search.Click();

					// Checks if no results pop up window found
					try {
						// No results, close pop up
						Wait(2);
						IWebElement noResultsPopup = driver.FindElement(By.XPath("/html/body/div[5]/div[3]/a"));
						noResultsPopup.Click();

					} catch (NoSuchElementException) {
						nextPageResults();

						// Revise search
						IWebElement reviseSearch = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/ul[2]/li[1]/a")));
						reviseSearch.Click();
						wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='cs-ContactsPerHousehold']")));
					}
				}

				exportToCSV(dt);

			} catch (Exception ex) {
				MessageBox.Show(ex.ToString());
			}
		}

		private void btnUploadLastNames_Click(object sender, RoutedEventArgs e) {
			OpenFileDialog openFile = new OpenFileDialog();
			openFile.Title = "Select file with last names";
			openFile.Filter = "Text files (*.txt)|*.txt";
			if (openFile.ShowDialog() == true) {
				string file = openFile.FileName;
				try {
					using (var reader = new StreamReader(file)) {
						foreach (var line in File.ReadLines(file)) {
							lstLastNames.Add(line.Trim());
						}
					}
				} catch (Exception ex) {
					MessageBox.Show(ex.ToString());
				}
			}
		}

		private void btnUploadZipCodes_Click(object sender, RoutedEventArgs e) {
			OpenFileDialog openFile = new OpenFileDialog();
			openFile.Title = "Select file with zip codes";
			openFile.Filter = "Text files (*.txt)|*.txt";
			if (openFile.ShowDialog() == true) {
				string file = openFile.FileName;
				try {
					using (var reader = new StreamReader(file)) {
						foreach (var line in File.ReadLines(file)) {
							lstZipCodes.Add(line.Trim());
						}
					}
				} catch (Exception ex) {
					MessageBox.Show(ex.ToString());
				}
			}
		}

		private void Wait(int intSeconds) {
			Thread.Sleep(intSeconds * 1000);
		}

		private void exportToCSV(DataTable dt) {
			string strFileName = "referenceUSASearch" + DateTime.Now.ToString("yyyy'-'MM'-'dd'-'HH'_'mm'_'ss") + ".csv";
			string strFilePath = System.IO.Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), strFileName);
			using (FileStream fs = File.Create(strFilePath)) { }

			StreamWriter sw = new StreamWriter(strFilePath, false);
			sw.Write(sw.NewLine);
			foreach (DataRow dr in dt.Rows) {
				for (int i = 0; i < dt.Columns.Count; i++) {
					if (!Convert.IsDBNull(dr[i])) {
						string value = dr[i].ToString();
						if (value.Contains(',')) {
							value = String.Format("\"{0}\"", value);
							sw.Write(value);
						} else {
							sw.Write(dr[i].ToString());
						}
					}
					if (i < dt.Columns.Count - 1) {
						sw.Write(",");
					}
				}
				sw.Write(sw.NewLine);
			}
			sw.Close();

			driver.Quit();

			removeQuotes(strFilePath);
		}

		private void removeQuotes(string strFilePath) {
			string strFileContent = File.ReadAllText(strFilePath);
			strFileContent = strFileContent.Replace("\"", "");
			File.WriteAllText(strFilePath, strFileContent);
		}

		private void startDriverLogin() {
			driver = new ChromeDriver("Resources");

			// Data table
			dt.Columns.Add("First Name");
			dt.Columns.Add("Last Name");
			dt.Columns.Add("Street Address");
			dt.Columns.Add("City State");
			dt.Columns.Add("Zip");
			dt.Columns.Add("Phone");

			driver.Manage().Window.Maximize();
			driver.Url = "http://www.referenceusa.com.ezproxy.kentonlibrary.org/UsConsumer/Search/Custom/2a1f8fddfb9d46308739d7fe382c5910";

			// Login
			IWebElement libraryCard = driver.FindElement(By.XPath("//*[@id='userName']"));
			libraryCard.SendKeys("23126001684261");

			IWebElement pin = driver.FindElement(By.XPath("//*[@id='password_form']"));
			pin.SendKeys("0102");

			IWebElement libraryLogin = driver.FindElement(By.XPath("/html/body/div/div[3]/form/div[3]/div[2]/button"));
			libraryLogin.Click();

			// Switch to new tab
			driver.SwitchTo().Window(driver.WindowHandles.Last());

			wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
		}

		private void initialSearchFilters() {
			driver.Navigate().Refresh();
			IWebElement clear = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/div[3]/div/a[3]")));
			clear.Click();
			Wait(2);

			// Name checkbox
			IWebElement nameCheckbox = driver.FindElement(By.XPath("//*[@id='cs-Name']"));
			if (!nameCheckbox.Selected) {
				nameCheckbox.Click();
			}
			Wait(2);

			// Zip code filter
			IWebElement zipCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='cs-ZipCode']")));
			if (!zipCheckbox.Selected) {
				zipCheckbox.Click();
			}
			Wait(2);

			// Textarea method
			IWebElement zipText = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='phZipCode']/div/div[2]/div/fieldset/ol/li[1]/div[2]/div/a")));
			zipText.Click();
			Wait(2);

			// Clear textares
			IWebElement zipClear = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='pastedZipCodesContainer']/a")));
			zipClear.Click();
			Wait(2);

			// Fill in zip codes
			string strZipBlock = "";
			foreach (string strZip in lstZipCodes) {
				strZipBlock += strZip + Environment.NewLine;
			}
			IWebElement zipCodes = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='pastedZipCodes']")));
			zipCodes.SendKeys(strZipBlock);
			Wait(2);

			// Contacts per household
			IWebElement contactsCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='cs-ContactsPerHousehold']")));
			try {
				driver.FindElement(By.XPath("//*[@id='phContactsPerHousehold']/div/div[1]/div/div/div"));
			} catch (NoSuchElementException) {
				if (!contactsCheckbox.Selected) {
					contactsCheckbox.Click();
				}
			}
		}

		private void nextPageResults() {
			// Check for next page
			string strResultCount = driver.FindElement(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/ul[1]/li[1]/span")).GetAttribute("innerText");
			int intResultCount = int.Parse(strResultCount, NumberStyles.AllowThousands);
			// Round up to get number of pages
			int intPages = (intResultCount + 25 - 1) / 25;
			int count = 1;
			while (count <= intPages) {

				// Results found, add to datatable
				string strHtmlSource = driver.FindElement(By.XPath("//*[@id='tblResults']")).GetAttribute("outerHTML");

				HtmlDocument htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(strHtmlSource);

				var nodes = htmlDoc.DocumentNode.SelectNodes("//*[@id='searchResultsPage']/tr");
				var rows = nodes.Select(tr => tr
								.Elements("td")
								.Select(td => td.InnerText.Trim()).Skip(1)
								.ToArray());
				foreach (var row in rows) {
					dt.Rows.Add(row);
				}

				count += 1;

				// Not on last page
				if (count <= intPages) {
					IWebElement nextPage = driver.FindElement(By.XPath("//*[@id='searchResults']/div[1]/div/div[1]/div[2]/div[3]"));
					nextPage.Click();
					wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='tblResults']")));
				}
			}
		}
	}
}
