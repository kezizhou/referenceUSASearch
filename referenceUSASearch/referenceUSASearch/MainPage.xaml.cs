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
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace referenceUSASearch
{
	/// <summary>
	/// Interaction logic for MainPage.xaml
	/// </summary>
	public partial class MainPage : Page
	{

		readonly List<string> lstLastNames = new List<string>();
		readonly List<string> lstZipCodes = new List<string>();
		IWebDriver driver;
		WebDriverWait wait;
		DataTable dt = new DataTable();

		public MainPage()
		{
			InitializeComponent();
		}

		private void Page_Loaded(object sender, RoutedEventArgs e)
		{
			if (Properties.Settings.Default.LibraryID != null)
			{
				txtLibraryID.Text = Properties.Settings.Default.LibraryID;
			}

			if (Properties.Settings.Default.LibraryPin != null)
			{
				txtLibraryPin.Password = Properties.Settings.Default.LibraryPin;
			}
		}

		private void btnSubmit_Click(object sender, RoutedEventArgs e)
		{
			if (Validate() == false)
			{
				return;
			}

			Thread t = new Thread(() => ButtonThread());
			t.Start();
		}

		private void ButtonThread()
		{
			try
			{
				NativeMethods.SetThreadExecutionState(NativeMethods.ES_CONTINUOUS | NativeMethods.ES_SYSTEM_REQUIRED);

				dt = new DataTable();

				StartDriverLogin();

				// First search, fill in all search filters
				WriteLog("Reference USA loaded, setting search filters.\n");
				InitialSearchFilters();
				Wait(2);

				// Loop through last names
				for (int i = 0; i < lstLastNames.Count; i++)
				{
					WriteLog("Searching for: " + lstLastNames[i] + "\n");

					SearchLastName(lstLastNames[i]);

					// Checks if no results pop up window found
					try
					{
						// No results, close pop up
						Wait(2);
						IWebElement noResultsPopup = driver.FindElement(By.XPath("/html/body/div[5]/div[3]/a"));
						noResultsPopup.Click();
						WriteLog("0 records found.\n\n");

					}
					catch (NoSuchElementException)
					{
						if (NextPageResults() == false)
						{
							// If search fails, try again
							if (!SearchLastName(lstLastNames[i]))
							{
								// Error occured, redo last search
								i -= 1;
								SearchLastName(lstLastNames[i]);
							}
						}

						// Revise search
						IWebElement reviseSearch = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/ul[2]/li[1]/a")));
						reviseSearch.Click();
						try
						{
							wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='cs-ContactsPerHousehold']")));
						}
						catch (Exception)
						{
							driver.Navigate().Refresh();
						}
					}
				}

				ExportToCSV(dt);

				MessageBox.Show("Search complete.");

			}
			catch (Exception ex)
			{
				if (driver != null)
				{
					driver.Quit();
					ExportToCSV(dt);
					WriteLog("The following error occured while processing. However, the current data file up to this point has been saved to your destination.\n\n");
				}
				else
				{
					WriteLog("The following error occured while processing.\n\n");
				}

				WriteLog(ex.ToString());
				NativeMethods.SetThreadExecutionState(NativeMethods.ES_CONTINUOUS);
			}
		}

		private bool Validate()
		{
			bool blnValidated = false;

			if (txtLibraryID.Text != "" && txtLibraryPin.Password != "" && txtLastNameFile.Text != "" && txtZipCodeFile.Text != "")
			{
				blnValidated = true;

				if (chkSaveCred.IsChecked == true)
				{
					Properties.Settings.Default.LibraryID = txtLibraryID.Text;
					Properties.Settings.Default.LibraryPin = txtLibraryPin.Password;
				}
			}
			else
			{
				if (txtLibraryID.Text == "")
				{
					lblLibraryIDReq.Visibility = Visibility.Visible;
				}
				else
				{
					lblLibraryPinReq.Visibility = Visibility.Hidden;
				}

				if (txtLibraryPin.Password == "")
				{
					lblLibraryPinReq.Visibility = Visibility.Visible;
				}
				else
				{
					lblLibraryPinReq.Visibility = Visibility.Hidden;
				}

				if (txtLastNameFile.Text == "")
				{
					lblLastNamesReq.Visibility = Visibility.Visible;
				}
				else
				{
					lblLastNamesReq.Visibility = Visibility.Hidden;
				}

				if (txtZipCodeFile.Text == "")
				{
					lblZipCodesReq.Visibility = Visibility.Visible;
				}
				else
				{
					lblZipCodesReq.Visibility = Visibility.Hidden;
				}
			}

			return blnValidated;
		}

		private void WriteLog(string strString)
		{
			if (!Dispatcher.CheckAccess())
			{
				Dispatcher.Invoke(() =>
				{
					txtLog.Text += strString;
					txtLog.ScrollToEnd();
				});
			}
		}

		private bool SearchLastName(string strLastName)
		{
			// Last name
			IWebElement lastName;
			try
			{
				lastName = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='LastName']")));
			}
			catch (WebDriverTimeoutException)
			{
				// Name checkbox was cleared
				try
				{
					IWebElement nameCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='cs-Name']")));
					if (!nameCheckbox.Selected)
					{
						nameCheckbox.Click();
					}
				}
				catch (WebDriverTimeoutException)
				{
					WriteLog("Error, restarting driver. (This sometimes occurs when requests are throttled.) Please double check if a last name was missed or duplicated. \n\n");

					// Try to restart driver
					driver.Quit();
					StartDriverLogin();

					// First search, fill in all search filters
					InitialSearchFilters();
					Wait(2);

					return false;
				}

				lastName = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='LastName']")));
			}
			lastName.Clear();
			lastName.SendKeys(strLastName);
			Wait(2);

			// One contact per household
			IWebElement oneContact;
			try
			{
				oneContact = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='primaryContacts']")));
			}
			catch (WebDriverTimeoutException)
			{
				// Contacts per household checkbox was cleared
				IWebElement contactsCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='cs-ContactsPerHousehold']")));
				oneContact = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='primaryContacts']")));
			}
			if (!oneContact.Selected)
			{
				oneContact.Click();
			}
			Wait(2);

			// Search button
			IWebElement search = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/div[3]/div/a[1]")));
			search.Click();

			return true;
		}

		private void btnUploadLastNames_Click(object sender, RoutedEventArgs e)
		{
			lstLastNames.Clear();

			OpenFileDialog openFile = new OpenFileDialog();
			openFile.Title = "Select file with last names";
			openFile.Filter = "Text files (*.txt)|*.txt";
			if (openFile.ShowDialog() == true)
			{
				string file = openFile.FileName;
				try
				{
					using (var reader = new StreamReader(file))
					{
						foreach (var line in File.ReadLines(file))
						{
							lstLastNames.Add(line.Trim());
						}
					}

					txtLastNameFile.Text = openFile.FileName;
				}
				catch (Exception ex)
				{
					WriteLog("The following error occured while processing your request. Please try again.\n\n");
					WriteLog(ex.ToString() + "\n\n");
				}
			}
		}

		private void btnUploadZipCodes_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFile = new OpenFileDialog
			{
				Title = "Select file with zip codes",
				Filter = "Text files (*.txt)|*.txt"
			};

			if (openFile.ShowDialog() == true)
			{
				string file = openFile.FileName;
				try
				{
					using (var reader = new StreamReader(file))
					{
						foreach (var line in File.ReadLines(file))
						{
							lstZipCodes.Add(line.Trim());
						}
					}

					txtZipCodeFile.Text = openFile.FileName;
				}
				catch (Exception ex)
				{
					WriteLog("The following error occured while processing your request. Please try again.\n\n");
					WriteLog(ex.ToString() + "\n\n");
				}
			}
		}

		private void Wait(int intSeconds)
		{
			Thread.Sleep(intSeconds * 1000);
		}

		private void ExportToCSV(DataTable dt)
		{
			string strFileName = "referenceUSASearch" + DateTime.Now.ToString("yyyy'-'MM'-'dd'-'HH'_'mm'_'ss") + ".csv";
			string strFilePath = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), strFileName);
			using (FileStream fs = File.Create(strFilePath)) { }

			StreamWriter sw = new StreamWriter(strFilePath, false);
			sw.Write(sw.NewLine);
			foreach (DataRow dr in dt.Rows)
			{
				for (int i = 0; i < dt.Columns.Count; i++)
				{
					if (!Convert.IsDBNull(dr[i]))
					{
						string value = dr[i].ToString();
						if (value.Contains(','))
						{
							value = string.Format("\"{0}\"", value);
							sw.Write(value);
						}
						else
						{
							sw.Write(dr[i].ToString());
						}
					}
					if (i < dt.Columns.Count - 1)
					{
						sw.Write(",");
					}
				}
				sw.Write(sw.NewLine);
			}
			sw.Close();

			driver.Quit();

			RemoveQuotes(strFilePath);
		}

		private void RemoveQuotes(string strFilePath)
		{
			string strFileContent = File.ReadAllText(strFilePath);
			strFileContent = strFileContent.Replace("\"", "");
			File.WriteAllText(strFilePath, strFileContent);
		}

		private void StartDriverLogin()
		{
			WriteLog("Starting driver...");

			ChromeOptions options = new ChromeOptions();

			// Do not show chrome window
			options.AddArgument("headless");

			// Hide logging window
			ChromeDriverService service = ChromeDriverService.CreateDefaultService("Resources");
			service.SuppressInitialDiagnosticInformation = true;
			service.HideCommandPromptWindow = true;

			driver = new ChromeDriver(service, options);

			// Data table
			DataColumnCollection columns = dt.Columns;
			if (!columns.Contains("First Name"))
			{
				dt.Columns.Add("First Name");
				dt.Columns.Add("Last Name");
				dt.Columns.Add("Street Address");
				dt.Columns.Add("City State");
				dt.Columns.Add("Zip");
				dt.Columns.Add("Phone");
			}

			driver.Manage().Window.Maximize();
			driver.Url = "http://www.referenceusa.com.kentucky.idm.oclc.org/UsConsumer/Search/Custom/2af92a95094f418c8de1c121ffc54219";
			wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

			// Login
			WriteLog("Logging into library account...");
			IWebElement libraryCard = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("user")));
			libraryCard.SendKeys(Properties.Settings.Default.LibraryID);

			IWebElement pin = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("pass")));
			pin.SendKeys(Properties.Settings.Default.LibraryPin);

			IWebElement libraryLogin = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector("input[type='submit']")));
			libraryLogin.Click();

			// Switch to new tab
			driver.SwitchTo().Window(driver.WindowHandles.Last());
		}

		private void InitialSearchFilters()
		{
			driver.Navigate().Refresh();
			IWebElement clear = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.ClassName("action-clear-search")));
			clear.Click();
			Wait(2);

			// Name checkbox
			IWebElement nameCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("cs-Name")));
			if (!nameCheckbox.Selected)
			{
				nameCheckbox.Click();
			}
			Wait(2);

			// Zip code filter
			IWebElement zipCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("cs-ZipCode")));
			if (!zipCheckbox.Selected)
			{
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
			foreach (string strZip in lstZipCodes)
			{
				strZipBlock += strZip + Environment.NewLine;
			}
			IWebElement zipCodes = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("pastedZipCodes")));
			zipCodes.SendKeys(strZipBlock);
			Wait(2);

			// Contacts per household
			IWebElement contactsCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("cs-ContactsPerHousehold")));
			try
			{
				driver.FindElement(By.XPath("//*[@id='phContactsPerHousehold']/div/div[1]/div/div/div"));
			}
			catch (NoSuchElementException)
			{
				if (!contactsCheckbox.Selected)
				{
					contactsCheckbox.Click();
				}
			}
		}

		private bool NextPageResults()
		{
			// Check for next page
			string strResultCount = "";
			int intResultCount = 0;
			bool blnResult = false;
			while (blnResult == false)
			{
				try
				{
					strResultCount = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='dbSelector']/div/div[2]/div[1]/ul[1]/li[1]/span"))).GetAttribute("innerText");
				}
				catch (Exception)
				{
					driver.Navigate().Refresh();
					return false;
				}
				blnResult = int.TryParse(strResultCount, NumberStyles.AllowThousands, CultureInfo.CurrentCulture.NumberFormat, out intResultCount);
			}

			WriteLog(intResultCount + " records found.\n\n");

			// Round up to get number of pages
			int intPages = (intResultCount + 25 - 1) / 25;
			int count = 1;
			while (count <= intPages)
			{

				// Results found, add to datatable
				string strHtmlSource = "";
				try
				{
					strHtmlSource = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("tblResults"))).GetAttribute("outerHTML");
				}
				catch (WebDriverException)
				{
					driver.Navigate().Refresh();
					strHtmlSource = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("tblResults"))).GetAttribute("outerHTML");
				}

				HtmlDocument htmlDoc = new HtmlDocument();
				htmlDoc.LoadHtml(strHtmlSource);

				var nodes = htmlDoc.DocumentNode.SelectNodes("//*[@id='searchResultsPage']/tr");
				var rows = nodes.Select(tr => tr
								.Elements("td")
								.Select(td => td.InnerText.Trim()).Skip(1)
								.ToArray());
				foreach (var row in rows)
				{
					dt.Rows.Add(row);
				}

				count += 1;

				// Not on last page
				if (count <= intPages)
				{
					IWebElement nextPage = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='searchResults']/div[1]/div/div[1]/div[2]/div[3]")));
					nextPage.Click();
					Wait(1);
				}
			}

			return true;
		}
	}

	internal static class NativeMethods
	{
		// Import SetThreadExecutionState Win32 API and necessary flags
		[DllImport("kernel32.dll")]
		public static extern uint SetThreadExecutionState(uint esFlags);
		public const uint ES_CONTINUOUS = 0x80000000;
		public const uint ES_SYSTEM_REQUIRED = 0x00000001;
	}
}
