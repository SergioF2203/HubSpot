using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace testHubSpot
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Contact> listOfContacts = GetRecentModifiedContactsByDate(new DateTime(2019, 10, 15));

            FillExcelDataSet(listOfContacts);
        }

        // Convert meеhod from DateTime to Unix Timestamp.
        static double ConvertToUnixStampDate(DateTime dateTime)
        {
            DateTime origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            TimeSpan diff = dateTime - origin;
            return Math.Floor(diff.TotalMilliseconds);
        }

        // Get method recently modifided contacts from HubSpot API.
        static List<Contact> GetRecentModifiedContactsByDate(DateTime dateTime)
        {
            double unixStampDate = ConvertToUnixStampDate(dateTime);
            List<OriginContact> listOfOriginContacts = new List<OriginContact>();
            string jsonData = "";
            List<long> listOfVids = new List<long>();
            List<Contact> listOfContacts = new List<Contact>();


            // Endpoint GET Recently mod contacts.
            string url = ConfigurationManager.AppSettings.Get("recentContactUrl");

            WebRequest webRequest = WebRequest.Create(url);
            WebResponse webResponse = webRequest.GetResponse();

            using (Stream stream = webResponse.GetResponseStream())
            {
                try
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        jsonData = reader.ReadToEnd();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

            }
            webResponse.Close();

            // First 'page' of contacts from API
            var parsedRequest = JsonConvert.DeserializeObject<RecentContacts>(jsonData);

            List<ContactElement> listOfRecentlyContacts = new List<ContactElement>();
            listOfRecentlyContacts = parsedRequest.Contacts.ToList();

            long vidoffset = parsedRequest.VidOffset;
            long timeoffset = parsedRequest.TimeOffset;

            // Chek 'has-more' value, if is it 'true' and 'time-Offset' > serach date will continue get data from endpoint
            if (parsedRequest.HasMore == true)
            {
                bool reapeatRequest = false;
                do
                {
                    url += "&vidOffset=" + vidoffset + "&timeOffset=" + timeoffset;

                    WebRequest nextRequest = WebRequest.Create(url);
                    WebResponse nextResponse = nextRequest.GetResponse();
                    using (Stream stream = nextResponse.GetResponseStream())
                    {
                        try
                        {
                            using (StreamReader reader = new StreamReader(stream))
                            {
                                jsonData = reader.ReadToEnd();
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                    }
                    nextResponse.Close();

                    parsedRequest = JsonConvert.DeserializeObject<RecentContacts>(jsonData);

                    if (parsedRequest.TimeOffset >= unixStampDate)
                    {
                        vidoffset = parsedRequest.VidOffset;
                        timeoffset = parsedRequest.TimeOffset;
                        reapeatRequest = true;
                    }

                    List<ContactElement> tempList = new List<ContactElement>();
                    tempList = parsedRequest.Contacts.ToList();

                    listOfRecentlyContacts.AddRange(tempList);

                } while (reapeatRequest);

            }

            // Select Contacts with more precision by Date.
            foreach (var data in listOfRecentlyContacts)
            {
                if (Convert.ToDouble(data.Properties.Lastmodifieddate.Value) >= unixStampDate)
                {
                    listOfVids.Add(data.Vid);
                }
            }

            // Get each Contact's detail by ID from API endpoint. 
            foreach (var data in listOfVids)
            {
                // Endpoint 'Contact by Id'
                string contactByIdUrl = ConfigurationManager.AppSettings.Get("contactByIdUrl");
                string secondPartOfUrl = ConfigurationManager.AppSettings.Get("secondPartOfContactByIdUrl");
                string partOfUrl = "";

                partOfUrl += data.ToString() + secondPartOfUrl;
                contactByIdUrl += partOfUrl;

                WebRequest contactRequest = WebRequest.Create(contactByIdUrl);
                WebResponse contactsResponse = contactRequest.GetResponse();

                using (Stream stream = contactsResponse.GetResponseStream())
                {
                    try
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            jsonData = reader.ReadToEnd();
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
                contactsResponse.Close();

                var newparsedRequest = JsonConvert.DeserializeObject<OriginContact>(jsonData);
                listOfOriginContacts.Add(newparsedRequest);
            }

            // Parsed Contact's data 
            foreach (var data in listOfOriginContacts)
            {
                Contact contact = new Contact();

                contact.Vid = data.Vid;
                contact.Lifecyclestage = data.Properties["lifecyclestage"].Value;
                if (data.Properties.ContainsKey("firstname"))
                {
                    contact.FirstName = data.Properties["firstname"].Value;
                }
                if (data.Properties.ContainsKey("lastname"))
                {
                    contact.SecondName = data.Properties["lastname"].Value;
                }
                if (data.AssociatedCompany != null)
                {
                    contact.Company_id = data.AssociatedCompany.CompanyId;
                    if (data.AssociatedCompany.Properties.ContainsKey("name"))
                        contact.CompanyName = data.AssociatedCompany.Properties["name"].Value;
                    if (data.AssociatedCompany.Properties.ContainsKey("city"))
                        contact.City = data.AssociatedCompany.Properties["city"].Value;
                    if (data.AssociatedCompany.Properties.ContainsKey("state"))
                        contact.State = data.AssociatedCompany.Properties["state"].Value;
                    if (data.AssociatedCompany.Properties.ContainsKey("website"))
                        contact.Website = data.AssociatedCompany.Properties["website"].Value;
                    if (data.AssociatedCompany.Properties.ContainsKey("phone"))
                        contact.Phone = data.AssociatedCompany.Properties["phone"].Value;
                    if (data.AssociatedCompany.Properties.ContainsKey("zip"))
                        contact.Zip = data.AssociatedCompany.Properties["zip"].Value;
                }

                listOfContacts.Add(contact);
            }

            return listOfContacts;
        }

        // Fill data Excel document method
        static void FillExcelDataSet(List<Contact> contacts)
        {
            Excel.Application applicationEx;
            Excel._Workbook workbook;
            Excel._Worksheet worksheet;
            Excel.Range range;

            try
            {
                applicationEx = new Excel.Application();
                applicationEx.Visible = true;

                workbook = (Excel._Workbook)(applicationEx.Workbooks.Add(Missing.Value));
                worksheet = (Excel._Worksheet)workbook.ActiveSheet;

                string[] headers = { "Vid", "First Name", "Last Name", "Lifecyclestage", "CompanyId", "Name", "Website", "City", "State", "Zip", "Phone" };

                range = worksheet.get_Range("A1", "K1");
                range.EntireRow.Font.Bold = true;

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1] = headers[i];
                }

                int row = 2;
                foreach (var data in contacts)
                {
                    worksheet.Cells[row, 1] = data.Vid;
                    worksheet.Cells[row, 2] = data.FirstName;
                    worksheet.Cells[row, 3] = data.SecondName;
                    worksheet.Cells[row, 4] = data.Lifecyclestage;
                    worksheet.Cells[row, 5] = data.Company_id;
                    worksheet.Cells[row, 6] = data.CompanyName;
                    worksheet.Cells[row, 7] = data.Website;
                    worksheet.Cells[row, 8] = data.City;
                    worksheet.Cells[row, 9] = data.State;
                    worksheet.Cells[row, 10] = data.Zip;
                    worksheet.Cells[row, 11] = data.Phone;

                    row++;
                }

                worksheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Columns.AutoFit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
