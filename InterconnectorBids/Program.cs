using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
 using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook; 
using System.Globalization;
namespace InterconnectorBids
{
    class Program
    {
        static List<string> exportBid = new List<string>();
        static List<string> hour = new List<string>();
        static List<string> halfHour = new List<string>();
        static List<string> importBid = new List<string>();

        static string[,] ImportBids = new string[50, 20];
        static string[,] ExportBids = new string[50, 20];

        static string date = null; 
        static string formattedDate;
        static int bidAmount; // From Cell C2 in excel sheet
        static string bidType = null; //From Cell B2 in excel sheet 
        static string resourceType = null;
        static string participantName = null; 

        static void Main(string[] args)
        {
            ReadFile();
            CreateXmlFile();
            createMail();             
            //Console.ReadKey();
        }

        private static void ReadFile()
        {
            string bookpath = @"F:\TRU-A\Team\Users\Shabab\Interconnector Bids\Bid Prices v2.xlsm";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook;

            //Open Excel workbook
            xlWorkbook = (Excel.Workbook)(xlApp.Workbooks.Open(bookpath, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing));

            //initialise date, amount and type for bids
            Excel.Worksheet excelSheet = (Excel.Worksheet)xlWorkbook.ActiveSheet;
            Excel.Range rngDate = (Excel.Range)excelSheet.Cells[2, 3];

            Excel.Range rngAmount = (Excel.Range)excelSheet.Cells[2, 2];
            bidAmount = int.Parse(rngAmount.Value2.ToString());

            Excel.Range rngType = (Excel.Range)excelSheet.Cells[3, 1];
            bidType = rngType.Value2.ToString();
            
            Excel.Range rngResType = (Excel.Range)excelSheet.Cells[5, 1];
            resourceType = rngResType.Value2.ToString();

            Excel.Range rngPartName = (Excel.Range)excelSheet.Cells[7, 1];
            participantName = rngPartName.Value2.ToString(); 

            date = rngDate.Value2.ToString();

            double test = double.Parse(date);
            DateTime finalDate = DateTime.FromOADate(test);
            formattedDate = finalDate.ToString("yyyy-MM-dd");

            //Iterate through each row- check Y/N
            for (int i = 0; i < 60; i++)
            {
                //Iterate and store the hour intervals from the spread sheets cells 
                Excel.Range rngHour = (Excel.Range)excelSheet.Cells[4 + i, 2];
                Excel.Range rngHalf = (Excel.Range)excelSheet.Cells[4 + i, 3];
                Excel.Range rngExport = (Excel.Range)excelSheet.Cells[4 + i, 4];
                Excel.Range rngImport;
    
                //Algorithm to calc cell position depending on amount of bids
                int timesBids = bidAmount * 2;
                rngImport = (Excel.Range)excelSheet.Cells[4 + i, 10 + timesBids - 2];                
                
                if (rngHour.Value2 == null)
                {
                    break;
                }
                else
                {
                    //Convert the excell value to string 
                    string getHour = rngHour.Value2.ToString();
                    string getHalf = rngHalf.Value2.ToString();
                    string getMaxExport = rngExport.Value2.ToString();
                    string getMaxImport = rngImport.Value2.ToString();

                    hour.Add(getHour);
                    halfHour.Add(getHalf);
                    exportBid.Add(getMaxExport);
                    importBid.Add(getMaxImport);

                    Console.WriteLine("{0}, {1}, {2}, {3} ", getHour, getHalf, getMaxExport, getMaxImport);
                }

                //Iterate each column 
                for (int x = 0; x < bidAmount * 2; x++)
                {
                    //Export Bids------------------- store each value into an array position 
                    Excel.Range exportBidCols = (Excel.Range)excelSheet.Cells[4 + i, 4 + x];
                    double convertBid2 = double.Parse(exportBidCols.Value2.ToString());
                    convertBid2 = Math.Round(convertBid2, 2);
                    string finalVal2 = convertBid2.ToString();
                    ExportBids[i, x] = finalVal2;

                    //Chose next 2 array positions to populate with extra export bid
                    if (x == bidAmount * 2 - 1)
                    {
                        int thirdPrice = bidAmount * 2;
                        int thirdMw = thirdPrice + 1;
                        ExportBids[i, thirdPrice] = "0";
                        double exportAddBid = convertBid2 + 0.01;
                        string finalExport = exportAddBid.ToString();
                        ExportBids[i, thirdMw] = finalExport;
                        ImportBids[i, thirdPrice] = "0";
                        ImportBids[i, thirdMw] = "0";
                    }

                    //Import Bids-----
                    Excel.Range importBidCols = (Excel.Range)excelSheet.Cells[4 + i, 10 + x];
                    double convertBid = double.Parse(importBidCols.Value2.ToString());
                    convertBid = Math.Round(convertBid, 2);
                    string finalVal = convertBid.ToString();
                    ImportBids[i, x] = finalVal;
                }
            }
            xlWorkbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            xlApp.Quit();
        }

        private static void CreateXmlFile()
        {
            int countPos1 = (bidAmount * 2) / 2 - 1;
            int countPos2 = bidAmount - 1;
            XmlDocument xmldox = new XmlDocument();
            XmlDeclaration decl = xmldox.CreateXmlDeclaration("1.0", "UTF-8", "");

            xmldox.InsertBefore(decl, xmldox.DocumentElement);

            //create root node
            XmlElement RootNode = xmldox.CreateElement("bids_offers");
            RootNode.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            RootNode.SetAttribute("noNamespaceSchemaLocation", "http://www.w3.org/2001/XMLSchema-instance", "mint_sem.xsd");

            //Create market submission nodes with user information
            XmlElement marketSub = xmldox.CreateElement("market_submit");
            marketSub.SetAttribute("application_type", "DAM");
            marketSub.SetAttribute("gate_window", bidType);
            marketSub.SetAttribute("trading_date", formattedDate);
            marketSub.SetAttribute("participant_name", participantName);
            marketSub.SetAttribute("user_name", "TNEWHAM");
            marketSub.SetAttribute("mode", "NORMAL");
            RootNode.AppendChild(marketSub);

            //Create offer information nodes
            XmlElement semOffer = xmldox.CreateElement("sem_interconnector_offer");
            semOffer.SetAttribute("resource_name", resourceType);
            semOffer.SetAttribute("resource_type", "INTERCONNECTOR");
            semOffer.SetAttribute("version_no", "1.0");
            marketSub.AppendChild(semOffer);

            //Dynamically create attributes for offers
            for (int i = 0; i < hour.Count; i++)
            {
                XmlElement interOffer = xmldox.CreateElement("interconnector_capacity");
                interOffer.SetAttribute("start_hr", hour[i]);
                interOffer.SetAttribute("start_int", halfHour[i]);
                interOffer.SetAttribute("end_hr", hour[i]);
                interOffer.SetAttribute("end_int", halfHour[i]);
                interOffer.SetAttribute("maximum_import_capacity_mw", importBid[i]);
                interOffer.SetAttribute("maximum_export_capacity_mw", exportBid[i]);
                semOffer.AppendChild(interOffer);
            }

            //Dynamically input price curve 
            for (int z = 0; z < hour.Count; z++)
            {
                XmlElement pqCurve = xmldox.CreateElement("pq_curve");
                pqCurve.SetAttribute("start_hr", hour[z]);
                pqCurve.SetAttribute("start_int", halfHour[z]);
                pqCurve.SetAttribute("end_hr", hour[z]);
                pqCurve.SetAttribute("end_int", halfHour[z]);
                semOffer.AppendChild(pqCurve);

                for (int y = 0; y < bidAmount; y++)
                {
                    if (!ImportBids[z, y + y].Equals("0"))
                    {           
                         //Import bids populated
                         XmlElement point = xmldox.CreateElement("point");
                         point.SetAttribute("price", ImportBids[z, y + y + 1]);
                         point.SetAttribute("quantity", ImportBids[z, y + y]);
                         pqCurve.AppendChild(point);                       
                    }
                    else //Import export link
                    {
                        if (bidAmount >= 2) 
                        {
                            if (y == countPos1 || y == countPos2)
                            {
                                if (ExportBids[z, y + y + 1] == "0" && ExportBids[z, y + y] == "0")
                                {
                                    //Do nothing   
                                }
                                else
                                {
                                    //Export bids populated
                                    XmlElement point = xmldox.CreateElement("point");
                                    point.SetAttribute("price", ExportBids[z, y + y + 1]);
                                    point.SetAttribute("quantity", ExportBids[z, y + y]);
                                    pqCurve.AppendChild(point);
                                }
                            }
                            else
                            {
                                //Export bids populated
                                XmlElement point = xmldox.CreateElement("point");
                                point.SetAttribute("price", ExportBids[z, y + y + 1]);
                                point.SetAttribute("quantity", ExportBids[z, y + y]);
                                pqCurve.AppendChild(point);
                            }
                        }                   
                        else
                        {
                            //Export bids populated
                            XmlElement point = xmldox.CreateElement("point");
                            point.SetAttribute("price", ExportBids[z, y + y + 1]);
                            point.SetAttribute("quantity", ExportBids[z, y + y]);
                            pqCurve.AppendChild(point);
                        }
                    }
                }                
                
                //Add third bid for Export
                for (int k = 0; k < 1; k++)
                {
                    if (!ExportBids[z, k + k].Equals("0"))
                    {
                        //Bug
                        XmlElement point = xmldox.CreateElement("point");
                        point.SetAttribute("price", ExportBids[z, bidAmount * 2 + 1]);
                        point.SetAttribute("quantity", ExportBids[z, bidAmount * 2]);
                        pqCurve.AppendChild(point);
                    }
                }
            }
            //add root to document
            xmldox.AppendChild(RootNode);
            xmldox.Save(@"F:\TRU-A\Team\Users\Shabab\Interconnector Bids\XMl Bid" + " " + resourceType + " " + ".xml");
        }

        //attach file to email 
        private static void createMail()
        { 
            try
            {
                List<string> toMail = new List<string>();
                toMail.Add("Shabab.Mahmood@axpo.com");         
                 
                //CC mail list
                List<string> ccMail = new List<string>();
               //ccMail.Add("Shabab.Mahmood@axpo.com");
               ccMail.Add("Alexandre.MaFat@axpo.com");
               //ccMail.Add("Kamal.Khoury@axpo.com");
                
                //Press F6 to rebuild after any changes         
                //Create new outlook application1`    
                Outlook.Application oApp = new Outlook.Application();
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                //Text to appear in Email body
                oMsg.HTMLBody = "Hi <br><br>Please find the bid attached <br><br>" + "<br><br>Kind Regards,<Br>Shabab Mahmood";

                String myAttachment = "Attatchment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttatchType = (int)Outlook.OlAttachmentType.olByValue;

                try
                {
                    if (System.IO.File.Exists(@"F:\TRU-A\Team\Users\Shabab\Interconnector Bids\XMl Bid" + " " + resourceType + " " + ".xml"))
                    {                        
                        //Retrieve todays file and date -5, specify email subject
                        Outlook.Attachment oAttatch = oMsg.Attachments.Add(@"F:\TRU-A\Team\Users\Shabab\Interconnector Bids\XMl Bid" + " " + resourceType + " " + ".xml", iAttatchType, iPosition, myAttachment);
              
                        var today = DateTime.Now.AddDays(1);
                        var date = today.Date;
                        oMsg.Subject = "Bid " + " " + resourceType + " "  + date.ToString("dd/MM/yyyy");

                        Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;

                        //iterating through recipients 
                        foreach (string receivers in toMail)
                        {
                            Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(receivers);
                            oRecip.Type = (int)Outlook.OlMailRecipientType.olTo;
                            oRecip.Resolve();
                        }

                        //Iterate through CC recipients 
                        foreach (string cc in ccMail)
                        {
                            Outlook.Recipient oCC = (Outlook.Recipient)oRecips.Add(cc);
                            oCC.Type = (int)Outlook.OlMailRecipientType.olCC;
                            oCC.Resolve();
                        }
                        oMsg.Send();
                        oRecips = null;
                        oMsg = null;
                        oApp = null;
                    }
                    else
                    {
                        Console.WriteLine("File failed to send");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("File not found");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to send email");
            }   

        }
    }
}
