using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BasicExcelFileOpenXML
{
    public class CustomersReport
    {
        public string Name { get; set; }
        public string RegisterDate { get; set; }
        public string LastBuy { get; set; }
        public string Item { get; set; }
        public int Quantity { get; set; }
        public double ItemCost { get; set; }

        public List<CustomersReport> GetRandomData()
        {
            List<CustomersReport> report = new List<CustomersReport>();
            report.Add(new CustomersReport { Name = "Jennifer W. Odom,", RegisterDate = "August 7th, 2018", LastBuy = "December 15th, 2017", Item = "Suspendisse sagittis.", Quantity = 2, ItemCost = 52.40 });
            report.Add(new CustomersReport { Name = "Jeanette K. Cleveland,", RegisterDate = "January 28th, 2018", LastBuy = "February 2nd, 2019", Item = "augue ac,", Quantity = 4, ItemCost = 26.07 });
            report.Add(new CustomersReport { Name = "Odette J. Hawkins,", RegisterDate = "March 5th, 2018", LastBuy = "January 22nd, 2018", Item = "tincidunt, n", Quantity = 4, ItemCost = 35.93 });
            report.Add(new CustomersReport { Name = "Adena A. Munoz,", RegisterDate = "December 28th, 2017", LastBuy = "July 27th, 2019", Item = "Vivamus molestie,", Quantity = 5, ItemCost = 35.03 });
            report.Add(new CustomersReport { Name = "Gil J. Mccarthy,", RegisterDate = "January 12th, 2018", LastBuy = "August 8th, 2019", Item = "nibh. D", Quantity = 6, ItemCost = 66.26 });
            report.Add(new CustomersReport { Name = "Devin E. Glover,", RegisterDate = "November 15th, 2017", LastBuy = "October 20th, 2018", Item = "Maecenas libero,", Quantity = 6, ItemCost = 61.76 });
            report.Add(new CustomersReport { Name = "Vera W. Lancaster,", RegisterDate = "September 17th, 2017", LastBuy = "June 15th, 2018", Item = "iaculis aliquet,", Quantity = 3, ItemCost = 24.65 });
            report.Add(new CustomersReport { Name = "Quail Y. Sanchez,", RegisterDate = "January 20th, 2018", LastBuy = "November 29th, 2017", Item = "erat semper,", Quantity = 2, ItemCost = 54.23 });
            report.Add(new CustomersReport { Name = "Tobias X. Burris,", RegisterDate = "February 26th, 2018", LastBuy = "March 6th, 2018", Item = "molestie orci,", Quantity = 0, ItemCost = 96.24 });
            report.Add(new CustomersReport { Name = "Dakota F. Saunders,", RegisterDate = "December 5th, 2017", LastBuy = "August 21st, 2018", Item = "mauris elit,", Quantity = 10, ItemCost = 36.82 });
            report.Add(new CustomersReport { Name = "Wylie Z. Weiss,", RegisterDate = "December 11th, 2017", LastBuy = "July 10th, 2018", Item = "adipiscing lacus.", Quantity = 5, ItemCost = 23.74 });
            report.Add(new CustomersReport { Name = "Nelle R. Price,", RegisterDate = "September 7th, 2017", LastBuy = "November 3rd, 2017", Item = "purus, i", Quantity = 10, ItemCost = 8.47 });
            report.Add(new CustomersReport { Name = "Kamal X. Jarvis,", RegisterDate = "January 11th, 2018", LastBuy = "February 24th, 2019", Item = "Donec sollicitudin,", Quantity = 2, ItemCost = 86.42 });
            report.Add(new CustomersReport { Name = "Christine Q. Johnson,", RegisterDate = "May 15th, 2018", LastBuy = "May 15th, 2019", Item = "Pellentesque habitant,", Quantity = 10, ItemCost = 67.81 });
            report.Add(new CustomersReport { Name = "Aiko E. Bass,", RegisterDate = "July 5th, 2018", LastBuy = "April 6th, 2018", Item = "et magnis,", Quantity = 2, ItemCost = 84.37 });
            report.Add(new CustomersReport { Name = "Rosalyn A. Decker,", RegisterDate = "September 17th, 2017", LastBuy = "August 31st, 2017", Item = "hendrerit consectetuer,", Quantity = 4, ItemCost = 93.52 });
            report.Add(new CustomersReport { Name = "Leilani I. Howell,", RegisterDate = "March 12th, 2018", LastBuy = "June 16th, 2019", Item = "Suspendisse ac,", Quantity = 5, ItemCost = 66.54 });
            report.Add(new CustomersReport { Name = "Echo K. Garza,", RegisterDate = "November 7th, 2017", LastBuy = "September 24th, 2018", Item = "mi, a", Quantity = 8, ItemCost = 95.27 });
            report.Add(new CustomersReport { Name = "Giselle E. Savage,", RegisterDate = "August 26th, 2017", LastBuy = "August 13th, 2019", Item = "sit amet,", Quantity = 5, ItemCost = 59.48 });
            report.Add(new CustomersReport { Name = "Amela A. Huff,", RegisterDate = "October 2nd, 2017", LastBuy = "May 23rd, 2018", Item = "Phasellus nulla.", Quantity = 10, ItemCost = 53.11 });
            report.Add(new CustomersReport { Name = "Jarrod J. Hull,", RegisterDate = "July 21st, 2018", LastBuy = "January 12th, 2018", Item = "eu, l", Quantity = 2, ItemCost = 8.57 });
            report.Add(new CustomersReport { Name = "Zorita C. Walter,", RegisterDate = "September 3rd, 2017", LastBuy = "February 4th, 2019", Item = "amet massa.", Quantity = 8, ItemCost = 31.42 });
            report.Add(new CustomersReport { Name = "Yoko O. Holman,", RegisterDate = "November 15th, 2017", LastBuy = "February 6th, 2019", Item = "tincidunt aliquam,", Quantity = 4, ItemCost = 17.92 });
            report.Add(new CustomersReport { Name = "Donovan S. Robles,", RegisterDate = "March 3rd, 2018", LastBuy = "April 7th, 2018", Item = "urna. N", Quantity = 2, ItemCost = 79.40 });
            report.Add(new CustomersReport { Name = "Ivor N. Daugherty,", RegisterDate = "February 28th, 2018", LastBuy = "September 15th, 2018", Item = "magnis dis,", Quantity = 5, ItemCost = 53.09 });
            report.Add(new CustomersReport { Name = "Dara Q. Hart,", RegisterDate = "September 20th, 2017", LastBuy = "July 10th, 2019", Item = "faucibus orci,", Quantity = 5, ItemCost = 10.94 });
            report.Add(new CustomersReport { Name = "Hall A. Strickland,", RegisterDate = "December 23rd, 2017", LastBuy = "September 29th, 2017", Item = "est. N", Quantity = 6, ItemCost = 25.31 });
            report.Add(new CustomersReport { Name = "Jakeem O. Williamson,", RegisterDate = "October 15th, 2017", LastBuy = "June 30th, 2018", Item = "eget varius,", Quantity = 5, ItemCost = 84.63 });
            report.Add(new CustomersReport { Name = "Aurelia K. Perez,", RegisterDate = "March 18th, 2018", LastBuy = "July 30th, 2019", Item = "Phasellus vitae,", Quantity = 5, ItemCost = 80.05 });
            report.Add(new CustomersReport { Name = "Shaeleigh Z. Bowen,", RegisterDate = "December 9th, 2017", LastBuy = "June 28th, 2019", Item = "ac nulla.", Quantity = 8, ItemCost = 88.51 });
            report.Add(new CustomersReport { Name = "Leonard G. Guy,", RegisterDate = "October 27th, 2017", LastBuy = "March 22nd, 2019", Item = "cubilia Curae;", Quantity = 6, ItemCost = 70.25 });
            report.Add(new CustomersReport { Name = "Cairo I. Gibbs,", RegisterDate = "October 8th, 2017", LastBuy = "December 21st, 2018", Item = "porttitor vulputate,", Quantity = 6, ItemCost = 67.43 });
            report.Add(new CustomersReport { Name = "Leah D. Sellers,", RegisterDate = "February 24th, 2018", LastBuy = "July 1st, 2019", Item = "Nunc lectus", Quantity = 4, ItemCost = 85.41 });
            report.Add(new CustomersReport { Name = "Rafael Q. Dawson,", RegisterDate = "February 19th, 2018", LastBuy = "May 7th, 2018", Item = "risus, at", Quantity = 1, ItemCost = 31.08 });
            report.Add(new CustomersReport { Name = "Aurelia A. Park,", RegisterDate = "October 9th, 2017", LastBuy = "August 31st, 2017", Item = "tincidunt pede,", Quantity = 5, ItemCost = 69.95 });
            report.Add(new CustomersReport { Name = "Matthew I. Baker,", RegisterDate = "January 17th, 2018", LastBuy = "June 30th, 2019", Item = "consectetuer adipiscing,", Quantity = 5, ItemCost = 93.23 });
            report.Add(new CustomersReport { Name = "Yen I. Schneider,", RegisterDate = "March 30th, 2018", LastBuy = "September 12th, 2017", Item = "viverra. Ma", Quantity = 3, ItemCost = 88.45 });
            report.Add(new CustomersReport { Name = "Zahir Q. Rowe,", RegisterDate = "April 19th, 2018", LastBuy = "November 8th, 2017", Item = "facilisis vitae,", Quantity = 5, ItemCost = 57.41 });
            report.Add(new CustomersReport { Name = "Jasper F. Vang,", RegisterDate = "August 26th, 2017", LastBuy = "January 8th, 2019", Item = "nec, e", Quantity = 4, ItemCost = 66.54 });
            report.Add(new CustomersReport { Name = "Rigel E. Foreman,", RegisterDate = "August 10th, 2018", LastBuy = "January 25th, 2018", Item = "at, l", Quantity = 3, ItemCost = 81.12 });
            report.Add(new CustomersReport { Name = "Breanna T. Gonzales,", RegisterDate = "October 10th, 2017", LastBuy = "February 28th, 2019", Item = "Proin non,", Quantity = 5, ItemCost = 14.03 });
            report.Add(new CustomersReport { Name = "Mikayla L. Haney,", RegisterDate = "December 19th, 2017", LastBuy = "May 26th, 2018", Item = "ligula. D", Quantity = 7, ItemCost = 66.06 });
            report.Add(new CustomersReport { Name = "Hanae G. Witt,", RegisterDate = "July 23rd, 2018", LastBuy = "January 29th, 2019", Item = "augue porttitor,", Quantity = 5, ItemCost = 29.70 });
            report.Add(new CustomersReport { Name = "Duncan Y. Browning,", RegisterDate = "April 20th, 2018", LastBuy = "September 13th, 2017", Item = "sociosqu ad,", Quantity = 5, ItemCost = 82.85 });
            report.Add(new CustomersReport { Name = "Nissim R. Walter,", RegisterDate = "April 23rd, 2018", LastBuy = "March 25th, 2019", Item = "Cras interdum.", Quantity = 1, ItemCost = 31.84 });
            report.Add(new CustomersReport { Name = "Amity T. Kline,", RegisterDate = "January 12th, 2018", LastBuy = "December 20th, 2018", Item = "eget, v", Quantity = 2, ItemCost = 63.66 });
            report.Add(new CustomersReport { Name = "Quinn D. Fisher,", RegisterDate = "April 15th, 2018", LastBuy = "October 27th, 2018", Item = "Nunc commodo,", Quantity = 6, ItemCost = 23.67 });
            report.Add(new CustomersReport { Name = "Castor Q. West,", RegisterDate = "August 3rd, 2018", LastBuy = "September 8th, 2017", Item = "Fusce aliquet,", Quantity = 5, ItemCost = 49.37 });
            report.Add(new CustomersReport { Name = "Clio V. Reilly,", RegisterDate = "June 4th, 2018", LastBuy = "February 18th, 2019", Item = "justo. P", Quantity = 9, ItemCost = 43.05 });
            report.Add(new CustomersReport { Name = "Jocelyn T. Allen,", RegisterDate = "September 16th, 2017", LastBuy = "December 23rd, 2017", Item = "libero est,", Quantity = 8, ItemCost = 37.91 });
            report.Add(new CustomersReport { Name = "Beverly C. Roberson,", RegisterDate = "July 17th, 2018", LastBuy = "December 26th, 2017", Item = "Praesent eu,", Quantity = 5, ItemCost = 92.42 });
            report.Add(new CustomersReport { Name = "Hermione T. Bray,", RegisterDate = "June 26th, 2018", LastBuy = "July 27th, 2018", Item = "Phasellus dapibus,", Quantity = 5, ItemCost = 53.66 });
            report.Add(new CustomersReport { Name = "Leo V. Gates,", RegisterDate = "July 21st, 2018", LastBuy = "February 4th, 2018", Item = "lectus pede,", Quantity = 5, ItemCost = 45.37 });
            report.Add(new CustomersReport { Name = "Lester K. Parks,", RegisterDate = "October 8th, 2017", LastBuy = "December 1st, 2017", Item = "leo elementum,", Quantity = 5, ItemCost = 29.48 });
            report.Add(new CustomersReport { Name = "Minerva O. Buchanan,", RegisterDate = "January 16th, 2018", LastBuy = "May 17th, 2019", Item = "eget ipsum.", Quantity = 1, ItemCost = 50.35 });
            report.Add(new CustomersReport { Name = "Nola C. Kirkland,", RegisterDate = "May 12th, 2018", LastBuy = "March 3rd, 2019", Item = "in consectetuer,", Quantity = 5, ItemCost = 10.81 });
            report.Add(new CustomersReport { Name = "Ralph G. Haley,", RegisterDate = "September 1st, 2017", LastBuy = "October 14th, 2018", Item = "erat eget,", Quantity = 5, ItemCost = 3.77 });
            report.Add(new CustomersReport { Name = "Colby D. Bradford,", RegisterDate = "April 22nd, 2018", LastBuy = "November 8th, 2017", Item = "Nulla interdum.", Quantity = 2, ItemCost = 56.15 });
            report.Add(new CustomersReport { Name = "Paki H. Chang,", RegisterDate = "July 2nd, 2018", LastBuy = "June 30th, 2019", Item = "orci quis,", Quantity = 5, ItemCost = 35.78 });
            report.Add(new CustomersReport { Name = "Timothy L. Glass,", RegisterDate = "November 1st, 2017", LastBuy = "May 22nd, 2018", Item = "sapien. N", Quantity = 6, ItemCost = 11.16 });
            report.Add(new CustomersReport { Name = "Colton Q. Carney,", RegisterDate = "June 29th, 2018", LastBuy = "July 23rd, 2019", Item = "luctus. C", Quantity = 3, ItemCost = 26.85 });
            report.Add(new CustomersReport { Name = "Madison Q. Townsend,", RegisterDate = "May 10th, 2018", LastBuy = "January 15th, 2018", Item = "dis parturient,", Quantity = 5, ItemCost = 46.81 });
            report.Add(new CustomersReport { Name = "Hadassah U. Wright,", RegisterDate = "August 13th, 2018", LastBuy = "May 24th, 2018", Item = "leo. C", Quantity = 8, ItemCost = 71.05 });
            report.Add(new CustomersReport { Name = "Shana M. Mcbride,", RegisterDate = "May 24th, 2018", LastBuy = "January 17th, 2018", Item = "blandit enim,", Quantity = 5, ItemCost = 2.20 });
            report.Add(new CustomersReport { Name = "Gloria I. Miller,", RegisterDate = "May 20th, 2018", LastBuy = "October 6th, 2017", Item = "tristique senectus,", Quantity = 5, ItemCost = 51.11 });
            report.Add(new CustomersReport { Name = "Brent F. Carter,", RegisterDate = "February 13th, 2018", LastBuy = "November 8th, 2017", Item = "luctus ut,", Quantity = 9, ItemCost = 27.70 });
            report.Add(new CustomersReport { Name = "Garrett B. Dalton,", RegisterDate = "August 26th, 2017", LastBuy = "July 26th, 2018", Item = "eget metus.", Quantity = 10, ItemCost = 64.87 });
            report.Add(new CustomersReport { Name = "Aileen Z. Lott,", RegisterDate = "December 12th, 2017", LastBuy = "September 2nd, 2018", Item = "tristique neque,", Quantity = 5, ItemCost = 24.49 });
            report.Add(new CustomersReport { Name = "Megan H. Kramer,", RegisterDate = "January 10th, 2018", LastBuy = "February 9th, 2019", Item = "risus. D", Quantity = 2, ItemCost = 25.10 });
            report.Add(new CustomersReport { Name = "Lane C. Haley,", RegisterDate = "November 9th, 2017", LastBuy = "November 13th, 2018", Item = "Etiam laoreet,", Quantity = 7, ItemCost = 95.30 });
            report.Add(new CustomersReport { Name = "Davis X. Estes,", RegisterDate = "March 8th, 2018", LastBuy = "December 12th, 2018", Item = "tristique ac,", Quantity = 10, ItemCost = 17.83 });
            report.Add(new CustomersReport { Name = "Noelle K. Bowman,", RegisterDate = "December 25th, 2017", LastBuy = "October 17th, 2018", Item = "accumsan neque,", Quantity = 5, ItemCost = 12.96 });
            report.Add(new CustomersReport { Name = "Akeem L. Crosby,", RegisterDate = "November 26th, 2017", LastBuy = "October 4th, 2018", Item = "dictum eu,", Quantity = 5, ItemCost = 89.98 });
            report.Add(new CustomersReport { Name = "Lillith B. Jordan,", RegisterDate = "August 25th, 2017", LastBuy = "August 13th, 2018", Item = "nulla. I", Quantity = 9, ItemCost = 95.23 });
            report.Add(new CustomersReport { Name = "Kevyn A. Owens,", RegisterDate = "October 24th, 2017", LastBuy = "October 12th, 2018", Item = "mattis ornare,", Quantity = 8, ItemCost = 95.68 });
            report.Add(new CustomersReport { Name = "Dorian W. Pickett,", RegisterDate = "November 8th, 2017", LastBuy = "February 13th, 2018", Item = "semper cursus.", Quantity = 8, ItemCost = 36.82 });
            report.Add(new CustomersReport { Name = "Coby Q. Osborne,", RegisterDate = "July 10th, 2018", LastBuy = "October 21st, 2018", Item = "Phasellus at,", Quantity = 5, ItemCost = 22.75 });
            report.Add(new CustomersReport { Name = "Peter X. Walsh,", RegisterDate = "November 7th, 2017", LastBuy = "March 14th, 2019", Item = "at augue,", Quantity = 5, ItemCost = 36.91 });
            report.Add(new CustomersReport { Name = "Maryam G. Santos,", RegisterDate = "September 24th, 2017", LastBuy = "August 28th, 2018", Item = "vulputate eu,", Quantity = 1, ItemCost = 92.91 });
            report.Add(new CustomersReport { Name = "Aspen P. Knowles,", RegisterDate = "February 27th, 2018", LastBuy = "March 25th, 2018", Item = "In tincidunt,", Quantity = 5, ItemCost = 50.23 });
            report.Add(new CustomersReport { Name = "Jana U. Newton,", RegisterDate = "June 23rd, 2018", LastBuy = "July 29th, 2019", Item = "facilisis non,", Quantity = 4, ItemCost = 41.72 });
            report.Add(new CustomersReport { Name = "Martena Y. Swanson,", RegisterDate = "September 15th, 2017", LastBuy = "August 26th, 2018", Item = "dui nec,", Quantity = 5, ItemCost = 54.90 });
            report.Add(new CustomersReport { Name = "Tarik R. Carter,", RegisterDate = "July 13th, 2018", LastBuy = "February 23rd, 2018", Item = "dui nec,", Quantity = 5, ItemCost = 89.92 });
            report.Add(new CustomersReport { Name = "Rinah R. Landry,", RegisterDate = "March 11th, 2018", LastBuy = "December 25th, 2017", Item = "Cras eget,", Quantity = 5, ItemCost = 44.32 });
            report.Add(new CustomersReport { Name = "Kermit I. Nielsen,", RegisterDate = "December 20th, 2017", LastBuy = "November 5th, 2018", Item = "a ultricies,", Quantity = 5, ItemCost = 97.43 });
            report.Add(new CustomersReport { Name = "Quinn D. Burks,", RegisterDate = "February 8th, 2018", LastBuy = "January 24th, 2019", Item = "tempor, e", Quantity = 4, ItemCost = 86.74 });
            report.Add(new CustomersReport { Name = "Perry E. Marquez,", RegisterDate = "June 3rd, 2018", LastBuy = "February 13th, 2018", Item = "Curabitur egestas,", Quantity = 5, ItemCost = 66.04 });
            report.Add(new CustomersReport { Name = "Summer X. Haley,", RegisterDate = "August 30th, 2017", LastBuy = "March 31st, 2018", Item = "justo faucibus,", Quantity = 5, ItemCost = 12.99 });
            report.Add(new CustomersReport { Name = "Cruz Z. May,", RegisterDate = "February 3rd, 2018", LastBuy = "February 11th, 2018", Item = "magnis dis,", Quantity = 5, ItemCost = 33.70 });
            report.Add(new CustomersReport { Name = "Yuri S. Blackburn,", RegisterDate = "March 23rd, 2018", LastBuy = "September 13th, 2017", Item = "Nam consequat,", Quantity = 5, ItemCost = 45.52 });
            report.Add(new CustomersReport { Name = "Curran W. Dillon,", RegisterDate = "March 2nd, 2018", LastBuy = "January 31st, 2019", Item = "lectus justo,", Quantity = 5, ItemCost = 31.22 });
            report.Add(new CustomersReport { Name = "Marshall F. Sharpe,", RegisterDate = "April 18th, 2018", LastBuy = "October 3rd, 2017", Item = "mollis vitae,", Quantity = 9, ItemCost = 24.15 });
            report.Add(new CustomersReport { Name = "Carl X. Rasmussen,", RegisterDate = "December 5th, 2017", LastBuy = "March 6th, 2019", Item = "libero. M", Quantity = 9, ItemCost = 76.16 });
            report.Add(new CustomersReport { Name = "Lawrence Z. Morgan,", RegisterDate = "December 29th, 2017", LastBuy = "March 13th, 2019", Item = "diam. P", Quantity = 9, ItemCost = 73.61 });
            report.Add(new CustomersReport { Name = "Chandler C. Yang,", RegisterDate = "August 14th, 2018", LastBuy = "August 30th, 2018", Item = "natoque penatibus,", Quantity = 5, ItemCost = 41.43 });
            report.Add(new CustomersReport { Name = "Rae Z. Walsh,", RegisterDate = "April 9th, 2018", LastBuy = "April 11th, 2018", Item = "ultricies ligula.", Quantity = 5, ItemCost = 40.17 });
            report.Add(new CustomersReport { Name = "Nathan F. Green,", RegisterDate = "January 11th, 2018", LastBuy = "February 23rd, 2018", Item = "varius. N", Quantity = 4, ItemCost = 27.87 });
            report.Add(new CustomersReport { Name = "Kiara J. Flynn,", RegisterDate = "April 14th, 2018", LastBuy = "April 23rd, 2019", Item = "Pellentesque habitant,", Quantity = 5, ItemCost = 26.24 });
            report.Add(new CustomersReport { Name = "Raya H. England,", RegisterDate = "July 2nd, 2018", LastBuy = "January 5th, 2019", Item = "In mi,", Quantity = 5, ItemCost = 9.02 });
            report.Add(new CustomersReport { Name = "Shellie O. Torres,", RegisterDate = "June 17th, 2018", LastBuy = "September 22nd, 2017", Item = "egestas a,", Quantity = 8, ItemCost = 50.13 });
            return report;
        }
    }
}
