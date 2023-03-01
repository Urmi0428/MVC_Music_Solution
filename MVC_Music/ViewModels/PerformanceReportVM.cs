using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace MVC_Music.ViewModels
{
    public class PerformanceReportVM
    {

        [Display(Name = "Musician")]
        public string FormalName
        {
            get
            {
                return LastName + ", " + FirstName
                    + (string.IsNullOrEmpty(MiddleName) ? "" :
                        (" " + (char?)MiddleName[0] + ".").ToUpper());
            }
        }

        public int ID { get; set; }

        public string FirstName { get; set; }

        public string MiddleName { get; set; }

        public string LastName { get; set; }

        [Display(Name = "Average Fee Paid")]
        [DataType(DataType.Currency)]
        public double AverageFeePaid { get; set; }

        [Display(Name = "Highest Fee Paid")]
        [DataType(DataType.Currency)]
        public double HighestFeePaid { get; set; }

        [Display(Name = "Lowest Fee Paid")]
        [DataType(DataType.Currency)]
        public double LowestFeePaid { get; set; }

        [Display(Name = "Total number of Performances")]
        public int NumberofPerformances { get; set; }


    }
}
