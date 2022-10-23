using Mail.WindowsService.DTOClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mail.WindowsService.DataProvider
{
    public class DataProvide
    {

        private List<Person> People;

        public DataProvide()
        {

            People = new List<Person>
            {

                new Person
                {
                    PersonName = "Rober Downe Jr.",
                    Address = "Somewhere Rainbow",
                    CarPlate = "34 APR 24",
                    Nation = "American",
                    Notes = "Clever",
                    StartDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day),
                    FinishDate = new DateTime((DateTime.Now.Year + 1), DateTime.Now.Month, DateTime.Now.Day)
                },
                new Person
                {
                    PersonName = "Chris Evan",
                    Address = "Somewhere California",
                    CarPlate = "12 TB 244",
                    Nation = "Canadian",
                    Notes = "Fast and Sturdy",
                    StartDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day),
                    FinishDate = new DateTime((DateTime.Now.Year + 1), DateTime.Now.Month, DateTime.Now.Day)
                },
                new Person
                {
                    PersonName = "Chris Prey",
                    Address = "Somewhere Around",
                    CarPlate = "POLK 12 24",
                    Nation = "Columbian",
                    Notes = "Clever",
                    StartDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day),
                    FinishDate = new DateTime((DateTime.Now.Year + 1), DateTime.Now.Month, DateTime.Now.Day)
                },
                new Person
                {
                    PersonName = "Tom Holland",
                    Address = "Rainbow Avenue",
                    CarPlate = "34 APR 24",
                    Nation = "Peruvian",
                    Notes = "Fast and clever",
                    StartDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day),
                    FinishDate = new DateTime((DateTime.Now.Year + 1), DateTime.Now.Month, DateTime.Now.Day)
                },

            };

        }

        public List<Person> ProvideData()
        {
            return People;
        }

    }
}
