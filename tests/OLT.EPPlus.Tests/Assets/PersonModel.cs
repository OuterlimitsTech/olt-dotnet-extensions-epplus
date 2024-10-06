using OLT.Core;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OLT.EPPlus.Tests.Assets
{
    public class PersonModel : OltPersonName
    {
        public string Email { get; set; }

        public static PersonModel FakerData()
        {
            return new PersonModel
            {
                First = Faker.Name.First(),
                Middle = Faker.Name.Middle(),
                Last = Faker.Name.Last(),
                Email = Faker.Internet.Email(),
            };
        }

        public static List<PersonModel> FakerList(int number)
        {
            var list = new List<PersonModel>();

            for (int i = 0; i < number; i++)
            {
                list.Add(FakerData());
            }

            return list;
        }
    }
}