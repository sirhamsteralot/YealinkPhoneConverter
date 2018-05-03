using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YealinkPhoneConverter
{
    public class Contact
    {
        private string firstName;
        private string lastName;
        private string mobile;
        private string work;

        public string FirstName { get => firstName; set => firstName = value; }
        public string LastName { get => lastName; set => lastName = value; }
        public string Mobile { get => mobile; set => mobile = value; }
        public string Work { get => work; set => work = value; }
    }
}
