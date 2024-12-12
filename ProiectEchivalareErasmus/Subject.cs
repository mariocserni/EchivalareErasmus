using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProiectEchivalareErasmus
{
    public class Subject
    {
        public int Number { get; set; }
        public string DisciplineUPT { get; set; }
        public int Year { get; set; }
        public int Semester { get; set; }
        public int CreditsUPT { get; set; }
        public string HostDiscipline { get; set; }
        public int HostCredits { get; set; }
        public int Grade { get; set; } 

        public Subject(int number, string disciplineUPT, int year, int semester, int creditsUPT, string hostDiscipline, int hostCredits, int grade)
        {
            Number = number;
            DisciplineUPT = disciplineUPT;
            Year = year;
            Semester = semester;
            CreditsUPT = creditsUPT;
            HostDiscipline = hostDiscipline;
            HostCredits = hostCredits;
            Grade = grade; 
        }


    }


}
