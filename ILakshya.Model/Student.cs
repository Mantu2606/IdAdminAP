using System.ComponentModel.DataAnnotations;

namespace ILakshya.Model
{
    public class Student
    {
        /*  public int Id { get; set; } // Auto-incrementing ID
          public int? EnrollNo { get; set; } // Enroll number (nullable)
          public string Name { get; set; }
          public string FatherName { get; set; }
          public string RollNo { get; set; }
          public int GenKnowledge { get; set; }
          public int Science { get; set; }
          public int EnglishI { get; set; }
          public int EnglishII { get; set; }
          public int HindiI { get; set; }
          public int HindiII { get; set; }
          public int Computer { get; set; }
          public int Sanskrit { get; set; }
          public int Mathematics { get; set; }
          public int SocialStudies { get; set; }
          public int MaxMarks { get; set; } = 5; // Assuming max marks are 5 for all subjects
          public int PassMarks { get; set; } = 2; // Assuming pass marks are 2 for all subjects
          */

        public int Id { get; set; } // Auto-incrementing ID
        public int? EnrollNo { get; set; } // Enroll number (nullable)
        public string Name { get; set; }
        public string FatherName { get; set; }
        public string RollNo { get; set; }
        public int GenKnowledge { get; set; }
        public int Science { get; set; }
        public int EnglishI { get; set; }
        public int EnglishII { get; set; }
        public int HindiI { get; set; }
        public int HindiII { get; set; }
        public int Computer { get; set; }
        public int Sanskrit { get; set; }
        public int Mathematics { get; set; }
        public int SocialStudies { get; set; }
        public int MaxMarks { get; set; } = 5; // Assuming max marks are 5 for all subjects
        public int PassMarks { get; set; } = 2; // Assuming pass marks are 2 for all subjects


    }
}


