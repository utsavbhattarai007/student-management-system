using System;
using System.Collections.Generic;

namespace Program
{

    public class Student
    {
        public string name, city;
        public int sno, totalMarks = 0;

        public Student(int sno, string name, string city, int totalMarks)
        {
            this.sno = sno;
            this.name = name.ToLower();
            this.city = city.ToLower();
            this.setTotalMarks(totalMarks);

        }

        public void setTotalMarks(int totalMarks)
        {
            if (totalMarks < 0 || totalMarks > 100)
            {
                Console.Write("\n\n *Marks should be Non-negative and less than 100");
                totalMarks = 0;
            }

            this.totalMarks = totalMarks;
        }

    }

    public class Program
    {
        public static List<Student> students = new List<Student>();
        static int sno, totalMarks;
        static string name, city;

        public static void Main(string[] args)
        {
            string ch = "";
            msg();

            do
            {

                Console.Write("\nSelect your choice : ");
                int choice = 0;
                bool IsConversionSuccessful = int.TryParse(Console.ReadLine(), out choice);

                if (!IsConversionSuccessful)
                    Console.Write("\nPlease enter a valid format");
                else
                {
                    switch (choice)
                    {
                        case 1:
                            enterDetails();
                            bool exists = students.Exists(x => x.sno == sno);
                            if (exists)
                            {
                                //clear the console and display the message
                                Console.Clear();
                                msg();
                                Console.Write("\n\n *Sno already exists");
                            }
                            else
                                students.Add(new Student(sno, name, city, totalMarks));
                            break;

                        case 2:
                            updateStudent(); break;

                        case 3:
                            deleteStudent(); break;

                        case 4:
                            enterSno();
                            viewStudent(sno); break;

                        case 5:
                            viewAllStudents(); break;

                        case 6:
                            viewStudentResult(); break;

                        case 7:
                            return;

                        default:
                            Console.Clear();
                            msg();
                            Console.Write("\nInvalid Choice, Please enter a valid choice");
                            break;
                    }

                    do
                    {

                        Console.Write("\n\nDo you want to continue to the Application, Say (yes or no): ");
                        ch = Console.ReadLine().Trim().ToLower();

                        if (ch != "yes" && ch != "no")
                        {

                            Console.Write("\n\nInvalid Choice, Please say (yes or no)");
                        }
                        Console.Clear();
                        msg();

                    } while (ch != "yes" && ch != "no");
                }

            } while (ch == "yes");

            Console.WriteLine("\n\n  *** Program Terminated *** ");
        }

        public static void msg()
        {
            Console.WriteLine("\n\tSTUDENT MANAGEMENT SYSTEM");
            Console.WriteLine("----------------------------------------");

            Console.WriteLine("1. New Student\n2. Update Student\n3. Delete Student\n4. View Student\n5. View All Students\n6. View Student Result\n7. Exit\n");
        }

        public static void enterDetails()
        {
            Console.WriteLine("\n\nPlease provide the details below");
            Console.WriteLine("------------------------------------");

            enterSno();
            enterName();
            enterCity();
            enterTotalMarks();

        }

        public static void enterSno()
        {
            Console.Write("\nEnter Sno : ");
            bool IsConversionSuccessful = int.TryParse(Console.ReadLine(), out sno);

            if (!IsConversionSuccessful)
                Console.Write("\nPlease enter a valid format");
        }

        public static void enterName()
        {
            Console.Write("\nEnter Name : ");
            name = Console.ReadLine().ToLower();
        }

        public static void enterCity()
        {
            Console.Write("\nEnter City : ");
            city = Console.ReadLine().ToLower();
        }

        public static void enterTotalMarks()
        {
            Console.Write("\nEnter Total Marks : ");
            bool IsConversionSuccessful = int.TryParse(Console.ReadLine(), out totalMarks);

            if (!IsConversionSuccessful)
                Console.Write("\nPlease enter a valid format");
        }


        public static void updateStudent()
        {
            enterDetails();
            Student student = students.Find(x => x.sno == sno && x.name == name && x.city == city && x.totalMarks == totalMarks);

            if (student == null)
                Console.Write("\n\n\t *No matchng records found...");

            else
            {
                Console.WriteLine("\n\nStudent Details");
                Console.WriteLine("------------------------------------");
                Console.WriteLine($"Sno : {student.sno}\nName : {student.name}\nCity : {student.city}\nTotal Marks : {student.totalMarks}");

                Console.WriteLine("\n1. Update Sno\n2. Update Name\n3. Update City\n4. Update Total Marks\n");
                string ch = "";

                do
                {
                    Console.Write("\nSelect the property which you want to update : ");
                    int choice = 0;
                    bool IsConversionSuccessful = int.TryParse(Console.ReadLine(), out choice);

                    if (!IsConversionSuccessful)
                        Console.Write("\nPlease enter a valid format");
                    else
                    {

                        switch (choice)
                        {
                            case 1:
                                enterSno();
                                if (sno != 0)
                                {
                                    bool exists = students.Exists(x => x.sno == sno);
                                    if (exists)
                                        Console.Write("\n\n *Sno already exists");
                                    else
                                        student.sno = sno;
                                }
                                break;

                            case 2:
                                enterName();
                                student.name = name; break;

                            case 3:
                                enterCity();
                                student.city = city; break;

                            case 4:
                                enterTotalMarks();
                                if (totalMarks != 0)
                                    student.setTotalMarks(totalMarks);
                                break;

                            default:
                                Console.WriteLine("\nInvalid Choice"); break;
                        }

                        do
                        {
                            Console.WriteLine("\n\nDo you want to continue to update more.., Say (yes or no)");
                            ch = Console.ReadLine().Trim().ToLower();
                            Console.Write("\n\nInvalid Choice, Please say (yes or no)");
                        } while (ch != "yes" && ch != "no");
                    }
                } while (ch == "yes");
            }
        }

        public static void deleteStudent()
        {
            enterSno();
            Student student = students.Find(x => x.sno == sno);

            bool res = students.Remove(student);

            if (res)
                Console.Write($"\n  *{student.sno}* removed successfully !!");
            else
                Console.Write("\n\n\t *No matching records found...");

        }

        public static Student viewStudent(int sno)
        {
            Student student = students.Find(x => x.sno == sno);

            if (student == null)
            {
                Console.WriteLine("\n\n\t *No matchng records found...");
                return null;
            }
            else
            {
                Console.WriteLine("\n\nStudent Details");
                Console.WriteLine("--------------------");

                Console.WriteLine($"Sno : {student.sno}\nName : {student.name}\nCity : {student.city}\nTotal Marks : {student.totalMarks}");

                return student;
            }

        }

        public static void viewAllStudents()
        {
            Console.WriteLine("\n\nStudents Details");
            Console.WriteLine("---------------------");

            if (students.Count == 0)
                Console.Write("\n0 records");
            else
            {
                int counter = 0;
                foreach (Student student in students)
                {
                    Console.WriteLine($"{counter++}.\tSno : {student.sno}\n\tName : {student.name}\n\tCity : {student.city}\n\tTotal Marks : {student.totalMarks}\n");
                }
            }

        }

        public static void viewStudentResult()
        {
            enterSno();

            Student student = viewStudent(sno);

            if (student != null)
            {
                if (student.totalMarks >= 75)
                    Console.WriteLine("Result : First Class");
                else if (student.totalMarks >= 50)
                    Console.WriteLine("Result : Second Class");
                else if (student.totalMarks >= 35)
                    Console.WriteLine("Result : Third Class");
                else
                    Console.WriteLine("Result : Fail");
            }
            else
            {
                Console.WriteLine("No Student Found!");
            }
        }
    }
}