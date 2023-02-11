using System;
using System.Collections.Generic;
using IronXL;

namespace Program
{
    public class Student
    {
        public int sno, totalMarks, tMarks = 100;
        public long contact_no;
        public string name, address, guardian_name, class_name;
        public DateTime added_on, s_dob;

        public Student(int sno, string name, string address, long contact_no, string guardian_name, string class_name, DateTime s_dob, DateTime added_on, int totalMarks)
        {
            this.sno = sno;
            this.name = name.ToLower();
            this.address = address.ToLower();
            this.contact_no = contact_no;
            this.guardian_name = guardian_name.ToLower();
            this.class_name = class_name.ToLower();
            this.s_dob = s_dob;
            this.added_on = added_on;
            this.setTotalMarks(totalMarks);
        }

        public void setTotalMarks(int totalMarks)
        {
            if (totalMarks < 0 || totalMarks > tMarks)
            {
                Console.Write("\n\n *Marks should be Non-negative and less than " + tMarks);
                totalMarks = 0;
            }

            this.totalMarks = totalMarks;
        }

    }

    public class Program
    {
        public static List<Student> students = new List<Student>();
        static int sno = 0, totalMarks, tMarks = 100;

        static long contact_no;
        static string name, address, guardian_name, class_name;
        static DateTime added_on, s_dob;

        static WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
        static WorkSheet sheet = workbook.CreateWorkSheet("Student");

        public static void Main(string[] args)
        {
            IronXL.License.LicenseKey = "IRONXL.UTSAVBHATTARAI.30717-2FC590FF9C-AAE6JN-OUSWNKNYAU6Y-UEEABLOHR52Y-FMDSDAYOYUSL-CVTWOHQZMVKD-I7FPEVJJFSKS-NCZUD5-T3BMSTK5IEGJEA-DEPLOYMENT.TRIAL-4CHDVT.TRIAL.EXPIRES.06.MAR.2023";
            sheet["A1"].Value = "Sno";
            sheet["B1"].Value = "Name";
            sheet["C1"].Value = "Address";
            sheet["D1"].Value = "Contact No";
            sheet["E1"].Value = "Guardian Name";
            sheet["F1"].Value = "Class";
            sheet["G1"].Value = "Date of Birth";
            sheet["H1"].Value = "Total Marks";
            sheet["I1"].Value = "Added on";

            string ch = "";
            getStudent();
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
                                students.Add(new Student(sno, name, address, contact_no, guardian_name, class_name, s_dob, added_on, totalMarks));
                            break;

                        case 2:
                            updateStudent(); break;

                        case 3:
                            deleteStudent(); break;

                        case 4:
                            enterName();
                            viewStudent(name); break;

                        case 5:
                            viewAllStudents(); break;

                        case 6:
                            viewStudentResult(); break;

                        case 7:
                            save(); break;
                        case 8:
                            fileAction();
                            break;
                        case 9:
                            fileProtection();
                            break;
                        case 10:
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

            Console.WriteLine("1. New Student\n2. Update Student\n3. Delete Student\n4. View Student\n5. View All Students\n6. View Student Result\n7. Save\n8. Hide/Unhide the File\n9. Protect/Unprotect file\n10. Exit\n");
        }

        public static void enterDetails()
        {
            Console.WriteLine("\n\nPlease provide the details below");
            Console.WriteLine("------------------------------------");

            enterSno();
            enterName();
            enterAddress();
            enterContactNo();
            enterGuardianName();
            enterClassName();
            enterDateOfBirth();
            enterTotalMarks();
            enterAddedOn();
        }

        public static void enterSno()
        {
            ++sno;
        }

        public static void enterName()
        {
            Console.Write("\nEnter Name : ");
            name = Console.ReadLine().ToLower();
        }

        public static void enterAddress()
        {
            Console.Write("\nEnter Address : ");
            address = Console.ReadLine().ToLower();
        }


        public static void enterContactNo()
        {
            //get the contact no and check if it is long numeric or not
            bool IsConversionSuccessful;
            do
            {
                Console.Write("\nEnter Contact No : ");
                IsConversionSuccessful = long.TryParse(Console.ReadLine(), out contact_no);
                if (!IsConversionSuccessful)
                    Console.Write("\nPlease enter a valid format");
            } while (!IsConversionSuccessful || contact_no.ToString().Length != 10 || contact_no.ToString().StartsWith("0"));
        }

        public static void enterGuardianName()
        {
            Console.Write("\nEnter Guardian Name : ");
            guardian_name = Console.ReadLine().ToLower();
        }

        public static void enterClassName()
        {
            Console.Write("\nEnter Class Name : ");
            class_name = Console.ReadLine().ToLower();
        }

        public static void enterDateOfBirth()
        {
            Console.Write("\nEnter Date of Birth (eg: 12/12/2061) : ");
            s_dob = DateTime.Parse(Console.ReadLine());
        }

        public static void enterAddedOn()
        {
            added_on = DateTime.Now;
        }

        public static void enterTotalMarks()
        {
            Console.Write("\nEnter Total Marks : ");
            bool IsConversionSuccessful = int.TryParse(Console.ReadLine(), out totalMarks);

            if (!IsConversionSuccessful)
                Console.Write("\nPlease enter a valid format");
        }
        public static void fileAction()
        {
            //hide the file if it exits and unhide if it is hidden
            if (File.Exists("Students.xlsx"))
            {
                FileAttributes attributes = File.GetAttributes("Students.xlsx");
                if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                {
                    File.SetAttributes("Students.xlsx", attributes & ~FileAttributes.Hidden);
                    Console.WriteLine("\nFile is Unhidden");
                }
                else
                {
                    File.SetAttributes("Students.xlsx", attributes | FileAttributes.Hidden);
                    Console.WriteLine("\nFile is Hidden");
                }
            }
            else
            {
                Console.WriteLine("\nFile does not exists");
            }
        }

        public static void fileProtection()
        {
            //make the file unopenable if it is openable and vice versa
            if (File.Exists("Students.xlsx"))
            {
                FileAttributes attributes = File.GetAttributes("Students.xlsx");
                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes("Students.xlsx", attributes & ~FileAttributes.ReadOnly);
                    Console.WriteLine("\nFile is Unprotected");
                }
                else
                {
                    File.SetAttributes("Students.xlsx", attributes | FileAttributes.ReadOnly);
                    Console.WriteLine("\nFile is Protected");
                }
            }
            else
            {
                Console.WriteLine("\nFile does not exists");
            }

        }
        public static void getStudent()
        {
            //Load the workbook from the disk if it exists
            if (File.Exists("Students.xlsx"))
            {
                workbook = WorkBook.Load("Students.xlsx");
                sheet = workbook.WorkSheets[0];
                // get the rows from the sheet
                var rows = sheet.Rows.Count();
                for (int i = 2; i <= rows; i++)
                {
                    //set the values to the variables
                    sno = Convert.ToInt32(sheet["A" + i].Value);
                    name = sheet["B" + i].Value.ToString();
                    address = sheet["C" + i].Value.ToString();
                    contact_no = Convert.ToInt64(sheet["D" + i].Value);
                    guardian_name = sheet["E" + i].Value.ToString();
                    class_name = sheet["F" + i].Value.ToString();
                    s_dob = Convert.ToDateTime(sheet["G" + i].Value);
                    totalMarks = Convert.ToInt32(sheet["H" + i].Value);
                    added_on = Convert.ToDateTime(sheet["I" + i].Value);

                    //add the student to the list
                    students.Add(new Student(sno, name, address, contact_no, guardian_name, class_name, s_dob, added_on, totalMarks));
                }
            }
        }

        public static void updateStudent()
        {
            enterName();
            Student student = students.Find(x => x.name == name);

            if (student == null)
                Console.Write("\n\n\t *No matchng records found...");

            else
            {
                Console.WriteLine("\n\nStudent Details");
                Console.WriteLine("------------------------------------");
                Console.WriteLine($"Sno : {student.sno}\nName : {student.name}\nAddress : {student.address}\nContact No. : {student.contact_no}\nGuardian Name : {student.guardian_name}\nClass Name : {student.class_name}\nDate of Birth : {student.s_dob}\nTotal Marks : {student.totalMarks}");

                Console.WriteLine("\n1. Update Name\n2. Update Address\n3. Update Contact No.\n4. Update Guardian Name\n5. Update Class Name\n6. Update Date of Birth\n7. Update Total Marks\n8. Exit ");
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
                                enterName();
                                student.name = name; break;

                            case 2:
                                enterAddress();
                                student.address = address; break;
                            case 3:
                                enterContactNo();
                                student.contact_no = contact_no; break;
                            case 4:
                                enterGuardianName();
                                student.guardian_name = guardian_name; break;
                            case 5:
                                enterClassName();
                                student.class_name = class_name; break;
                            case 6:
                                enterDateOfBirth();
                                student.s_dob = s_dob; break;
                            case 7:
                                enterTotalMarks();
                                if (totalMarks != 0)
                                    student.setTotalMarks(totalMarks);
                                break;
                            case 8:
                                break;

                            default:
                                Console.WriteLine("\nInvalid Choice"); break;
                        }

                        do
                        {
                            Console.Write("\nDo you want to continue to update more.., Say (yes or no):");
                            ch = Console.ReadLine().Trim().ToLower();
                        } while (ch != "yes" && ch != "no");
                    }
                } while (ch == "yes");
            }
        }

        public static void deleteStudent()
        {
            enterName();
            Student student = students.Find(x => x.name == name);

            bool res = students.Remove(student);

            if (res)
                Console.Write($"\n  *{student.name}* removed successfully !!");
            else
                Console.Write("\n\n\t *No matching records found...");

        }

        public static Student viewStudent(string name)
        {
            Student student = students.Find(x => x.name == name);

            if (student == null)
            {
                Console.WriteLine("\n\n\t *No matchng records found...");
                return null;
            }
            else
            {
                Console.WriteLine("\n\nStudent Details");
                Console.WriteLine("--------------------");

                Console.WriteLine($"Sno : {student.sno}\nName : {student.name}\nAddress : {student.address}\nContact No. : {student.contact_no}\nGuardian Name : {student.guardian_name}\nClass Name : {student.class_name}\nDate of Birth : {student.s_dob}\nTotal Marks : {student.totalMarks}");

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
                    Console.WriteLine($"{counter++}.\tSno : {student.sno}\n\tName : {student.name}\n\tAddress : {student.address}\n\tContact No. : {student.contact_no}\n\tGuardian Name : {student.guardian_name}\n\tClass Name : {student.class_name}\n\tDate of Birth : {student.s_dob}\n\tTotal Marks : {student.totalMarks}\n");
                }
            }

        }

        public static void viewStudentResult()
        {
            enterName();

            Student student = viewStudent(name);

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
                Console.WriteLine("\t No Student Found!");
            }
        }

        public static void save()
        {
            var l = sheet.Rows.Count();
            for (int i = 2; i <= l; i++)
            {
                sheet["A" + i].Value = "";
                sheet["B" + i].Value = "";
                sheet["C" + i].Value = "";
                sheet["D" + i].Value = "";
                sheet["E" + i].Value = "";
                sheet["F" + i].Value = "";
                sheet["G" + i].Value = "";
                sheet["H" + i].Value = "";
                sheet["I" + i].Value = "";
            }
            int counter = 2;
            foreach (Student student in students)
            {
                sheet["A" + counter].Value = student.sno;
                sheet["B" + counter].Value = student.name;
                sheet["C" + counter].Value = student.address;
                sheet["D" + counter].Value = student.contact_no;
                sheet["E" + counter].Value = student.guardian_name;
                sheet["F" + counter].Value = student.class_name;
                sheet["G" + counter].Value = student.s_dob;
                sheet["H" + counter].Value = student.totalMarks;
                sheet["I" + counter].Value = student.added_on;
                counter++;
            }
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write("\nSaving");
            for (int i = 0; i < 3; i++)
            {
                Console.Write(".");
                Thread.Sleep(500);
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\nSaved");
            Console.ResetColor();
            workbook.SaveAs("Students.xlsx");
        }
    }
}