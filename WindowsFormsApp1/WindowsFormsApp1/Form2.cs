using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp1
{

    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();



        }

        private void  FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {

            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);



        }

        private void CreateWordDocument(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                this.FindAndReplace(wordApp, "<Coursename>", txtCourseName.Text);




                this.FindAndReplace(wordApp, "<coursecode>", txtCourseId.Text);
                this.FindAndReplace(wordApp, "<fall>", txtFall.Text);
                this.FindAndReplace(wordApp, "<spring>", txtSpring.Text);
                this.FindAndReplace(wordApp, "<Theory>", txtTheory.Text);
                this.FindAndReplace(wordApp, "<Application>", txtApplication.Text);
                this.FindAndReplace(wordApp, "<Local>", txtLocalCredits.Text);

                this.FindAndReplace(wordApp, "<prere>", txtPrere.Text);

                this.FindAndReplace(wordApp, "<English>", labeleng.Text);
                this.FindAndReplace(wordApp, "<Turkish>", labelturk.Text);
                this.FindAndReplace(wordApp, "<SecondForeignLanguage>", labelsecond.Text);
                this.FindAndReplace(wordApp, "<Required>", labelrequired.Text);
                this.FindAndReplace(wordApp, "<Elective>", labelelective.Text);
                this.FindAndReplace(wordApp, "<ShortCycle>", labelshort.Text);
                this.FindAndReplace(wordApp, "<FirstCycle>", labelfirst.Text);
                this.FindAndReplace(wordApp, "<SecondCycle>", labelscnd.Text);
                this.FindAndReplace(wordApp, "<ThirdCycle>", labelthird.Text);

                this.FindAndReplace(wordApp, "<coursecor>", txtCourseCoordinator.Text);
                this.FindAndReplace(wordApp, "<courselec>", txtCourseLecturer.Text);
                this.FindAndReplace(wordApp, "<asistant>", txtAssistan.Text);

                this.FindAndReplace(wordApp, "<courseobj>", txtcourseobj.Text);
                this.FindAndReplace(wordApp, "<learning>", txtlearning.Text);
                this.FindAndReplace(wordApp, "<coursedesc>", txtcoursedesc.Text);


                this.FindAndReplace(wordApp, "<cor>", labelcore.Text);
                this.FindAndReplace(wordApp, "<maj>", labelmajor.Text);
                this.FindAndReplace(wordApp, "<sup>", labelsupportive.Text);
                this.FindAndReplace(wordApp, "<com>", labelcommunication.Text);
                this.FindAndReplace(wordApp, "<tran>", labeltransferable.Text);

                this.FindAndReplace(wordApp, "<s1>", s1.Text);
                this.FindAndReplace(wordApp, "<s2>", s2.Text);
                this.FindAndReplace(wordApp, "<s3>", s3.Text);
                this.FindAndReplace(wordApp, "<s4>", s4.Text);
                this.FindAndReplace(wordApp, "<s5>", s5.Text);
                this.FindAndReplace(wordApp, "<s6>", s6.Text);
                this.FindAndReplace(wordApp, "<s7>", s7.Text);
                this.FindAndReplace(wordApp, "<s8>", s8.Text);
                this.FindAndReplace(wordApp, "<s9>", s9.Text);
                this.FindAndReplace(wordApp, "<s10>", s10.Text);
                this.FindAndReplace(wordApp, "<s11>", s11.Text);
                this.FindAndReplace(wordApp, "<s12>", s12.Text);
                this.FindAndReplace(wordApp, "<s13>", s13.Text);
                this.FindAndReplace(wordApp, "<s14>", s14.Text);
                this.FindAndReplace(wordApp, "<s15>", s15.Text);
                this.FindAndReplace(wordApp, "<s16>", s16.Text);

                this.FindAndReplace(wordApp, "<r1>", r1.Text);
                this.FindAndReplace(wordApp, "<r2>", r2.Text);
                this.FindAndReplace(wordApp, "<r3>", r3.Text);
                this.FindAndReplace(wordApp, "<r4>", r4.Text);
                this.FindAndReplace(wordApp, "<r5>", r5.Text);
                this.FindAndReplace(wordApp, "<r6>", r6.Text);
                this.FindAndReplace(wordApp, "<r7>", r7.Text);
                this.FindAndReplace(wordApp, "<r8>", r8.Text);
                this.FindAndReplace(wordApp, "<r9>", r9.Text);
                this.FindAndReplace(wordApp, "<r10>", r10.Text);
                this.FindAndReplace(wordApp, "<r11>", r11.Text);
                this.FindAndReplace(wordApp, "<r12>", r12.Text);
                this.FindAndReplace(wordApp, "<r13>", r13.Text);
                this.FindAndReplace(wordApp, "<r14>", r14.Text);
                this.FindAndReplace(wordApp, "<r15>", r15.Text);
                this.FindAndReplace(wordApp, "<r16>", r16.Text);


                this.FindAndReplace(wordApp, "<coursenotes>", txtcoursenotes.Text);
                this.FindAndReplace(wordApp, "<suggested>", txtsuggested.Text);


                this.FindAndReplace(wordApp, "<n1>", n1.Text);
                this.FindAndReplace(wordApp, "<n2>", n2.Text);
                this.FindAndReplace(wordApp, "<n3>", n3.Text);
                this.FindAndReplace(wordApp, "<n4>", n4.Text);
                this.FindAndReplace(wordApp, "<n5>", n5.Text);
                this.FindAndReplace(wordApp, "<n6>", n6.Text);
                this.FindAndReplace(wordApp, "<n7>", n7.Text);
                this.FindAndReplace(wordApp, "<n8>", n8.Text);
                this.FindAndReplace(wordApp, "<n9>", n9.Text);
                this.FindAndReplace(wordApp, "<n10>", n10.Text);
                this.FindAndReplace(wordApp, "<n11>", n11.Text);
                this.FindAndReplace(wordApp, "<n12>", n12.Text);


                this.FindAndReplace(wordApp, "<w1>", w1.Text);
                this.FindAndReplace(wordApp, "<w2>", w2.Text);
                this.FindAndReplace(wordApp, "<w3>", w3.Text);
                this.FindAndReplace(wordApp, "<w4>", w4.Text);
                this.FindAndReplace(wordApp, "<w5>", w5.Text);
                this.FindAndReplace(wordApp, "<w6>", w6.Text);
                this.FindAndReplace(wordApp, "<w7>", w7.Text);
                this.FindAndReplace(wordApp, "<w8>", w8.Text);
                this.FindAndReplace(wordApp, "<w9>", w9.Text);
                this.FindAndReplace(wordApp, "<w10>", w10.Text);
                this.FindAndReplace(wordApp, "<w11>", w11.Text);
                this.FindAndReplace(wordApp, "<w12>", w12.Text);


                this.FindAndReplace(wordApp, "<wsem1>", w9.Text);
                this.FindAndReplace(wordApp, "<wsem2>", w10.Text);
                this.FindAndReplace(wordApp, "<wendsem1>", w11.Text);
                this.FindAndReplace(wordApp, "<wendsem2>", w12.Text);


                
                this.FindAndReplace(wordApp, "<nu3>", nu3.Text);
                this.FindAndReplace(wordApp, "<nu4>", nu4.Text);
                this.FindAndReplace(wordApp, "<nu5>", nu5.Text);
                this.FindAndReplace(wordApp, "<nu6>", nu6.Text);
                this.FindAndReplace(wordApp, "<nu7>", nu7.Text);
                this.FindAndReplace(wordApp, "<nu8>", nu8.Text);
                this.FindAndReplace(wordApp, "<nu9>", nu9.Text);
                this.FindAndReplace(wordApp, "<nu10>", nu10.Text);


                this.FindAndReplace(wordApp, "<d1>", d1.Text);
                this.FindAndReplace(wordApp, "<d2>", d2.Text);
                this.FindAndReplace(wordApp, "<d3>", d3.Text);
                this.FindAndReplace(wordApp, "<d4>", d4.Text);
                this.FindAndReplace(wordApp, "<d5>", d5.Text);
                this.FindAndReplace(wordApp, "<d6>", d6.Text);
                this.FindAndReplace(wordApp, "<d7>", d7.Text);
                this.FindAndReplace(wordApp, "<d8>", d8.Text);
                this.FindAndReplace(wordApp, "<d9>", d9.Text);
                this.FindAndReplace(wordApp, "<d10>", d10.Text);


                this.FindAndReplace(wordApp, "<wo1>", wo1.Text);
                this.FindAndReplace(wordApp, "<wo2>", wo2.Text);
                this.FindAndReplace(wordApp, "<wo3>", wo3.Text);
                this.FindAndReplace(wordApp, "<wo4>", wo4.Text);
                this.FindAndReplace(wordApp, "<wo5>", wo5.Text);
                this.FindAndReplace(wordApp, "<wo6>", wo6.Text);
                this.FindAndReplace(wordApp, "<wo7>", wo7.Text);
                this.FindAndReplace(wordApp, "<wo8>", wo8.Text);
                this.FindAndReplace(wordApp, "<wo9>", wo9.Text);
                this.FindAndReplace(wordApp, "<wo10>", wo10.Text);



            }
            else
            {
                MessageBox.Show("File not Found!");
            }

            //Save as
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created!");
        }


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {


        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtCourseName_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtTheory_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if(!Char.IsDigit(ch)&& ch!= 8 && ch != 46) // 8 is a backspace key 56 is a delete key
            {

                e.Handled = true;

            }
        }

        private void txtApplication_KeyPress(object sender, KeyPressEventArgs e)
        {

            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46) // 8 is a backspace key 56 is a delete key
            {

                e.Handled = true;

            }

        }

        private void txtLocalCredits_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46) // 8 is a backspace key 56 is a delete key
            {

                e.Handled = true;

            }
        }

        private void txtECTS_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46) // 8 is a backspace key 56 is a delete key
            {

                e.Handled = true;

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int firsttextvalue;
            int secondtextvalue;
            int atmost14;
            int number = 16;
            int thirdtextvalue;
            int fourthtextvalue;

            wsem1.Text = total1.Text;
            wsem2.Text = total2.Text;
            wendsem1.Text = total1.Text;
            wendsem2.Text = total2.Text;
            

            
            atmost14 = int.Parse(nu3.Text);
            if (atmost14 > 14)
            {
                MessageBox.Show("Study Hours out Class can be at most 14");
            }

            firsttextvalue = int.Parse(txtTheory.Text);
            secondtextvalue = int.Parse(txtApplication.Text) / 2;
              txtLocalCredits.Text = (firsttextvalue + secondtextvalue).ToString();

            thirdtextvalue = int.Parse(d1.Text);
           fourthtextvalue = int.Parse(d2.Text);


            thirdtextvalue = firsttextvalue;
            fourthtextvalue = firsttextvalue;



            wo1.Text = (firsttextvalue * number).ToString();
            wo2.Text = (firsttextvalue * number).ToString();



            CreateWordDocument(@"C:\Users\BARTH\OneDrive\Masaüstü\syllabus.docx", @"C:\Users\BARTH\OneDrive\Masaüstü\newcreatedsyllabs.docx");




        }

        private void txtLocalCredits_TextChanged(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtCourseObjectives_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void checkEnglish_CheckedChanged(object sender, EventArgs e)
        {
            labeleng.Text = checkEnglish.Checked ? "X English" : "English";


        }

        private void checkTurkish_CheckedChanged(object sender, EventArgs e)
        {
            labelturk.Text = checkTurkish.Checked ? "X Turkish" : "Turkish";
        }

        private void checkSecond_CheckedChanged(object sender, EventArgs e)
        {
            labelsecond.Text = checkSecond.Checked ? "X Second Foreign Language" : "Second Foreign Language";

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void checkShortCycle_CheckedChanged(object sender, EventArgs e)
        {
            labelshort.Text = checkShortCycle.Checked ? "X Short Cycle" : "Short Cycle";

        }

        private void labelrequired_Click(object sender, EventArgs e)
        {

        }

        private void labelelective_Click(object sender, EventArgs e)
        {

        }

        private void checkRequired_CheckedChanged(object sender, EventArgs e)
        {
            labelrequired.Text = checkRequired.Checked ? "X Required" : "Required";
        }

        private void checkElective_CheckedChanged(object sender, EventArgs e)
        {
            labelelective.Text = checkElective.Checked ? "X Elective" : "Elective";
        }

        private void checkFirstCycle_CheckedChanged(object sender, EventArgs e)
        {
            labelfirst.Text = checkFirstCycle.Checked ? "X First Cycle" : "First Cycle";
        }

        private void checkSecondCycle_CheckedChanged(object sender, EventArgs e)
        {
            labelscnd.Text = checkSecondCycle.Checked ? "X Second Cycle" : "Second Cycle";
        }

        private void checkThirdCycle_CheckedChanged(object sender, EventArgs e)
        {
            labelthird.Text = checkThirdCycle.Checked ? "X Third Cycle" : "Third Cycle";
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label20_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkcore_CheckedChanged(object sender, EventArgs e)
        {
            labelcore.Text = checkcore.Checked ? "X" : " ";
        }

        private void checkmajor_CheckedChanged(object sender, EventArgs e)
        {
            labelmajor.Text = checkmajor.Checked ? "X" : " ";
        }

        private void checksupportive_CheckedChanged(object sender, EventArgs e)
        {
            labelsupportive.Text = checksupportive.Checked ? "X" : " ";
        }

        private void checkcommunication_CheckedChanged(object sender, EventArgs e)
        {
            labelcommunication.Text = checkcommunication.Checked ? "X" : " ";
        }

        private void checktransferable_CheckedChanged(object sender, EventArgs e)
        {
            labeltransferable.Text = checktransferable.Checked ? "X" : " ";
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtCourseCoordinator_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTheory_TextChanged(object sender, EventArgs e)
        {

        }

        private void label82_Click(object sender, EventArgs e)
        {

        }

        private void label84_Click(object sender, EventArgs e)
        {

        }

        private void txtECTS_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
