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
using System.Xml.Serialization;


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
                this.FindAndReplace(wordApp, "<Ects>", txtECTS.Text);

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



                this.FindAndReplace(wordApp, "<lo11>", txtlo11.Text);
                this.FindAndReplace(wordApp, "<lo12>", txtlo12.Text);
                this.FindAndReplace(wordApp, "<lo13>", txtlo13.Text);
                this.FindAndReplace(wordApp, "<lo14>", txtlo14.Text);
                this.FindAndReplace(wordApp, "<lo15>", txtlo15.Text);
                this.FindAndReplace(wordApp, "<lo16>", txtlo16.Text);
                this.FindAndReplace(wordApp, "<lo17>", txtlo17.Text);
                this.FindAndReplace(wordApp, "<lo18>", txtlo18.Text);
                this.FindAndReplace(wordApp, "<lo19>", txtlo19.Text);
                this.FindAndReplace(wordApp, "<lo110>", txtlo110.Text);
                this.FindAndReplace(wordApp, "<lo111>", txtlo111.Text);
                this.FindAndReplace(wordApp, "<totallo1>", totallo1.Text);



                this.FindAndReplace(wordApp, "<lo21>", txtlo21.Text);
                this.FindAndReplace(wordApp, "<lo22>", txtlo22.Text);
                this.FindAndReplace(wordApp, "<lo23>", txtlo23.Text);
                this.FindAndReplace(wordApp, "<lo24>", txtlo24.Text);
                this.FindAndReplace(wordApp, "<lo25>", txtlo25.Text);
                this.FindAndReplace(wordApp, "<lo26>", txtlo26.Text);
                this.FindAndReplace(wordApp, "<lo27>", txtlo27.Text);
                this.FindAndReplace(wordApp, "<lo28>", txtlo28.Text);
                this.FindAndReplace(wordApp, "<lo29>", txtlo29.Text);
                this.FindAndReplace(wordApp, "<lo210>", txtlo210.Text);
                this.FindAndReplace(wordApp, "<lo211>", txtlo211.Text);
                this.FindAndReplace(wordApp, "<totallo2>", totallo2.Text);

                this.FindAndReplace(wordApp, "<lo31>", txtlo31.Text);
                this.FindAndReplace(wordApp, "<lo32>", txtlo32.Text);
                this.FindAndReplace(wordApp, "<lo33>", txtlo33.Text);
                this.FindAndReplace(wordApp, "<lo34>", txtlo34.Text);
                this.FindAndReplace(wordApp, "<lo35>", txtlo35.Text);
                this.FindAndReplace(wordApp, "<lo36>", txtlo36.Text);
                this.FindAndReplace(wordApp, "<lo37>", txtlo37.Text);
                this.FindAndReplace(wordApp, "<lo38>", txtlo38.Text);
                this.FindAndReplace(wordApp, "<lo39>", txtlo39.Text);
                this.FindAndReplace(wordApp, "<lo310>", txtlo310.Text);
                this.FindAndReplace(wordApp, "<lo311>", txtlo311.Text);
                this.FindAndReplace(wordApp, "<totallo3>", totallo3.Text);

                this.FindAndReplace(wordApp, "<lo41>", txtlo41.Text);
                this.FindAndReplace(wordApp, "<lo42>", txtlo42.Text);
                this.FindAndReplace(wordApp, "<lo43>", txtlo43.Text);
                this.FindAndReplace(wordApp, "<lo44>", txtlo44.Text);
                this.FindAndReplace(wordApp, "<lo45>", txtlo45.Text);
                this.FindAndReplace(wordApp, "<lo46>", txtlo46.Text);
                this.FindAndReplace(wordApp, "<lo47>", txtlo47.Text);
                this.FindAndReplace(wordApp, "<lo48>", txtlo48.Text);
                this.FindAndReplace(wordApp, "<lo49>", txtlo49.Text);
                this.FindAndReplace(wordApp, "<lo410>", txtlo410.Text);
                this.FindAndReplace(wordApp, "<lo411>", txtlo411.Text);
                this.FindAndReplace(wordApp, "<totallo4>", totallo4.Text);


                this.FindAndReplace(wordApp, "<lo51>", txtlo51.Text);
                this.FindAndReplace(wordApp, "<lo52>", txtlo52.Text);
                this.FindAndReplace(wordApp, "<lo53>", txtlo53.Text);
                this.FindAndReplace(wordApp, "<lo54>", txtlo54.Text);
                this.FindAndReplace(wordApp, "<lo55>", txtlo55.Text);
                this.FindAndReplace(wordApp, "<lo56>", txtlo56.Text);
                this.FindAndReplace(wordApp, "<lo57>", txtlo57.Text);
                this.FindAndReplace(wordApp, "<lo58>", txtlo58.Text);
                this.FindAndReplace(wordApp, "<lo59>", txtlo59.Text);
                this.FindAndReplace(wordApp, "<lo510>", txtlo510.Text);
                this.FindAndReplace(wordApp, "<lo511>", txtlo511.Text);
                this.FindAndReplace(wordApp, "<totallo5>", totallo5.Text);

                this.FindAndReplace(wordApp, "<lo61>", txtlo61.Text);
                this.FindAndReplace(wordApp, "<lo62>", txtlo62.Text);
                this.FindAndReplace(wordApp, "<lo63>", txtlo63.Text);
                this.FindAndReplace(wordApp, "<lo64>", txtlo64.Text);
                this.FindAndReplace(wordApp, "<lo65>", txtlo65.Text);
                this.FindAndReplace(wordApp, "<lo66>", txtlo66.Text);
                this.FindAndReplace(wordApp, "<lo67>", txtlo67.Text);
                this.FindAndReplace(wordApp, "<lo68>", txtlo68.Text);
                this.FindAndReplace(wordApp, "<lo69>", txtlo69.Text);
                this.FindAndReplace(wordApp, "<lo610>", txtlo610.Text);
                this.FindAndReplace(wordApp, "<lo611>", txtlo611.Text);
                this.FindAndReplace(wordApp, "<totallo6>", totallo6.Text);

                this.FindAndReplace(wordApp, "<lo71>", txtlo71.Text);
                this.FindAndReplace(wordApp, "<lo72>", txtlo72.Text);
                this.FindAndReplace(wordApp, "<lo73>", txtlo73.Text);
                this.FindAndReplace(wordApp, "<lo74>", txtlo74.Text);
                this.FindAndReplace(wordApp, "<lo75>", txtlo75.Text);
                this.FindAndReplace(wordApp, "<lo76>", txtlo76.Text);
                this.FindAndReplace(wordApp, "<lo77>", txtlo77.Text);
                this.FindAndReplace(wordApp, "<lo78>", txtlo78.Text);
                this.FindAndReplace(wordApp, "<lo79>", txtlo79.Text);
                this.FindAndReplace(wordApp, "<lo710>", txtlo710.Text);
                this.FindAndReplace(wordApp, "<lo711>", txtlo711.Text);
                this.FindAndReplace(wordApp, "<totallo7>", totallo7.Text);







                this.FindAndReplace(wordApp, "<wsem1>", wsem1.Text);
                this.FindAndReplace(wordApp, "<wsem2>", wsem2.Text);
                this.FindAndReplace(wordApp, "<wendsem1>", wendsem1.Text);
                this.FindAndReplace(wordApp, "<wendsem2>", wendsem2.Text);


                this.FindAndReplace(wordApp, "<totalgrade1>", total1.Text);
                this.FindAndReplace(wordApp, "<totalgrade2>", total2.Text);


                this.FindAndReplace(wordApp, "<nu1>", nu1.Text);
                this.FindAndReplace(wordApp, "<nu2>", nu2.Text);
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
                this.FindAndReplace(wordApp, "<d11>", d11.Text);
                this.FindAndReplace(wordApp, "<d12>", d12.Text);

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
                this.FindAndReplace(wordApp, "<wo11>", wo11.Text);
                this.FindAndReplace(wordApp, "<wo12>", wo12.Text);

                this.FindAndReplace(wordApp, "<totalwo>", totalwo.Text);



                this.FindAndReplace(wordApp, "<prg1>", prg1.Text);
                this.FindAndReplace(wordApp, "<prg2>", prg2.Text);
                this.FindAndReplace(wordApp, "<prg3>", prg3.Text);
                this.FindAndReplace(wordApp, "<prg4>", prg4.Text);
                this.FindAndReplace(wordApp, "<prg5>", prg5.Text);
                this.FindAndReplace(wordApp, "<prg6>", prg6.Text);
                this.FindAndReplace(wordApp, "<prg7>", prg7.Text);
                this.FindAndReplace(wordApp, "<prg8>", prg8.Text);
                this.FindAndReplace(wordApp, "<prg9>", prg9.Text);
                this.FindAndReplace(wordApp, "<prg10>", prg10.Text);
                this.FindAndReplace(wordApp, "<prg11>", prg11.Text);
                



                this.FindAndReplace(wordApp, "<cfirst1>", cfirst1.Text);
                this.FindAndReplace(wordApp, "<cfirst2>", cfirst2.Text);
                this.FindAndReplace(wordApp, "<cfirst3>", cfirst3.Text);
                this.FindAndReplace(wordApp, "<cfirst4>", cfirst4.Text);
                this.FindAndReplace(wordApp, "<cfirst5>", cfirst5.Text);
                this.FindAndReplace(wordApp, "<cfirst6>", cfirst6.Text);
                this.FindAndReplace(wordApp, "<cfirst7>", cfirst7.Text);
                this.FindAndReplace(wordApp, "<cfirst8>", cfirst8.Text);
                this.FindAndReplace(wordApp, "<cfirst9>", cfirst9.Text);
                this.FindAndReplace(wordApp, "<cfirst10>", cfirst10.Text);
                this.FindAndReplace(wordApp, "<cfirst11>", cfirst11.Text);
                this.FindAndReplace(wordApp, "<cfirst12>", cfirst12.Text);
                this.FindAndReplace(wordApp, "<cfirst13>", cfirst13.Text);


                this.FindAndReplace(wordApp, "<csecond1>", csecond1.Text);
                this.FindAndReplace(wordApp, "<csecond2>", csecond2.Text);
                this.FindAndReplace(wordApp, "<csecond3>", csecond3.Text);
                this.FindAndReplace(wordApp, "<csecond4>", csecond4.Text);
                this.FindAndReplace(wordApp, "<csecond5>", csecond5.Text);
                this.FindAndReplace(wordApp, "<csecond6>", csecond6.Text);
                this.FindAndReplace(wordApp, "<csecond7>", csecond7.Text);
                this.FindAndReplace(wordApp, "<csecond8>", csecond8.Text);
                this.FindAndReplace(wordApp, "<csecond9>", csecond9.Text);
                this.FindAndReplace(wordApp, "<csecond10>", csecond10.Text);
                this.FindAndReplace(wordApp, "<csecond11>", csecond11.Text);
                this.FindAndReplace(wordApp, "<csecond12>", csecond12.Text);
                this.FindAndReplace(wordApp, "<csecond13>", csecond13.Text);

                this.FindAndReplace(wordApp, "<cthird1>", cthird1.Text);
                this.FindAndReplace(wordApp, "<cthird2>", cthird2.Text);
                this.FindAndReplace(wordApp, "<cthird3>", cthird3.Text);
                this.FindAndReplace(wordApp, "<cthird4>", cthird4.Text);
                this.FindAndReplace(wordApp, "<cthird5>", cthird5.Text);
                this.FindAndReplace(wordApp, "<cthird6>", cthird6.Text);
                this.FindAndReplace(wordApp, "<cthird7>", cthird7.Text);
                this.FindAndReplace(wordApp, "<cthird8>", cthird8.Text);
                this.FindAndReplace(wordApp, "<cthird9>", cthird9.Text);
                this.FindAndReplace(wordApp, "<cthird10>", cthird10.Text);
                this.FindAndReplace(wordApp, "<cthird11>", cthird11.Text);
                this.FindAndReplace(wordApp, "<cthird12>", cthird12.Text);
                this.FindAndReplace(wordApp, "<cthird13>", cthird13.Text);

                this.FindAndReplace(wordApp, "<cforth1>", cforth1.Text);
                this.FindAndReplace(wordApp, "<cforth2>", cforth2.Text);
                this.FindAndReplace(wordApp, "<cforth3>", cforth3.Text);
                this.FindAndReplace(wordApp, "<cforth4>", cforth4.Text);
                this.FindAndReplace(wordApp, "<cforth5>", cforth5.Text);
                this.FindAndReplace(wordApp, "<cforth6>", cforth6.Text);
                this.FindAndReplace(wordApp, "<cforth7>", cforth7.Text);
                this.FindAndReplace(wordApp, "<cforth8>", cforth8.Text);
                this.FindAndReplace(wordApp, "<cforth9>", cforth9.Text);
                this.FindAndReplace(wordApp, "<cforth10>", cforth10.Text);
                this.FindAndReplace(wordApp, "<cforth11>", cforth11.Text);
                this.FindAndReplace(wordApp, "<cforth12>", cforth12.Text);
                this.FindAndReplace(wordApp, "<cforth13>", cforth13.Text);


                this.FindAndReplace(wordApp, "<cfifth1>", cfifth1.Text);
                this.FindAndReplace(wordApp, "<cfifth2>", cfifth2.Text);
                this.FindAndReplace(wordApp, "<cfifth3>", cfifth3.Text);
                this.FindAndReplace(wordApp, "<cfifth4>", cfifth4.Text);
                this.FindAndReplace(wordApp, "<cfifth5>", cfifth5.Text);
                this.FindAndReplace(wordApp, "<cfifth6>", cfifth6.Text);
                this.FindAndReplace(wordApp, "<cfifth7>", cfifth7.Text);
                this.FindAndReplace(wordApp, "<cfifth8>", cfifth8.Text);
                this.FindAndReplace(wordApp, "<cfifth9>", cfifth9.Text);
                this.FindAndReplace(wordApp, "<cfifth10>", cfifth10.Text);
                this.FindAndReplace(wordApp, "<cfifth11>", cfifth11.Text);
                this.FindAndReplace(wordApp, "<cfifth12>", cfifth12.Text);
                this.FindAndReplace(wordApp, "<cfifth13>", cfifth13.Text);


                this.FindAndReplace(wordApp, "<LO#1>", lo_1.Text);
                this.FindAndReplace(wordApp, "<LO#2>", lo_2.Text);
                this.FindAndReplace(wordApp, "<LO#3>", lo_3.Text);
                this.FindAndReplace(wordApp, "<LO#4>", lo_4.Text);
                this.FindAndReplace(wordApp, "<LO#5>", lo_5.Text);
                this.FindAndReplace(wordApp, "<LO#6>", lo_6.Text);
                this.FindAndReplace(wordApp, "<LO#7>", lo_7.Text);
                this.FindAndReplace(wordApp, "<LO#8>", lo_8.Text);
                this.FindAndReplace(wordApp, "<LO#9>", lo_9.Text);
                this.FindAndReplace(wordApp, "<LO#10>", lo_10.Text);
                this.FindAndReplace(wordApp, "<LO#11>", lo_11.Text);
                this.FindAndReplace(wordApp, "<LO#12>", lo_12.Text);
                this.FindAndReplace(wordApp, "<LO#13>", lo_13.Text);





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



            firsttextvalue = int.Parse(txtTheory.Text);
            secondtextvalue = int.Parse(txtApplication.Text) / 2;
            txtLocalCredits.Text = (firsttextvalue + secondtextvalue).ToString();











             
            



            /*  atmost14 = int.Parse(nu3.Text);
              if (atmost14 > 14)
              {
                  MessageBox.Show("Study Hours out Class can be at most 14");
              }
              thirdtextvalue = int.Parse(d1.Text);
              fourthtextvalue = int.Parse(d2.Text);


              thirdtextvalue = firsttextvalue;
              fourthtextvalue = firsttextvalue; */


            d1.Text = firsttextvalue.ToString();
            d2.Text = firsttextvalue.ToString();
            wo1.Text = (firsttextvalue * number).ToString();
            wo2.Text = (firsttextvalue * number).ToString();




            int n1_,n2_,n3_,n4_,n5_,n6_,n7_,n8_,n9_,n10_,n11_;
            int w1_, w2_, w3_, w4_, w5_, w6_, w7_, w8_, w9_, w10_, w11_;
            int lo11_, lo12_, lo13_, lo14_, lo15_, lo16_, lo17_, lo18_, lo19_, lo110_, lo111_;
            int lo21_, lo22_, lo23_, lo24_, lo25_, lo26_, lo27_, lo28_, lo29_, lo210_, lo211_;
            int lo31_, lo32_, lo33_, lo34_, lo35_, lo36_, lo37_, lo38_, lo39_, lo310_, lo311_;
            int lo41_, lo42_, lo43_, lo44_, lo45_, lo46_, lo47_, lo48_, lo49_, lo410_, lo411_;
            int lo51_, lo52_, lo53_, lo54_, lo55_, lo56_, lo57_, lo58_, lo59_, lo510_, lo511_;
            int lo61_, lo62_, lo63_, lo64_, lo65_, lo66_, lo67_, lo68_, lo69_, lo610_, lo611_;
            int lo71_, lo72_, lo73_, lo74_, lo75_, lo76_, lo77_, lo78_, lo79_, lo710_, lo711_;
            int totalnumber, totalweigthing, totallo1_, totallo2_, totallo3_, totallo4_, totallo5_, totallo6_, totallo7_;
            int wsem1_, wsem2_, wendsem1_, wendsem2_;
            int totalgrade1_, totalgrade2_;
            int wo1_, wo2_, wo3_, wo4_, wo5_, wo6_, wo7_, wo8_, wo9_, wo10_, wo11_, wo12_;
            int totalwo_;
            int ECTS;
            







            wo1_ = int.Parse(wo1.Text);
            wo2_ = int.Parse(wo2.Text);
            wo3_ = int.Parse(wo3.Text);
            wo4_ = int.Parse(wo4.Text);
            wo5_ = int.Parse(wo5.Text);
            wo6_ = int.Parse(wo6.Text);
            wo7_ = int.Parse(wo7.Text);
            wo8_ = int.Parse(wo8.Text);
            wo9_ = int.Parse(wo9.Text);
            wo10_ = int.Parse(wo10.Text);
            wo11_ = int.Parse(wo11.Text);
            wo12_ = int.Parse(wo12.Text);

            totalwo_ = wo1_ + wo2_ + wo3_ + wo4_ + wo5_ + wo6_ + wo7_ + wo8_ + wo9_ + wo10_ + wo11_ + wo12_;
            totalwo.Text = totalwo_.ToString();

            


            txtECTS.Text = (totalwo_/30).ToString();




            n1_ = int.Parse(n1.Text);
            n2_ = int.Parse(n2.Text);
            n3_ = int.Parse(n3.Text);
            n4_ = int.Parse(n4.Text);
            n5_ = int.Parse(n5.Text);
            n6_ = int.Parse(n6.Text);
            n7_ = int.Parse(n7.Text);
            n8_ = int.Parse(n8.Text);
            n9_ = int.Parse(n9.Text);
            n10_ = int.Parse(n10.Text);
            n11_ = int.Parse(n11.Text);

            totalnumber = n1_ + n2_ + n3_ + n4_ + n5_ + n6_ + n7_ + n8_ + n9_ + n10_ + n11_;
            n12.Text = totalnumber.ToString();


             w1_ = int.Parse(w1.Text);
            w2_ = int.Parse(w2.Text);
            w3_ = int.Parse(w3.Text);
            w4_ = int.Parse(w4.Text);
            w5_ = int.Parse(w5.Text);
            w6_ = int.Parse(w6.Text);
            w7_ = int.Parse(w7.Text);
            w8_ = int.Parse(w8.Text);
            w9_ = int.Parse(w9.Text);
            w10_ = int.Parse(w10.Text);
            w11_ = int.Parse(w11.Text);

            totalweigthing = w1_ + w2_ + w3_ + w4_ + w5_ + w6_ + w7_ + w8_ + w9_ + w10_ + w11_;
            w12.Text = totalweigthing.ToString();

            lo11_ = int.Parse(txtlo11.Text);
            lo12_ = int.Parse(txtlo12.Text);
            lo13_ = int.Parse(txtlo13.Text);
            lo14_ = int.Parse(txtlo14.Text);
            lo15_ = int.Parse(txtlo15.Text);
            lo16_ = int.Parse(txtlo16.Text);
            lo17_ = int.Parse(txtlo17.Text);
            lo18_ = int.Parse(txtlo18.Text);
            lo19_ = int.Parse(txtlo19.Text);
            lo110_ = int.Parse(txtlo110.Text);
            lo111_ = int.Parse(txtlo110.Text);

            totallo1_ = lo11_+ lo12_ + lo13_ + lo14_ + lo15_ + lo16_ + lo17_ + lo18_ + lo19_ + lo110_ + lo111_;

            totallo1.Text = totallo1_.ToString();

            lo21_ = int.Parse(txtlo21.Text);
            lo22_ = int.Parse(txtlo22.Text);
            lo23_ = int.Parse(txtlo23.Text);
            lo24_ = int.Parse(txtlo24.Text);
            lo25_ = int.Parse(txtlo25.Text);
            lo26_ = int.Parse(txtlo26.Text);
            lo27_ = int.Parse(txtlo27.Text);
            lo28_ = int.Parse(txtlo28.Text);
            lo29_ = int.Parse(txtlo29.Text);
            lo210_ = int.Parse(txtlo210.Text);
            lo211_ = int.Parse(txtlo211.Text);


            totallo2_ = lo21_ + lo22_ + lo23_ + lo24_ + lo25_ + lo26_ + lo27_ + lo28_ + lo29_ + lo210_ + lo211_;

            totallo2.Text = totallo2_.ToString();

            lo31_ = int.Parse(txtlo31.Text);
            lo32_ = int.Parse(txtlo32.Text);
            lo33_ = int.Parse(txtlo33.Text);
            lo34_ = int.Parse(txtlo34.Text);
            lo35_ = int.Parse(txtlo35.Text);
            lo36_ = int.Parse(txtlo36.Text);
            lo37_ = int.Parse(txtlo37.Text);
            lo38_ = int.Parse(txtlo38.Text);
            lo39_ = int.Parse(txtlo39.Text);
            lo310_ = int.Parse(txtlo310.Text);
            lo311_ = int.Parse(txtlo310.Text);

            totallo3_ = lo31_ + lo32_ + lo33_ + lo34_ + lo35_ + lo36_ + lo37_ + lo38_ + lo39_ + lo310_ + lo311_;

            totallo3.Text = totallo3_.ToString();

            lo41_ = int.Parse(txtlo41.Text);
            lo42_ = int.Parse(txtlo42.Text);
            lo43_ = int.Parse(txtlo43.Text);
            lo44_ = int.Parse(txtlo44.Text);
            lo45_ = int.Parse(txtlo45.Text);
            lo46_ = int.Parse(txtlo46.Text);
            lo47_ = int.Parse(txtlo47.Text);
            lo48_ = int.Parse(txtlo48.Text);
            lo49_ = int.Parse(txtlo49.Text);
            lo410_ = int.Parse(txtlo410.Text);
            lo411_ = int.Parse(txtlo411.Text);

            totallo4_ = lo41_ + lo42_ + lo43_ + lo44_ + lo45_ + lo46_ + lo47_ + lo48_ + lo49_ + lo410_ + lo411_;
            totallo4.Text = totallo4_.ToString();

            lo51_ = int.Parse(txtlo51.Text);
            lo52_ = int.Parse(txtlo52.Text);
            lo53_ = int.Parse(txtlo53.Text);
            lo54_ = int.Parse(txtlo54.Text);
            lo55_ = int.Parse(txtlo55.Text);
            lo56_ = int.Parse(txtlo56.Text);
            lo57_ = int.Parse(txtlo57.Text);
            lo58_ = int.Parse(txtlo58.Text);
            lo59_ = int.Parse(txtlo59.Text);
            lo510_ = int.Parse(txtlo510.Text);
            lo511_ = int.Parse(txtlo510.Text);

            totallo5_ = lo51_ + lo52_ + lo53_ + lo54_ + lo55_ + lo56_ + lo57_ + lo58_ + lo59_ + lo510_ + lo511_;
            totallo5.Text = totallo5_.ToString();

            lo61_ = int.Parse(txtlo61.Text);
            lo62_ = int.Parse(txtlo62.Text);
            lo63_ = int.Parse(txtlo63.Text);
            lo64_ = int.Parse(txtlo64.Text);
            lo65_ = int.Parse(txtlo65.Text);
            lo66_ = int.Parse(txtlo66.Text);
            lo67_ = int.Parse(txtlo67.Text);
            lo68_ = int.Parse(txtlo68.Text);
            lo69_ = int.Parse(txtlo69.Text);
            lo610_ = int.Parse(txtlo610.Text);
            lo611_ = int.Parse(txtlo610.Text);

            totallo6_ = lo61_ + lo62_ + lo63_ + lo64_ + lo65_ + lo66_ + lo67_ + lo68_ + lo69_ + lo610_ + lo611_;

            totallo6.Text = totallo6_.ToString();
            lo71_ = int.Parse(txtlo71.Text);
            lo72_ = int.Parse(txtlo72.Text);
            lo73_ = int.Parse(txtlo73.Text);
            lo74_ = int.Parse(txtlo74.Text);
            lo75_ = int.Parse(txtlo75.Text);
            lo76_ = int.Parse(txtlo76.Text);
            lo77_ = int.Parse(txtlo77.Text);
            lo78_ = int.Parse(txtlo78.Text);
            lo79_ = int.Parse(txtlo79.Text);
            lo710_ = int.Parse(txtlo710.Text);
            lo711_ = int.Parse(txtlo710.Text);

            totallo7_ = lo71_ + lo72_ + lo73_ + lo74_ + lo75_ + lo76_ + lo77_ + lo78_ + lo79_ + lo710_ + lo711_;

            totallo7.Text = totallo7_.ToString();


            wsem1_ = int.Parse(wsem1.Text);
            wsem2_ = int.Parse(wsem2.Text);
            wendsem1_ = int.Parse(wendsem1.Text);
            wendsem2_ = int.Parse(wendsem2.Text);

            totalgrade1_ = wsem1_ + wendsem1_;
            totalgrade2_ = wsem2_ + wendsem2_;

            total1.Text = totalgrade1_.ToString();

            total2.Text = totalgrade2_.ToString();


            




            CreateWordDocument(@"C:\Users\BARTH\OneDrive\Masaüstü\syllabus12.docx", @"C:\Users\BARTH\OneDrive\Masaüstü\newcreatedsyllabs.docx");




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

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46) // 8 is a backspace key 56 is a delete key
            {

                e.Handled = true;

            }
        }

        private void btn_Click(object sender, EventArgs e)
        {
         

            

        }

        private void tableLayoutPanel13_Paint(object sender, PaintEventArgs e)
        {
 cthird6.Text = cthird6.Checked ? "X" : " ";
        }

        private void cfirst1_CheckedChanged(object sender, EventArgs e)
        {
            cfirst1.Text = cfirst1.Checked ? "X" : " ";
        }

        private void cfirst2_CheckedChanged(object sender, EventArgs e)
        {
            cfirst2.Text = cfirst2.Checked ? "X" : " ";
        }

        private void cfirst3_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst4.Text = cfirst4.Checked ? "X" : " ";
        }

        private void cfirst4_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst4.Text = cfirst4.Checked ? "X" : " ";
        }

        private void cfirst5_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst5.Text = cfirst5.Checked ? "X" : " ";
        }

        private void cfirst6_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst6.Text = cfirst6.Checked ? "X" : " ";
        }

        private void cfirst7_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst7.Text = cfirst7.Checked ? "X" : " ";
        }
        private void cfirst8_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst8.Text = cfirst8.Checked ? "X" : " ";
        }
        private void cfirst9_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst9.Text = cfirst9.Checked ? "X" : " ";
        }
        private void cfirst10_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst10.Text = cfirst10.Checked ? "X" : " ";
        }
        private void cfirst11_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst11.Text = cfirst11.Checked ? "X" : " ";
        }
        private void cfirst12_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst12.Text = cfirst12.Checked ? "X" : " ";
        }
        private void cfirst13_CheckedChanged_1(object sender, EventArgs e)
        {
            cfirst13.Text = cfirst13.Checked ? "X" : " ";
        }

        private void csecond1_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond1.Text = csecond1.Checked ? "X" : " ";
        }
        private void csecond2_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond2.Text = csecond2.Checked ? "X" : " ";
        }
        private void csecond3_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond3.Text = csecond3.Checked ? "X" : " ";
        }
        private void csecond4_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond4.Text = csecond4.Checked ? "X" : " ";
        }
        private void csecond5_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond5.Text = csecond5.Checked ? "X" : " ";
        }
        private void csecond6_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond6.Text = csecond6.Checked ? "X" : " ";
        }
        private void csecond7_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond7.Text = csecond7.Checked ? "X" : " ";
        }
        private void csecond8_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond8.Text = csecond8.Checked ? "X" : " ";
        }
        private void csecond9_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond9.Text = csecond9.Checked ? "X" : " ";
        }
        private void csecond10_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond10.Text = csecond10.Checked ? "X" : " ";
        }
        private void csecond11_CheckedChanged_1(object sender, EventArgs e)
        {
            csecond11.Text = csecond11.Checked ? "X" : " ";
        }
        private void csecond12_CheckedChanged(object sender, EventArgs e)
        {
            csecond12.Text = csecond12.Checked ? "X" : " ";
        }
        private void csecond13_CheckedChanged(object sender, EventArgs e)
        {
            csecond13.Text = csecond13.Checked ? "X" : " ";
        }

        private void cthird1_CheckedChanged(object sender, EventArgs e)
        {
            cthird1.Text = cthird1.Checked ? "X" : " ";
        }
        private void cthird2_CheckedChanged(object sender, EventArgs e)
        {
            cthird2.Text = cthird2.Checked ? "X" : " ";
        }
        private void cthird3_CheckedChanged(object sender, EventArgs e)
        {
            cthird3.Text = cthird3.Checked ? "X" : " ";
        }
        private void cthird4_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird4.Text = cthird4.Checked ? "X" : " ";
        }
        private void cthird5_CheckedChanged(object sender, EventArgs e)
        {
            cthird5.Text = cthird5.Checked ? "X" : " ";
        }
        private void cthird6_CheckedChanged(object sender, EventArgs e)
        {
            cthird6.Text = cthird6.Checked ? "X" : " ";
        }
        private void cthird7_CheckedChanged(object sender, EventArgs e)
        {
            cthird7.Text = cthird7.Checked ? "X" : " ";
        }
        private void cthird8_CheckedChanged(object sender, EventArgs e)
        {
            cthird8.Text = cthird8.Checked ? "X" : " ";
        }
        private void cthird9_CheckedChanged(object sender, EventArgs e)
        {
            cthird9.Text = cthird9.Checked ? "X" : " ";
        }
        private void cthird10_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird10.Text = cthird10.Checked ? "X" : " ";
        }
        private void cthird11_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird11.Text = cthird11.Checked ? "X" : " ";
        }
        private void cthird12_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird12.Text = cthird12.Checked ? "X" : " ";
        }
        private void cthird13_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird13.Text = cthird13.Checked ? "X" : " ";
        }





        private void cforth1_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth1.Text = cforth1.Checked ? "X" : " ";
        }
        private void cforth2_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth2.Text = cforth2.Checked ? "X" : " ";
        }
        private void cforth3_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth3.Text = cforth3.Checked ? "X" : " ";
        }
        private void cforth4_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth4.Text = cforth4.Checked ? "X" : " ";
        }
        private void cforth5_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth5.Text = cforth5.Checked ? "X" : " ";
        }
        private void cforth6_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth6.Text = cforth6.Checked ? "X" : " ";
        }
        private void cforth7_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth7.Text = cforth7.Checked ? "X" : " ";
        }
        private void cforth8_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth8.Text = cforth8.Checked ? "X" : " ";
        }
        private void cforth9_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth9.Text = cforth9.Checked ? "X" : " ";
        }
        private void cforth10_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth10.Text = cforth10.Checked ? "X" : " ";
        }
        private void cforth11_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth11.Text = cforth11.Checked ? "X" : " ";
        }
        private void cforth12_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth12.Text = cforth12.Checked ? "X" : " ";
        }
        private void cforth13_CheckedChanged_1(object sender, EventArgs e)
        {
            cforth13.Text = cforth13.Checked ? "X" : " ";
        }

        private void cfifth1_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth1.Text = cfifth1.Checked ? "X" : " ";
        }
        private void cfifth2_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth2.Text = cfifth2.Checked ? "X" : " ";
        }
        private void cfifth3_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth3.Text = cfifth3.Checked ? "X" : " ";
        }
        private void cfifth4_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth4.Text = cfifth4.Checked ? "X" : " ";
        }
        private void cfifth5_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth5.Text = cfifth5.Checked ? "X" : " ";
        }
        private void cfifth6_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth6.Text = cfifth6.Checked ? "X" : " ";
        }
        private void cfifth7_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth7.Text = cfifth7.Checked ? "X" : " ";
        }
        private void cfifth8_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth8.Text = cfifth8.Checked ? "X" : " ";
        }
        private void cfifth9_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth9.Text = cfifth9.Checked ? "X" : " ";
        }
        private void cfifth10_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth10.Text = cfifth10.Checked ? "X" : " ";
        }
        private void cfifth11_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth11.Text = cfifth11.Checked ? "X" : " ";
        }
        private void cfifth12_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth12.Text = cfifth12.Checked ? "X" : " ";
        }
        private void cfifth13_CheckedChanged_1(object sender, EventArgs e)
        {
            cfifth13.Text = cfifth13.Checked ? "X" : " ";
        }

        private void cthird4_CheckedChanged(object sender, EventArgs e)
        {
            cthird4.Text = cthird4.Checked ? "X" : " ";
        }

        private void cfifth5_CheckedChanged(object sender, EventArgs e)
        {
            cfifth5.Text = cfifth5.Checked ? "X" : " ";
        }

       

        private void cfirst3_CheckedChanged(object sender, EventArgs e)
        {
            cfirst3.Text = cfirst3.Checked ? "X" : " ";
        }

        private void cfirst4_CheckedChanged(object sender, EventArgs e)
        {
            cfirst4.Text = cfirst4.Checked ? "X" : " ";
        }

        private void cfirst5_CheckedChanged(object sender, EventArgs e)
        {
            cfirst5.Text = cfirst5.Checked ? "X" : " ";
        }

        private void cfirst6_CheckedChanged(object sender, EventArgs e)
        {
            cfirst6.Text = cfirst6.Checked ? "X" : " ";
        }

        private void cfirst7_CheckedChanged(object sender, EventArgs e)
        {
            cfirst7.Text = cfirst7.Checked ? "X" : " ";
        }

        private void cfirst8_CheckedChanged(object sender, EventArgs e)
        {
            cfirst8.Text = cfirst8.Checked ? "X" : " ";
        }

        private void cfirst9_CheckedChanged(object sender, EventArgs e)
        {
            cfirst9.Text = cfirst9.Checked ? "X" : " ";
        }

        private void cfirst10_CheckedChanged(object sender, EventArgs e)
        {
            cfirst10.Text = cfirst10.Checked ? "X" : " ";
        }

        private void cfirst11_CheckedChanged(object sender, EventArgs e)
        {
            cfirst11.Text = cfirst11.Checked ? "X" : " ";
        }

        private void cfirst12_CheckedChanged(object sender, EventArgs e)
        {
            cfirst12.Text = cfirst12.Checked ? "X" : " ";
        }

        private void cfirst13_CheckedChanged(object sender, EventArgs e)
        {
            cfirst13.Text = cfirst13.Checked ? "X" : " ";
        }

        private void csecond1_CheckedChanged(object sender, EventArgs e)
        {
            csecond1.Text = csecond1.Checked ? "X" : " ";
        }

        private void csecond3_CheckedChanged(object sender, EventArgs e)
        {
            csecond3.Text = csecond3.Checked ? "X" : " ";
        }

        private void csecond4_CheckedChanged(object sender, EventArgs e)
        {
            csecond4.Text = csecond4.Checked ? "X" : " ";
        }

        private void csecond5_CheckedChanged(object sender, EventArgs e)
        {
            csecond5.Text = csecond5.Checked ? "X" : " ";
        }

        private void csecond6_CheckedChanged(object sender, EventArgs e)
        {
            csecond6.Text = csecond6.Checked ? "X" : " ";
        }

        private void csecond7_CheckedChanged(object sender, EventArgs e)
        {
            csecond7.Text = csecond7.Checked ? "X" : " ";
        }

        private void csecond8_CheckedChanged(object sender, EventArgs e)
        {
            csecond8.Text = csecond8.Checked ? "X" : " ";
        }

        private void csecond9_CheckedChanged(object sender, EventArgs e)
        {
            csecond9.Text = csecond9.Checked ? "X" : " ";
        }

        private void csecond10_CheckedChanged(object sender, EventArgs e)
        {
            csecond10.Text = csecond10.Checked ? "X" : " ";
        }

        private void csecond11_CheckedChanged(object sender, EventArgs e)
        {
            csecond11.Text = csecond11.Checked ? "X" : " ";
        }

        private void cthird2_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird2.Text = cthird2.Checked ? "X" : " ";
        }

        private void cthird3_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird3.Text = cthird3.Checked ? "X" : " ";
        }

        private void cthird5_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird5.Text = cthird5.Checked ? "X" : " ";
        }

        private void cthird6_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird6.Text = cthird6.Checked ? "X" : " ";
        }

        private void cthird7_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird7.Text = cthird7.Checked ? "X" : " ";
        }

        private void cthird8_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird8.Text = cthird8.Checked ? "X" : " ";
        }

        private void cthird9_CheckedChanged_1(object sender, EventArgs e)
        {
            cthird9.Text = cthird9.Checked ? "X" : " ";
        }

        private void cthird10_CheckedChanged(object sender, EventArgs e)
        {
            cthird10.Text = cthird10.Checked ? "X" : " ";
        }

        private void cthird11_CheckedChanged(object sender, EventArgs e)
        {
            cthird11.Text = cthird11.Checked ? "X" : " ";
        }

        private void cthird12_CheckedChanged(object sender, EventArgs e)
        {
            cthird12.Text = cthird12.Checked ? "X" : " ";
        }

        private void cthird13_CheckedChanged(object sender, EventArgs e)
        {
            cthird13.Text = cthird13.Checked ? "X" : " ";
        }

        private void cforth1_CheckedChanged(object sender, EventArgs e)
        {
            cforth1.Text = cforth1.Checked ? "X" : " ";
        }

        private void cforth2_CheckedChanged(object sender, EventArgs e)
        {
            cforth2.Text = cforth2.Checked ? "X" : " ";
        }

        private void cforth3_CheckedChanged(object sender, EventArgs e)
        {
            cforth3.Text = cforth3.Checked ? "X" : " ";
        }

        private void cforth5_CheckedChanged(object sender, EventArgs e)
        {
            cforth5.Text = cforth5.Checked ? "X" : " ";
        }

        private void cforth6_CheckedChanged(object sender, EventArgs e)
        {
            cforth6.Text = cforth6.Checked ? "X" : " ";
        }

        private void cforth7_CheckedChanged(object sender, EventArgs e)
        {
            cforth7.Text = cforth7.Checked ? "X" : " ";
        }

        private void cforth8_CheckedChanged(object sender, EventArgs e)
        {
            cforth8.Text = cforth8.Checked ? "X" : " ";
        }

        private void cforth9_CheckedChanged(object sender, EventArgs e)
        {
            cforth9.Text = cforth9.Checked ? "X" : " ";
        }

        private void cforth10_CheckedChanged(object sender, EventArgs e)
        {
            cforth10.Text = cforth10.Checked ? "X" : " ";
        }

        private void cforth11_CheckedChanged(object sender, EventArgs e)
        {
            cforth11.Text = cforth11.Checked ? "X" : " ";
        }

        private void cforth12_CheckedChanged(object sender, EventArgs e)
        {
            cforth12.Text = cforth12.Checked ? "X" : " ";
        }

        private void cforth13_CheckedChanged(object sender, EventArgs e)
        {
            cforth13.Text = cforth13.Checked ? "X" : " ";
        }

        private void cfifth1_CheckedChanged(object sender, EventArgs e)
        {
            cfifth1.Text = cfifth1.Checked ? "X" : " ";
        }

        private void cfifth2_CheckedChanged(object sender, EventArgs e)
        {
            cfifth2.Text = cfifth2.Checked ? "X" : " ";
        }

        private void cfifth3_CheckedChanged(object sender, EventArgs e)
        {
            cfifth3.Text = cfifth3.Checked ? "X" : " ";
        }

        private void cfifth4_CheckedChanged(object sender, EventArgs e)
        {
            cfifth4.Text = cfifth4.Checked ? "X" : " ";
        }

        private void cfifth6_CheckedChanged(object sender, EventArgs e)
        {
            cfifth6.Text = cfifth6.Checked ? "X" : " ";
        }

        private void cfifth7_CheckedChanged(object sender, EventArgs e)
        {
            cfifth7.Text = cfifth7.Checked ? "X" : " ";
        }

        private void cfifth8_CheckedChanged(object sender, EventArgs e)
        {
            cfifth8.Text = cfifth8.Checked ? "X" : " ";
        }

        private void cfifth9_CheckedChanged(object sender, EventArgs e)
        {
            cfifth9.Text = cfifth9.Checked ? "X" : " ";
        }

        private void cfifth10_CheckedChanged(object sender, EventArgs e)
        {
            cfifth10.Text = cfifth10.Checked ? "X" : " ";
        }

        private void cfifth11_CheckedChanged(object sender, EventArgs e)
        {
            cfifth11.Text = cfifth11.Checked ? "X" : " ";
        }

        private void cfifth12_CheckedChanged(object sender, EventArgs e)
        {
            cfifth12.Text = cfifth12.Checked ? "X" : " ";
        }

        private void cfifth13_CheckedChanged(object sender, EventArgs e)
        {
            cfifth13.Text = cfifth13.Checked ? "X" : " ";
        }

        private void totallo7_TextChanged(object sender, EventArgs e)
        {

        }

        private void w9_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtlo63_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtApplication_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {


           int a = int.Parse(txtTheory.Text);

            List<Class1> c1 = new List<Class1>();

            XmlSerializer serial = new XmlSerializer(typeof(List<Class1>));
            c1.Add(new Class1() { coursename = txtCourseName.Text, theory = a });

            using (FileStream fs = new FileStream(Environment.CurrentDirectory + "\\dnm.xml", FileMode.Create, FileAccess.Write))

            {

                serial.Serialize(fs, c1);
                MessageBox.Show("Created");





            }


        }
    }

}
