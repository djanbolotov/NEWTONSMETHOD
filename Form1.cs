using System;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using TextBox = System.Windows.Forms.TextBox;
using MessageBox = System.Windows.Forms.MessageBox;
using aziretParser;
using System.Diagnostics;
using System.Drawing;

namespace POCKETSEARCHMETHOD
{
    public partial class Form1 : Form
    {
        private const string nameOfExcel = @"\Zhanbolot_uulu_Askabek_LookingForOnePoint.xlsm";
        string inputFuncFX = "";
        decimal df0, df1, ddf0, ddf1, dp, dp0, relerror;
        decimal x0 = 0;
        decimal x1 = 0;
        decimal f0;
        decimal f1;
        decimal e_tol = 0;
        decimal delta = 0;
        int k_max = 0, cond, QND;
        decimal t_max = 0;
        decimal parameterR = 0;
        decimal fplusTol;
        decimal fminusTol;
        Application xls;
        Workbook book = null;
        Worksheet sheet = null;
        public Form1()
        {
            InitializeComponent();
            xls = new Application();
        }

        public int getSign(decimal number)
        {
            if (number < 0)
            {
                return -1;
            }
            else
            {
                return 1;
            }
        }

        public void OpenExcel()
        {
            if (!checkFunction(1)) return;
            string function;
            decimal startPoint;

            try
            {
                if (book == null)
                {
                    book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                }
                if (sheet == null)
                {
                    sheet = book.Sheets["Russian"];
                    sheet.Activate();
                }

                xls.Visible = true;
                function = Function.Text;
                if (InitialApproximation.Text != "" && InitialApproximation.Text != "-" && InitialApproximation.Text != "+" && InitialApproximation.Text != ".")
                {
                    startPoint = Decimal.Parse(InitialApproximation.Text);
                }
                else
                {
                    startPoint = 1;
                }

                sheet.Cells[4, 9] = startPoint;
                sheet.Cells[2, 1] = "f(x)=" + Function.Text;

                StringBuilder builder = new StringBuilder(function);
                builder.Replace("exp", ":");
                builder.Replace("x", "D4");
                builder.Replace(":", "exp");
                function = builder.ToString();
                sheet.Range["E4:E10003"].Value = "=" + function;
            }
            catch
            {
                book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                sheet = book.Sheets["Russian"];
                sheet.Activate();
                xls.Visible = true;
                function = Function.Text;
                if (InitialApproximation.Text != "" && InitialApproximation.Text != "-" && InitialApproximation.Text != "+" && InitialApproximation.Text != ".")
                {
                    startPoint = Decimal.Parse(InitialApproximation.Text);
                }
                else
                {
                    startPoint = 1;
                }

                sheet.Cells[4, 9] = startPoint;
                sheet.Cells[2, 1] = "f(x)=" + Function.Text;

                StringBuilder builder = new StringBuilder(function);
                builder.Replace("exp", ":");
                builder.Replace("x", "D4");
                builder.Replace(":", "exp");
                function = builder.ToString();
                sheet.Range["E4:E10003"].Value = "=" + function;
            }
        }

        private bool parseTry(TextBox t, String type)
        {
            try
            {
                if (type == "Decimal")
                    Decimal.Parse(t.Text, System.Globalization.NumberStyles.Float);
                else if (type == "Integer")
                    int.Parse(t.Text);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void Clean(Control control)
        {
            foreach (var  element in control.Controls)
            {
                switch (element.GetType().Name)
                {
                    case "TextBox":
                        ((TextBox)element).Text = String.Empty;
                        break;
                    case "RadioButton":
                        ((RadioButton)element).Checked = false;
                        break;
                    case "RichTextBox":
                        ((RichTextBox)element).Text = String.Empty;
                        break;
                    case "GroupBox":
                        Clean((Control)element);
                        break;
                    default:
                        break;
                }
            }
        }


        private bool IsOKForDecimalTextBox(char theCharacter, TextBox theTextBox, bool positive)
        {
            if (!char.IsControl(theCharacter) && !char.IsDigit(theCharacter) && (theCharacter != ',') && (theCharacter != '.')
                && (theCharacter != '-') && (theCharacter != '+') && (theCharacter != 'E') && (theCharacter != 'e'))
            {
                return false;
            }
            if(positive && theCharacter == '-' && (theTextBox.Text.IndexOf('E') == -1 && theTextBox.Text.IndexOf('e') == -1))
            {
                return false;
            }
            if (theCharacter == ',' && (theTextBox.Text.IndexOf(',') > -1 || theTextBox.Text.IndexOf('.') > -1))
            {
                return false;
            }
            if (theCharacter == '.' && (theTextBox.Text.IndexOf('.') > -1 || theTextBox.Text.IndexOf(',') > -1))
            {
                return false;
            }
            if (theCharacter == 'e' && (theTextBox.Text.IndexOf('e') > -1 || theTextBox.Text.IndexOf('E') > -1))
            {
                return false;
            }
            if (theCharacter == 'E' && (theTextBox.Text.IndexOf('E') > -1 || theTextBox.Text.IndexOf('e') > -1))
            {
                return false;
            }
            if (theCharacter == '-' && (theTextBox.Text.IndexOf('-') > -1 || theTextBox.Text.IndexOf('+') > -1))
            {
                return false;
            }
            if (theCharacter == '+' && (theTextBox.Text.IndexOf('+') > -1 || theTextBox.Text.IndexOf('-') > -1))
            {
                return false;
            }
            if (((theCharacter == '-') || (theCharacter == '+')) && (theTextBox.SelectionStart != 0 && (theTextBox.Text.IndexOf('E') == -1 && theTextBox.Text.IndexOf('e') == -1)))
            {
                return false;
            }
            if ((char.IsDigit(theCharacter) || (theCharacter == ',') || (theCharacter == '.')) && ((theTextBox.Text.IndexOf('-') > -1) 
                || (theTextBox.Text.IndexOf('+') > -1)) && theTextBox.SelectionStart == 0)
            {
                return false;
            }
            return true;
        }

        public decimal Fx(decimal x)
        {
            decimal result;
            result = aziretParser.Computer.Compute(inputFuncFX, x);
            return result;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenExcel();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Clean(this);
            progressBar1.Visible = false;
        }

        private void InitialApproximation_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, InitialApproximation, false);
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void Tolerance_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, Tolerance, true);
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void SearchStep_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, SearchStep, true);
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void ParametrR_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == (int)48 && ParametrR.Text == "")
            {
                e.Handled = true;
                return;
            }
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private String checkParse()
        {
            String errorMessage = "";
            if (!parseTry(InitialApproximation, "Decimal"))
            {
                errorMessage += "Invalid value of the field x0 (the starting point of the approximation)! Change the input and perform the calculation!\n\n";
            }
            else
            {
                x0 = Decimal.Parse(InitialApproximation.Text, System.Globalization.NumberStyles.Float);
            }

            if (!parseTry(SearchStep, "Decimal"))
            {
                errorMessage += "Invalid value of the field search step! Change the input and perform the calculation!\n\n";
            }
            else
            {
                delta = Decimal.Parse(SearchStep.Text, System.Globalization.NumberStyles.Float);
            }

            if (parseTry(Tolerance, "Decimal"))
            {
                e_tol = Decimal.Parse(Tolerance.Text, System.Globalization.NumberStyles.Float);
            }
            else
            {
                errorMessage += "Invalid value of the Tolerance(e) field (entered tolerance)! Change the input and perform the calculation!\n\n";
            }

            if (!parseTry(LimitOfIterations, "Integer"))
            {
                errorMessage += "Invalid value of the field limit of iterations! Change the input and perform the calculation!\n\n";
            }
            else
            {
                k_max = Int32.Parse(LimitOfIterations.Text);
            }

            if (!parseTry(ParametrR, "Integer"))
            {
                errorMessage += "Invalid value of the field parameter R! Change the input and perform the calculation!\n\n";
            }
            else
            {
                parameterR = Decimal.Parse(ParametrR.Text);   
            }

            if (!parseTry(LimitOfTime, "Decimal"))
            {
                errorMessage += "Invalid value of the field limit of time! Change the input and perform the calculation!\n\n";
            }
            else
            {
                t_max = Decimal.Parse(LimitOfTime.Text, System.Globalization.NumberStyles.Float);
            }

            return errorMessage;
        }

        public bool fullCheck()
        {
            bool check = false;
            if (Function.Text == "" || InitialApproximation.Text == "" ||
                Tolerance.Text == "" || LimitOfIterations.Text == "" ||
                LimitOfTime.Text == "" || SearchStep.Text == "" || ParametrR.Text == "")
            {
                MessageBox.Show("All fields must be filled in! Enter the missing information and make the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (checkParse() != "")
                {
                    MessageBox.Show(checkParse(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (isRigth() && checkFunction(x0))
                    {
                        check = true;
                    }
                }
            }
            return check;
        }

        public string getComparisonSign(decimal a, decimal b)
        {
            if (a > b)
            {
                return ">";
            }
            else if (a < b)
            {
                return "<";
            }
            else
            {
                return "=";
            }
        }

        private bool isRigth()
        {
            bool b = true;
            if (e_tol <= 0)
            {
                MessageBox.Show("The value of the tolerance field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                b = false;
            }
            if (delta <= 0)
            {
                MessageBox.Show("The value of the search step field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                b = false;
            }
            if (k_max <= 0)
            {
                MessageBox.Show("The value of the limit of iterations field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                b = false;
            }
            if (t_max <= 0)
            {
                MessageBox.Show("The value of the limit of time field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                b = false;
            }
            if (parameterR <= 1)
            {
                MessageBox.Show("The value of the parameter R field must be greater than 1! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                b = false;
            }
            if (b)
            {
                return true;
            }
            return false;
        }

        private bool checkFunction(decimal x0)
        {
            inputFuncFX = Function.Text;

            if (inputFuncFX == "" || inputFuncFX.IndexOf('x') == -1)
            {
                MessageBox.Show("The function is entered incorrectly! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                return false;
            }
            try
            {
                if (inputFuncFX.Contains("log") && x0 <= 0 || inputFuncFX.Contains("ln") && x0 <= 0)
                {
                    MessageBox.Show("If you entered function with 'log' or 'ln' value of X0 must greater than zero!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    decimal F1 = Fx(x0);
                    return true;
                }
            }
            catch
            {
                MessageBox.Show("The function or initial approximation is entered incorrectly! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                return false;
            }
        }

        public void FillResult(string solution, string iterations, string resultTolerance, string fminustol, string fplustol, string fxvalue, string fminusplus, string fminusminus, string searchStep)
        {
            ResultX.Text = solution;
            countofiterations.Text = iterations;
            fxplustolerance.Text = fplustol;
            fxminustolerance.Text = fminustol;
            fxminusplustolerance.Text = fminusplus;
            fxminusminustolerance.Text = fminusminus;
            fx.Text = fxvalue;
        }

        public string getError(TextBox tol, decimal error)
        {
            Console.WriteLine(tol);
            if (tol.Text.Contains("E"))
            {
                return error.ToString("0E0");
            }
            else if (tol.Text.Contains("e"))
            {
                return error.ToString("0e0");
            }
            else
            {
                return error.ToString();
            }
        }

        private void LimitOfTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, LimitOfTime, true);
            if(e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void LimitOfIterations_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == (int)48 && LimitOfIterations.Text == "")
            {
                e.Handled = true;
                return;
            }
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("1) Choose a function or write your's on field 'Function'\n" +
                     "2) Click on the button 'Show function graph'\n" +
                     "3) In the opened file select the values for a\n" +
                     "then save the document and return to the program\n" +
                     "4) If you need 'a' value to insert,\n" +
                     "click the button 'Set 'a' like 'X0'' or write your's\n" +
                     "5) Enter tolerance\n" +
                     "6) Enter delta\n" +
                     "7) Enter parameter R for constriction area size\n" +
                     "8) Enter limit of time in sec\n" +
                     "9) Enter limit of iterations \n" +
                     "10) Select search parameter\n" +
                     "Then click the button 'Run Method'.", "Information",
                     MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (book == null)
                {
                    book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                }
                if (sheet == null)
                {
                    sheet = book.Sheets["Russian"];
                    sheet.Activate();
                }
                book.Save();
                InitialApproximation.Text = sheet.Cells[4, 9].Value.ToString();
            }
            catch
            {
                book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                sheet = book.Sheets["Russian"];
                sheet.Activate();
                book.Save();
                InitialApproximation.Text = sheet.Cells[4, 9].Value.ToString();
            }
            xls.Visible = false;
            book = null;
            sheet = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            delta = 0;
            x0 = 0;
            x1 = 0;
            f0 = 0;
            f1 = 0;
            e_tol = 0;
            k_max = 0;
            t_max = 0;
            parameterR = 0;
            inputFuncFX = "";
            fminusTol = 0;
            fplusTol = 0;

            string extremium;

            try
            {
                if (fullCheck())
                {
                    xls.Visible = false;
                    book = null;
                    sheet = null;

                    progressBar1.Value = 0;
                    Clean(groupBox2);
                    validation.Text = String.Empty;
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();

                    //
                    cond = 0;
                    f0 = Fx(x0);
                    string firstDerivative = Derivative.ReturnDerivative(inputFuncFX);
                    string secondDerivative = Derivative.ReturnDerivative(firstDerivative);
                    df0 = Computer.Compute(firstDerivative, x0);
                    ddf0 = Computer.Compute(secondDerivative, x0);
                    dp = df0 / ddf0;
                    int k = 0;

                    do
                    {
                        k = k + 1;

                        progressBar1.Visible = true;
                        progressBar1.Maximum = (int)(k + 0.00000001);
                        progressBar1.Value = k;

                        if (k > k_max)
                        {
                            stopwatch.Stop();
                            f1 = Fx(x1);
                            fminusTol = Fx(x1 - e_tol);
                            fplusTol = Fx(x1 + e_tol);
                            DialogResult result = MessageBox.Show("Iteration limit reached. Do you want to add iterations?",
                                "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (result == DialogResult.Yes)
                            {
                                k_max += k_max;
                                LimitOfIterations.Text = k_max.ToString();
                            }
                            else
                            {
                                k--;
                                validation.Text += "Result X* not found because of limit of iterations = " + k_max + "." +
                                    "\nSince the following condition is false, namely:" +
                                    "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                    "\nResult X* is not extremum of the function.";
                                validation.ForeColor = Color.Red;

                                FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"), getError(absError, Math.Abs(delta)));
                                absError.Text = getError(Tolerance, Math.Abs(x1 - x0));

                                DialogResult answer = MessageBox.Show("Result X* not found because of maximum limit of iterations = " + k_max + "." +
                                "\nSince the following condition is false, namely:" +
                                "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                "\nResult X* is not extremum of the function." +
                                "\n\nYou probably entered the value of 'a' incorrectly on Ecxel!" +
                                "\nYou need to open the graph and select the correct points [a;b]!" +
                                "\n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                                if (answer == DialogResult.Yes)
                                {
                                    OpenExcel();
                                }
                                break;
                            }
                            stopwatch.Start();
                        }

                        if (stopwatch.ElapsedMilliseconds >= t_max * 1000)
                        {
                            stopwatch.Stop();
                            f1 = Fx(x1);
                            fminusTol = Fx(x1 - e_tol);
                            fplusTol = Fx(x1 + e_tol);
                            DialogResult result = MessageBox.Show("Time limit reached. Do you want to add time?",
                                "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (result == DialogResult.Yes)
                            {
                                t_max += t_max;
                                LimitOfTime.Text = t_max.ToString();
                            }
                            else
                            {
                                validation.Text += "Result X* not found because of limit of time = " + t_max + " sec." +
                                    "\nSince the following condition is false, namely:" +
                                    "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                    "\nResult X* is not extremum of the function.";
                                validation.ForeColor = Color.Red;

                                FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"), getError(absError, Math.Abs(delta)));
                                absError.Text = getError(Tolerance, Math.Abs(x1 - x0));

                                DialogResult answer = MessageBox.Show("Result X* not found because of maximum time limit = " + t_max + " sec." +
                                "\nSince the following condition is false, namely:" +
                                "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                "\nResult X* is not extremum of the function." +
                                "\n\nYou probably entered the value of 'a' incorrectly on Ecxel!" +
                                "\nYou need to open the graph and select the correct points [a;b]!" +
                                "\n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                                if (answer == DialogResult.Yes)
                                {
                                    OpenExcel();
                                }
                                break;
                            }
                            stopwatch.Start();
                        }

                        if (Math.Abs(ddf0) <= e_tol)
                        {
                            cond = 1;
                            QND = 0;
                        }
                        else
                        {
                            dp = df0 / ddf0;
                        }
                        if (k == 1)
                        {
                            dp0 = dp;
                        }

                        if (Math.Sign(dp0) == Math.Sign(dp))
                        {
                            x1 = x0 - dp;
                        }
                        else
                        {
                            x1 = x0 - dp / parameterR;
                        }
                        f1 = Fx(x1);
                        firstDerivative = Derivative.ReturnDerivative(inputFuncFX);
                        secondDerivative = Derivative.ReturnDerivative(firstDerivative);
                        df1 = Computer.Compute(firstDerivative, x1);
                        ddf1 = Computer.Compute(secondDerivative, x1);
                        dp0 = dp = df1 / ddf1;
                        relerror = Convert.ToDecimal(2) * Math.Abs(dp) / (Math.Abs(x1) + e_tol);
                        if (relerror < delta)
                        {
                            if (cond != 1)
                            {
                                cond = 2;
                            }
                        }
                        x0 = x1;
                        df0 = df1;
                        ddf0 = ddf1;

                        fminusTol = Fx(x1 - e_tol);
                        fplusTol = Fx(x1 + e_tol);
                        decimal controlFminusTol = Fx(x1 - Convert.ToDecimal(0.0001));
                        decimal controlFplusTol = Fx(x1 - Convert.ToDecimal(0.0001));

                        if (cond != 0)
                        {
                                if (f1 < fminusTol && f1 < fplusTol)
                                {
                                    FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"), getError(absError, Math.Abs(relerror)));

                                    validation.Text += "Since the following condition is true, namely:" +
                                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                            "\nResult X* is minimizer of the function. It has been found with the error = " +  relerror + ". This is less than or equal to given Tolerance!";
                                    validation.ForeColor = Color.Green;
                                absError.Text = Convert.ToString(relerror);
                                break;
                                }
                                if ((f1 >= fminusTol && f1 >= fplusTol))
                                {
                                FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"), getError(absError, Math.Abs(relerror)));

                                if (f1 >= controlFminusTol && f1 >= controlFplusTol)
                                {
                                    validation.Text += "Since the following condition is true, namely:" +
                                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                            "\nResult X* is maximizer of the function. It has been found with the error = " + relerror + ". This is less than or equal to given Tolerance!";
                                    validation.ForeColor = Color.Green;
                                }
                                else
                                {
                                    validation.Text += "Since the following condition is false, namely: " + "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" + "Cannot be proven to be the minimum or maximum. Because " + "\nSign(f(X*)-f(X*+1e-10)) = " + getSign(f1 - controlFplusTol) + " and Sign(f(X*)-f(X*-1e-10)) = " + getSign(f1 - controlFminusTol) + "!";
                                    validation.ForeColor = Color.Red;
                                }

                                absError.Text = Convert.ToString(relerror);
                                break;
                                }
                            FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"), getError(absError, Math.Abs(relerror)));

                            validation.Text += "Since the following condition is false, namely:" +
                                    "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                    "\nResult X* is not maximizer or maximizer of the function. It's inflection point!";
                            validation.ForeColor = Color.Red;
                            absError.Text = Convert.ToString(relerror);
                            break;
                        }
                    } while (true);

                    stopwatch.Stop();
                    elapsedtime.Text = stopwatch.ElapsedMilliseconds / 1000.0 + " sec";
                    timer1.Enabled = true;
                    timer1.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                progressBar1.Value = 0;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            timer1.Enabled = false;
            timer1.Stop();
        }

        private void Function_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void InitialApproximation_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Tolerance_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void SearchStep_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void ParametrR_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void LimitOfTime_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void LimitOfIterations_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Maximum_CheckedChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Minimum_CheckedChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            xls.Quit();
        }
    }
}
